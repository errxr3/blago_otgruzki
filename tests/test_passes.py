"""Пункт 7: отгрузка и пропуска на обе Газели — только в первом кабинете; КГТ — без пропуска."""

import pytest
import requests

import blago_otgruzki as b
from fakes import CABINETS

CFG = {"driver": "Алексей Солохов", "phone": "+7 900 000-00-00", "cars": "а139уо69 газель\nв678ха69 газель"}
FBS = {"warehouse_name": "FBS Тверь", "warehouse_id": 501, "delivery_method_id": 11, "carriages": []}
KGT = {"warehouse_name": "FBS КГТ Тверь", "warehouse_id": 502, "delivery_method_id": 12, "carriages": []}
FORM = {"Z": {"fbs тверь", "fbs кгт тверь"}, "R": {"fbs - сгт тверь", "fbs тверь"}, "IZ": {"fbs тверь"}}


@pytest.fixture
def carriages(ozon):
    ozon.carriages = {"Z": [dict(FBS), dict(KGT)],
                      "R": [dict(FBS, delivery_method_id=21), dict(KGT, warehouse_name="FBS - СГТ Тверь")],
                      "IZ": [dict(FBS, delivery_method_id=31)]}
    return ozon


def bodies(ozon, path):
    return [body for _, p, body in ozon.calls if p == path]


def test_only_first_cabinet_gets_carriage_and_pass(carriages, log):
    ozon = carriages
    made = b.ozon_passes(CABINETS, CFG, log, form=FORM)
    assert made == 1
    assert {cab for cab, _, _ in ozon.calls} == {"Z"}          # во 2-й и 3-й кабинет — ни одного запроса
    assert [(cab, mid) for cab, mid, _ in ozon.created] == [("Z", 11), ("Z", 12)]
    passes = bodies(ozon, "/v1/carriage/pass/create")
    assert len(passes) == 1                                     # на КГТ пропуска нет
    assert passes[0]["carriage_id"] == 7000
    assert [p["vehicle_license_plate"] for p in passes[0]["arrival_passes"]] == ["А139УО69", "В678ХА69"]
    assert {p["driver_name"] for p in passes[0]["arrival_passes"]} == {"Солохов Алексей"}
    text = log.text()
    assert "FBS КГТ Тверь — КГТ, пропуск не нужен" in text
    assert "✔ Отгрузка сформирована Z · FBS Тверь · № 7000" in text


def test_user_says_no_nothing_is_formed(carriages, log):
    made = b.ozon_passes(CABINETS, CFG, log, form=FORM, confirm=lambda text: False)
    assert made == 0 and carriages.created == []
    assert not bodies(carriages, "/v1/carriage/pass/create")


def test_confirm_text_says_kgt_has_no_pass(carriages, log):
    asked = []
    b.ozon_passes(CABINETS, CFG, log, form=FORM, confirm=lambda text: asked.append(text) or False)
    assert "Z · FBS Тверь" in asked[0] and "Z · FBS КГТ Тверь" in asked[0]
    assert "R ·" not in asked[0] and "IZ ·" not in asked[0]
    assert "На КГТ — только отгрузка, без пропуска" in asked[0]


def test_passes_only_for_existing_carriages_without_form(carriages, log):
    """Кнопка «Пропуска»: отгрузок не создаёт, пропуск — на сформированные без пропуска, КГТ мимо."""
    ozon = carriages
    ozon.carriages["Z"][0]["carriages"] = [{"id": 555, "status": "formed"}, {"id": 556, "status": "closed"}]
    ozon.carriages["Z"][1]["carriages"] = [{"id": 777, "status": "formed"}]
    ozon.carriages["R"][0]["carriages"] = [{"id": 888, "status": "formed"}]
    made = b.ozon_passes(CABINETS, CFG, log)
    assert made == 1 and ozon.created == []
    assert [x["carriage_id"] for x in bodies(ozon, "/v1/carriage/pass/create")] == [555]
    assert {cab for cab, _, _ in ozon.calls} == {"Z"}


def test_approve_with_containers_is_repeated_with_one(carriages, log):
    ozon = carriages
    ozon.fail[("Z", "/v1/carriage/approve")] = [(400, {"message": "INCORRECT_CONTAINERS_COUNT"})]
    b.ozon_passes(CABINETS, CFG, log, form={"Z": {"fbs тверь"}})
    approves = bodies(ozon, "/v1/carriage/approve")
    assert approves == [{"carriage_id": 7000}, {"carriage_id": 7000, "containers_count": 1}]
    assert "✔ Отгрузка сформирована" in log.text()


def test_carriage_create_error_is_reported_and_no_pass(carriages, log):
    carriages.fail[("Z", "/v1/carriage/create")] = (500, {"message": "boom"})
    assert b.ozon_passes(CABINETS, CFG, log, form={"Z": {"fbs тверь"}}) == 0
    assert "не создана" in log.text()
    assert not bodies(carriages, "/v1/carriage/pass/create")


def test_carriage_create_timeout_is_not_repeated(carriages, log):
    carriages.fail[("Z", "/v1/carriage/create")] = requests.Timeout("timeout")
    b.ozon_passes(CABINETS, CFG, log, form={"Z": {"fbs тверь"}})
    assert carriages.count("/v1/carriage/create") == 1
    assert "проверь в кабинете, создалась ли" in log.text()


def test_no_driver_no_requests(carriages, log):
    assert b.ozon_passes(CABINETS, {"driver": "", "cars": ""}, log, form=FORM) == 0
    assert carriages.calls == []


def test_car_and_driver_helpers():
    assert b.parse_cars("a139yo69 газель; в678ха69") == [("А139УО69", "Газель"), ("В678ХА69", "Газель")]
    assert b.driver_full("Алексей Солохов") == "Солохов Алексей"
    assert b.is_kgt("FBS - CГТ Тверь") and b.is_kgt("fbs кгт") and not b.is_kgt("FBS Тверь")
