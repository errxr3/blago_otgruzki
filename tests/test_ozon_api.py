"""Запросы к Ozon: повторы при 429 и обрыве, ошибки 500, страницы, задание на этикетки."""

import json
import re

import pytest
import requests
import responses

import blago_otgruzki as b
from fakes import label_pdf

H = {"Client-Id": "1001", "Api-Key": "test", "Content-Type": "application/json"}
URL = f"{b.API}/v4/posting/fbs/list"


def test_429_forever_gives_up_and_returns_429(isolated):
    isolated.add(responses.POST, URL, status=429, json={"message": "limit"})
    r = b.ozon_post(URL, headers=H, json={}, timeout=5)
    assert r.status_code == 429
    assert len(isolated.calls) == len(b.RETRY_WAITS) + 1
    assert "ошибка 429" in b.api_error(r)


def test_429_then_ok(isolated):
    isolated.add(responses.POST, URL, status=429, json={})
    isolated.add(responses.POST, URL, status=200, json={"ok": 1})
    assert b.ozon_post(URL, headers=H, json={}, timeout=5).json() == {"ok": 1}


def test_network_error_is_retried_for_reads(isolated):
    isolated.add(responses.POST, URL, body=requests.ConnectionError("reset"))
    with pytest.raises(requests.ConnectionError):
        b.ozon_post(URL, headers=H, json={}, timeout=5)
    assert len(isolated.calls) == len(b.NET_WAITS) + 1


def test_network_error_is_not_retried_for_changes(isolated):
    """Сборка и отгрузка идут с net_retry=False: при обрыве повторять опасно."""
    isolated.add(responses.POST, URL, body=requests.Timeout("timeout"))
    with pytest.raises(requests.Timeout):
        b.ozon_post(URL, net_retry=False, headers=H, json={}, timeout=5)
    assert len(isolated.calls) == 1


def test_500_is_returned_with_readable_error(isolated):
    isolated.add(responses.POST, URL, status=500, json={"message": "internal error"})
    with pytest.raises(RuntimeError, match="ошибка 500.*internal error"):
        b.list_postings(H, 14)


def test_wrong_key_has_hint(isolated):
    isolated.add(responses.POST, URL, status=403, json={"message": "forbidden"})
    with pytest.raises(RuntimeError, match="ключ неверный"):
        b.check_key(H)


def test_disabled_method_is_named(isolated):
    isolated.add(responses.POST, URL, status=410, json={"message": "method is obsolete"})
    with pytest.raises(RuntimeError, match="OZON ОТКЛЮЧИЛ МЕТОД /v4/posting/fbs/list"):
        b.list_postings(H, 14)


def test_list_reads_all_pages(isolated):
    def page(request):
        cursor = json.loads(request.body)["cursor"]
        n = int(cursor or 0)
        items = [{"posting_number": f"{n}-{i}", "status": b.READY, "products": []} for i in range(100)]
        return 200, {}, json.dumps({"postings": items, "has_next": n < 2, "cursor": str(n + 1)})

    isolated.add_callback(responses.POST, URL, callback=page)
    assert len(b.list_postings(H, 14)) == 300
    assert len(isolated.calls) == 3


def test_list_works_when_ozon_rejects_barcodes(isolated):
    def page(request):
        if "with" in json.loads(request.body):
            return 400, {}, json.dumps({"message": "unknown field"})
        return 200, {}, json.dumps({"postings": [{"posting_number": "1-1", "status": b.READY}], "has_next": False})

    isolated.add_callback(responses.POST, URL, callback=page)
    assert list(b.list_postings(H, 14)) == ["1-1"]


def test_units_of_and_plural():
    assert b.units_of({"products": [{"quantity": 2}, {"quantity": 1}]}) == 3
    assert b.units_of({"products": []}) == 1
    assert [b.plural(n, "заказ", "заказа", "заказов") for n in (1, 2, 5, 11, 21, 112)] == [
        "1 заказ", "2 заказа", "5 заказов", "11 заказов", "21 заказ", "112 заказов"]


def test_ship_splits_every_piece_into_own_package(isolated):
    seen = {}

    def ship(request):
        seen.update(json.loads(request.body))
        return 200, {}, json.dumps({"result": ["5-1", "5-3", "5-5"],
                                    "additional_data": [{"posting_number": "5-3", "products": [{"offer_id": "B"}]}]})

    isolated.add_callback(responses.POST, f"{b.API}/v4/posting/fbs/ship", callback=ship)
    numbers, arts, err = b.ship_posting(H, "5-1", [{"offer_id": "A", "sku": 11, "quantity": 2},
                                                   {"offer_id": "B", "sku": 22, "quantity": 1}])
    assert (numbers, arts, err) == (["5-1", "5-3", "5-5"], {"5-3": "B"}, "")
    assert seen["packages"] == [{"products": [{"product_id": 11, "quantity": 1}]},
                                {"products": [{"product_id": 11, "quantity": 1}]},
                                {"products": [{"product_id": 22, "quantity": 1}]}]


# ---------------------------------------------------------------- этикетки заданием

CREATE = f"{b.API}/v3/posting/fbs/package-label/create"
GET = f"{b.API}/v2/posting/fbs/package-label/get"


def test_label_task_prefers_small_label(isolated):
    isolated.add(responses.POST, CREATE, json={"result": {"tasks": [
        {"task_id": 1, "task_type": "big_label"}, {"task_id": 2, "task_type": "small_label"}]}})
    asked = []

    def get(request):
        asked.append(json.loads(request.body)["task_id"])
        return 200, {}, json.dumps({"result": {"file_url": "https://files.test/x.pdf", "status": {"code": "ok"}}})

    isolated.add_callback(responses.POST, GET, callback=get)
    isolated.add(responses.GET, re.compile(r"https://files\.test/.*"), body=label_pdf(["1-1"]))
    pdf, unprinted, err = b.label_task(H, ["1-1"])
    assert pdf[:4] == b"%PDF" and not unprinted and not err and asked == [2]


def test_label_task_without_tasks(isolated):
    isolated.add(responses.POST, CREATE, json={"result": {"tasks": []}})
    assert b.label_task(H, ["1-1"])[2].startswith("Ozon не дал задание")


def test_label_task_error_from_ozon(isolated):
    isolated.add(responses.POST, CREATE, json={"tasks": [{"task_id": 1, "task_type": "small_label"}]})
    isolated.add(responses.POST, GET, json={"error": {"code": "X", "message": "нет этикетки"}})
    assert b.label_task(H, ["1-1"]) == (None, {}, "Ozon: нет этикетки")


def test_label_task_never_ready(isolated):
    isolated.add(responses.POST, CREATE, json={"tasks": [{"task_id": 1, "task_type": "small_label"}]})
    isolated.add(responses.POST, GET, json={"status": {"code": "pending"}})
    assert "долго готовит" in b.label_task(H, ["1-1"])[2]


def test_grab_labels_network_error_marks_all(isolated):
    isolated.add(responses.POST, CREATE, body=requests.ConnectionError("down"))
    good, bad = b.grab_labels(H, ["1-1", "2-1"])
    assert good == [] and [n for n, _ in bad] == ["1-1", "2-1"] and "нет связи" in bad[0][1]
