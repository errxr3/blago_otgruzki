"""Пункт 8: WB — склады МГТ и СГТ, 4 поставки в день, вчерашняя поставка, стикеры и лист подбора."""

from datetime import date

import openpyxl
import pytest
import requests
from pypdf import PdfReader

import blago_otgruzki as b
from fakes import wb_report

TODAY = date(2026, 11, 10)
MGT_WH, SGT_WH = "FBS Тверь - Эммаус", "Тверь СГТ"
TOKEN = "test-wb-token"


@pytest.fixture
def shop(wb, tmp_path):
    wb.order(501, "Zerkalo_50x130", cargo=1, warehouse_id=1, created="2026-11-10T05:00:00Z")
    wb.order(502, "Ambient 80 krug", cargo=1, warehouse_id=1)
    wb.order(505, "bp_Zerkalo_50x130", cargo=1, warehouse_id=1)
    wb.order(503, "Zerkalo_100x200", cargo=2, warehouse_id=2)
    wb.order(504, "Ambient 200x100", cargo=2, warehouse_id=2)
    wb.order(506, "Ambient 100 krug", cargo=2, warehouse_id=2)   # доставить завтра — сегодня не едет
    folder = tmp_path / "ВБ 10.11"
    folder.mkdir()
    mgt = wb_report(folder / "ВБ 10.11.xlsx", [
        (501, "Zerkalo_50x130", MGT_WH, "-", 1), (502, "Ambient 80 krug", MGT_WH, "-", 1),
        (505, "bp_Zerkalo_50x130", MGT_WH, "-", 1),
    ], warehouse=MGT_WH)
    sgt = wb_report(folder / "СГТ 10.11.xlsx", [
        (503, "Zerkalo_100x200", SGT_WH, "10.11.2026", 1), (504, "Ambient 200x100", SGT_WH, "10.11.2026", 1),
        (506, "Ambient 100 krug", SGT_WH, "11.11.2026", 1),
    ], warehouse=SGT_WH, deliver="10.11.2026")
    return wb, [str(mgt), str(sgt)], tmp_path / "Этикетки ВБ"


def run(shop, log, progress, mode="run", confirm=lambda text: True):
    wb, files, out = shop
    return b.process_wb(files, TOKEN, str(out), log, progress, mode=mode, confirm=confirm, today=TODAY)


def changed(wb):
    return [(m, p) for m, p in wb.methods() if (m, p) == ("POST", "/api/v3/supplies") or m in ("PATCH", "PUT", "DELETE")]


def test_supply_name_and_parse():
    name = b.wb_supply_name(TODAY, "МГТ", "Коминтерна")
    assert name == "Поставка от 10.11.2026 МГТ коминтерна"
    assert b.wb_parse_name(name) == (TODAY, "МГТ", "Коминтерна", False)
    assert b.wb_parse_name("Поставка от 10.11.2026 СГТ учительская юр.лицо") == (TODAY, "СГТ", "Учительская", True)
    assert b.wb_parse_name("Моя поставка") is None and b.wb_parse_name(None) is None


def test_sgt_is_filtered_by_delivery_date(shop):
    keep, desc = b.wb_report_filter(shop[1][1])
    assert "Доставить на склад: 10.11.2026" in desc and "Склад продавца: Тверь СГТ" in desc
    rpt = b.read_report(shop[1][1], keep=keep)
    assert set(rpt.postings) == {"503", "504"}


def test_report_without_filters_takes_everything(shop, tmp_path):
    path = wb_report(tmp_path / "все.xlsx", [(1, "A", MGT_WH, "-", 1), (2, "B", SGT_WH, "-", 1)])
    keep, desc = b.wb_report_filter(path)
    assert keep is None and "берутся все" in desc


def test_check_changes_nothing_in_wb(shop, log, progress):
    wb = shop[0]
    run(shop, log, progress, mode="check")
    text = log.text()
    assert "ВСЕГО ШК БУДЕТ СКАЧАНО: 5" in text
    for name in ("МГТ коминтерна — 2 задания", "МГТ учительская — 1 задание",
                 "СГТ коминтерна — 1 задание", "СГТ учительская — 1 задание"):
        assert f"Поставка от 10.11.2026 {name}" in text
    assert changed(wb) == [] and wb.supplies == []      # ни создания поставок, ни добавления заданий
    assert not shop[2].exists()


def test_run_makes_four_supplies_stickers_and_pick_lists(shop, log, progress):
    wb, _, out = shop
    run(shop, log, progress)
    assert sorted(s["name"] for s in wb.supplies) == [
        "Поставка от 10.11.2026 МГТ коминтерна", "Поставка от 10.11.2026 МГТ учительская",
        "Поставка от 10.11.2026 СГТ коминтерна", "Поставка от 10.11.2026 СГТ учительская"]
    by_name = {s["name"]: s["orders"] for s in wb.supplies}
    assert sorted(by_name["Поставка от 10.11.2026 МГТ коминтерна"]) == ["501", "505"]
    assert by_name["Поставка от 10.11.2026 СГТ учительская"] == ["504"]
    assert "506" in wb.new                                # завтрашнее задание СГТ не тронуто
    kom = out / b.KOMINTERNA
    pdf = PdfReader(str(kom / "ШК Поставка от 10.11.2026 МГТ коминтерна.pdf"))
    assert len(pdf.pages) == 2
    assert all((page.extract_text() or "").strip() == "" for page in pdf.pages)   # стикер без наложения текста
    assert round(float(pdf.pages[0].mediabox.width) / b.MM) == 58
    # лист подбора: порядок как у WB — Zerkalo_50x130, сразу за ним bp_Zerkalo_50x130
    ws = openpyxl.load_workbook(kom / "Лист подбора Поставка от 10.11.2026 МГТ коминтерна.xlsx").active
    rows = [(r[0].value, r[1].value) for r in ws.iter_rows(min_row=6) if r[0].value]
    assert rows == [("Zerkalo_50x130", "501"), ("bp_Zerkalo_50x130", "505")]
    assert (out / b.UCHITELSKAYA / "ШК Поставка от 10.11.2026 СГТ учительская.pdf").exists()
    assert "СКАЧАНО ШК: 5" in log.text()


def test_sticker_order_follows_pick_list():
    arts = ["bp_Zerkalo_160x50", "Zerkalo_60x160", "Ambient 80", "Zerkalo_160x50", "alum rama"]
    assert sorted(arts, key=b.pick_order) == ["alum rama", "Ambient 80", "Zerkalo_160x50", "bp_Zerkalo_160x50",
                                              "Zerkalo_60x160"]


def test_rerun_same_day_creates_no_duplicates(shop, log, progress):
    wb = shop[0]
    run(shop, log, progress)
    creates = wb.methods().count(("POST", "/api/v3/supplies"))
    run(shop, log, progress)
    assert wb.methods().count(("POST", "/api/v3/supplies")) == creates == 4
    assert len(wb.supplies) == 4
    assert "на сборке (уже в поставке)" in log.text()


def test_new_orders_later_same_day_go_into_todays_supply(shop, log, progress, tmp_path):
    wb, _, out = shop
    run(shop, log, progress)
    wb.order(507, "Zerkalo_60x160", cargo=1, warehouse_id=1)
    extra = wb_report(tmp_path / "ВБ 2.xlsx", [(507, "Zerkalo_60x160", MGT_WH, "-", 1)])
    b.process_wb([str(extra)], TOKEN, str(out), log, progress, mode="run", today=TODAY)
    assert len(wb.supplies) == 4                           # новая поставка не создана
    mgt = next(s for s in wb.supplies if s["name"] == "Поставка от 10.11.2026 МГТ коминтерна")
    assert sorted(mgt["orders"]) == ["501", "505", "507"]


def test_final_no_changes_nothing(shop, log, progress):
    wb = shop[0]
    run(shop, log, progress, confirm=lambda text: False)
    assert changed(wb) == [] and "Отменено — в WB ничего не изменено" in log.text()


def test_b2b_orders_go_to_separate_supply(shop, log, progress, tmp_path):
    wb, _, out = shop
    wb.new.clear()
    wb.order(601, "Zerkalo_50x130", cargo=1)
    wb.order(602, "Zerkalo_50x130", cargo=1, options={"isB2B": True})
    rep = wb_report(tmp_path / "b2b.xlsx", [(601, "Zerkalo_50x130", MGT_WH, "-", 1), (602, "Zerkalo_50x130", MGT_WH, "-", 1)])
    b.process_wb([str(rep)], TOKEN, str(out), log, progress, mode="run", today=TODAY)
    assert sorted(s["name"] for s in wb.supplies) == ["Поставка от 10.11.2026 МГТ коминтерна",
                                                       "Поставка от 10.11.2026 МГТ коминтерна юр.лицо"]


def test_red_and_yellow_in_wb_report(wb, tmp_path, log, progress):
    wb.order(701, "Zerkalo_50x130", created="2026-11-10T09:00:00Z")
    wb.order(702, "Zerkalo_50x130", created="2026-11-10T05:00:00Z")   # старше — едет
    wb.order(703, "Ambient 80 krug")
    rep = wb_report(tmp_path / "цвет.xlsx", [(701, "Zerkalo_50x130", MGT_WH, "-", 1), (702, "Zerkalo_50x130", MGT_WH, "-", 1),
                                             (703, "Ambient 80 krug", MGT_WH, "-", 1)],
                    red={"Ambient 80 krug"}, yellow={"Zerkalo_50x130": 1})
    out = tmp_path / "out"
    b.process_wb([str(rep)], TOKEN, str(out), log, progress, mode="run", today=TODAY)
    assert [s["orders"] for s in wb.supplies] == [["702"]]
    skipped = (out / "НЕ_ЕДУТ.txt").read_text(encoding="utf-8")
    assert "703\tкрасный" in skipped and "701" in skipped


def test_unknown_order_is_reported(shop, log, progress, tmp_path):
    wb = shop[0]
    wb.status["999"] = ("cancel", "canceled_by_client")
    rep = wb_report(tmp_path / "x.xlsx", [(999, "Zerkalo_50x130", MGT_WH, "-", 1), (998, "Zerkalo_50x130", MGT_WH, "-", 1)])
    b.process_wb([str(rep)], TOKEN, str(tmp_path / "o"), log, progress, mode="check", today=TODAY)
    text = log.text()
    assert "статус «отменено продавцом»" in text and "не найдено в WB" in text


def test_empty_wb_report(wb, tmp_path, log, progress):
    rep = wb_report(tmp_path / "пусто.xlsx", [])
    b.process_wb([str(rep)], TOKEN, str(tmp_path / "o"), log, progress, mode="run", today=TODAY)
    assert "В отчётах нет заданий, которые едут" in log.text()
    assert changed(wb) == []


def test_broken_sticker_does_not_crash_after_supplies_created(shop, log, progress):
    """WB отдал битую картинку стикера уже после добавления в поставку: не падаем, пишем в НЕ_СКАЧАНО."""
    wb, _, out = shop
    wb.bad_sticker.add("505")
    run(shop, log, progress)
    pdf = PdfReader(str(out / b.KOMINTERNA / "ШК Поставка от 10.11.2026 МГТ коминтерна.pdf"))
    assert len(pdf.pages) == 1
    assert "505" in (out / "НЕ_СКАЧАНО.txt").read_text(encoding="utf-8")
    assert "СКАЧАНО ШК: 4" in log.text()


def test_sticker_error_500_is_reported(shop, log, progress):
    wb, _, out = shop
    wb.fail[("POST", "/api/v3/orders/stickers")] = (500, {"message": "boom"})
    run(shop, log, progress)
    assert "ошибка 500" in (out / "НЕ_СКАЧАНО.txt").read_text(encoding="utf-8")
    assert len(wb.supplies) == 4


def test_429_is_retried(shop, log, progress):
    wb = shop[0]
    wb.fail[("GET", "/api/v3/orders/new")] = [(429, {"message": "limit"}), (429, {"message": "limit"})]
    run(shop, log, progress, mode="check")
    assert "ВСЕГО ШК БУДЕТ СКАЧАНО: 5" in log.text()
    assert wb.methods().count(("GET", "/api/v3/orders/new")) == 3


def test_500_on_orders_gives_readable_error(shop, log, progress):
    shop[0].fail[("GET", "/api/v3/orders/new")] = (500, {"message": "internal"})
    with pytest.raises(RuntimeError, match="ошибка 500"):
        run(shop, log, progress, mode="check")


def test_bad_token_has_hint(shop, log, progress):
    shop[0].fail[("GET", "/api/v3/orders/new")] = (401, {"message": "unauthorized"})
    with pytest.raises(RuntimeError, match="токен неверный или просрочен"):
        run(shop, log, progress, mode="check")


def test_supply_create_error_keeps_other_supplies(shop, log, progress):
    wb, _, out = shop
    wb.fail[("POST", "/api/v3/supplies")] = [(500, {"message": "boom"})]
    run(shop, log, progress)
    assert len(wb.supplies) == 3
    assert "не добавлено" in (out / "НЕ_СКАЧАНО.txt").read_text(encoding="utf-8")


def test_timeout_on_wb_is_a_requests_error(wb, log, progress, tmp_path):
    wb.fail[("GET", "/api/v3/orders/new")] = requests.Timeout("timeout")
    rep = wb_report(tmp_path / "t.xlsx", [(1, "Zerkalo_50x130", MGT_WH, "-", 1)])
    with pytest.raises(requests.RequestException):
        b.process_wb([str(rep)], TOKEN, str(tmp_path / "o"), log, progress, mode="check", today=TODAY)


# ---------------------------------------------------------------- не уехало вчера (1.37)

def leftover(wb):
    """Вчерашняя «МГТ учительская» не уехала: в ней одно задание, его ШК выкинули."""
    wb.history["401"] = {"id": 401, "article": "Ambient warm 80 krug multi", "cargoType": 1, "warehouseId": 1,
                         "createdAt": "2026-11-09T08:00:00Z"}
    wb.status["401"] = ("confirm", "waiting")
    wb.supply("WB-GI-OLD", "Поставка от 09.11.2026 МГТ учительская", orders=[401])


def test_morning_puts_yesterdays_leftover_into_report_and_sborka(wb, tmp_path, log, progress, monkeypatch):
    leftover(wb)
    wb.order(501, "Zerkalo_50x130")
    monkeypatch.setattr(b, "wb_warehouses", lambda token: {1: MGT_WH, 2: SGT_WH})
    out = tmp_path / "Сборка"
    files = b.morning({}, TOKEN, None, None, out, 14, log, progress, today=TODAY)
    ws = openpyxl.load_workbook(files["wb"]).active
    ids = [r[0] for r in ws.iter_rows(min_row=2, max_col=1, values_only=True) if r[0]]
    assert ids == [401, 501]                       # вчерашнее — первым
    assert "Не уехало вчера: «Поставка от 09.11.2026 МГТ учительская» — 1 задание" in log.text()
    assert b.read_totals(files["wb"])["Ambient warm 80 krug multi"] == 1   # попадёт в «Сборку»
    assert changed(wb) == []


def test_leftover_in_report_is_not_failed_and_counted(wb, tmp_path, log, progress):
    leftover(wb)
    wb.order(501, "Zerkalo_50x130")
    rep = wb_report(tmp_path / "ВБ.xlsx", [(401, "Ambient warm 80 krug multi", MGT_WH, "-", 1),
                                           (501, "Zerkalo_50x130", MGT_WH, "-", 1)])
    b.process_wb([str(rep)], TOKEN, str(tmp_path / "o"), log, progress, mode="check", today=TODAY)
    text = log.text()
    assert "Не поедут" not in text
    assert "Поставка от 10.11.2026 МГТ учительская — 0 заданий + не уехало вчера 1" in text
    assert "401  Ambient warm 80 krug multi  →  Поставка от 10.11.2026 МГТ учительская" in text
    assert "ВСЕГО ШК БУДЕТ СКАЧАНО: 2" in text
    assert changed(wb) == []


def test_leftover_moved_to_todays_supply_and_old_deleted(wb, tmp_path, log, progress):
    leftover(wb)
    wb.order(501, "Zerkalo_50x130")
    wb.order(502, "Ambient 80 krug")
    rep = wb_report(tmp_path / "ВБ.xlsx", [(401, "Ambient warm 80 krug multi", MGT_WH, "-", 1),
                                           (501, "Zerkalo_50x130", MGT_WH, "-", 1),
                                           (502, "Ambient 80 krug", MGT_WH, "-", 1)])
    out = tmp_path / "o"
    asked = []
    b.process_wb([str(rep)], TOKEN, str(out), log, progress, mode="run", today=TODAY,
                 confirm=lambda text: asked.append(text) or True)
    assert any("Не уехало вчера: 1 задание" in q and "401" in q for q in asked)
    by_name = {s["name"]: sorted(s["orders"]) for s in wb.supplies}
    assert by_name == {"Поставка от 10.11.2026 МГТ коминтерна": ["501"],
                       "Поставка от 10.11.2026 МГТ учительская": ["401", "502"]}   # вчерашняя удалена
    assert "Пустая «Поставка от 09.11.2026 МГТ учительская» удалена" in log.text()
    pdf = PdfReader(str(out / b.UCHITELSKAYA / "ШК Поставка от 10.11.2026 МГТ учительская.pdf"))
    assert len(pdf.pages) == 2
    assert "СКАЧАНО ШК: 3" in log.text()


def test_leftover_goes_by_article_not_by_old_supply(wb, tmp_path, log, progress):
    """Вчера Zerkalo по ошибке лежало в «учительская» — переносится в «коминтерна»."""
    wb.history["402"] = {"id": 402, "article": "Zerkalo_60x160", "cargoType": 1, "warehouseId": 1}
    wb.status["402"] = ("confirm", "waiting")
    wb.supply("WB-GI-OLD", "Поставка от 09.11.2026 МГТ учительская", orders=[402])
    rep = wb_report(tmp_path / "ВБ.xlsx", [(402, "Zerkalo_60x160", MGT_WH, "-", 1)])
    b.process_wb([str(rep)], TOKEN, str(tmp_path / "c"), log, progress, mode="check", today=TODAY)
    assert "402  Zerkalo_60x160  →  Поставка от 10.11.2026 МГТ коминтерна  ← другое производство" in log.text()
    b.process_wb([str(rep)], TOKEN, str(tmp_path / "o"), log, progress, mode="run", today=TODAY)
    assert [(s["name"], s["orders"]) for s in wb.supplies] == [("Поставка от 10.11.2026 МГТ коминтерна", ["402"])]


def test_leftover_answer_no_keeps_old_supply_and_downloads_it(wb, tmp_path, log, progress):
    leftover(wb)
    wb.order(501, "Zerkalo_50x130")
    rep = wb_report(tmp_path / "ВБ.xlsx", [(401, "Ambient warm 80 krug multi", MGT_WH, "-", 1),
                                           (501, "Zerkalo_50x130", MGT_WH, "-", 1)])
    out = tmp_path / "o"
    b.process_wb([str(rep)], TOKEN, str(out), log, progress, mode="run", today=TODAY,
                 confirm=lambda text: "Не уехало вчера" not in text)
    names = sorted(s["name"] for s in wb.supplies)
    assert names == ["Поставка от 09.11.2026 МГТ учительская", "Поставка от 10.11.2026 МГТ коминтерна"]
    assert (out / b.UCHITELSKAYA / "ШК Поставка от 09.11.2026 МГТ учительская.pdf").exists()
    assert "СКАЧАНО ШК: 2" in log.text()


def test_move_refused_by_wb_keeps_old_and_downloads(wb, tmp_path, log, progress):
    leftover(wb)
    wb.no_move.add("401")
    rep = wb_report(tmp_path / "ВБ.xlsx", [(401, "Ambient warm 80 krug multi", MGT_WH, "-", 1)])
    out = tmp_path / "o"
    b.process_wb([str(rep)], TOKEN, str(out), log, progress, mode="run", today=TODAY)
    text = log.text()
    assert "Не перенесено 401" in text and "Осталось во вчерашней поставке" in text
    assert any(s["name"] == "Поставка от 09.11.2026 МГТ учительская" for s in wb.supplies)
    assert (out / b.UCHITELSKAYA / "ШК Поставка от 09.11.2026 МГТ учительская.pdf").exists()


def test_only_leftovers_no_new_moves_without_adding_question(wb, tmp_path, log, progress):
    leftover(wb)
    rep = wb_report(tmp_path / "ВБ.xlsx", [(401, "Ambient warm 80 krug multi", MGT_WH, "-", 1)])
    out = tmp_path / "o"
    asked = []
    b.process_wb([str(rep)], TOKEN, str(out), log, progress, mode="run", today=TODAY,
                 confirm=lambda text: asked.append(text) or True)
    assert not any(q.startswith("Добавить") for q in asked)   # новых нет — не спрашиваем про добавление
    assert [(s["name"], s["orders"]) for s in wb.supplies] == [("Поставка от 10.11.2026 МГТ учительская", ["401"])]
    assert (out / b.UCHITELSKAYA / "ШК Поставка от 10.11.2026 МГТ учительская.pdf").exists()


def test_leftover_not_in_report_is_shown_and_moved(wb, tmp_path, log, progress):
    """Отчёт сделан старой версией: вчерашнего в нём нет — всё равно видно и переносится (1.39)."""
    leftover(wb)
    wb.order(501, "Zerkalo_50x130")
    rep = wb_report(tmp_path / "ВБ.xlsx", [(501, "Zerkalo_50x130", MGT_WH, "-", 1)])
    b.process_wb([str(rep)], TOKEN, str(tmp_path / "c"), log, progress, mode="check", today=TODAY)
    text = log.text()
    assert "Не уехало вчера, а в отчёте нет: 1 задание" in text
    assert "401  Ambient warm 80 krug multi  (было в «Поставка от 09.11.2026 МГТ учительская»)" in text
    assert "Поставка от 10.11.2026 МГТ учительская — 0 заданий + не уехало вчера 1" in text
    assert "ВСЕГО ШК БУДЕТ СКАЧАНО: 2" in text
    b.process_wb([str(rep)], TOKEN, str(tmp_path / "o"), log, progress, mode="run", today=TODAY)
    assert {s["name"]: s["orders"] for s in wb.supplies}["Поставка от 10.11.2026 МГТ учительская"] == ["401"]


def test_leftover_painted_red_is_left_alone(wb, tmp_path, log, progress):
    leftover(wb)
    rep = wb_report(tmp_path / "ВБ.xlsx", [(401, "Ambient warm 80 krug multi", MGT_WH, "-", 1)],
                    red={"Ambient warm 80 krug multi"})
    b.process_wb([str(rep)], TOKEN, str(tmp_path / "o"), log, progress, mode="run", today=TODAY)
    assert [(s["name"], s["orders"]) for s in wb.supplies] == [("Поставка от 09.11.2026 МГТ учительская", ["401"])]
