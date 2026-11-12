"""СДЭК: одна накладная — несколько одинаковых зеркал (колонка количества), 1.40."""

import re
import zipfile

import openpyxl
import pytest
from pypdf import PdfReader

import blago_otgruzki as b
from fakes import label_pdf


def cdek_file(path, rows, qty_title="Количество"):
    wbk = openpyxl.Workbook()
    ws = wbk.active
    ws.append(["№ задания", "Артикул продавца", "трэк", qty_title])
    for row in rows:
        ws.append(list(row))
    wbk.save(path)
    return path


def act_text(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    return " ".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))


ROWS = [("111", "Zerkalo_60x160", "10000000001", 3),     # одна накладная — 3 зеркала
        ("222", "Zerkalo_50x130", "10000000002", 1),
        ("333", "Ambient 80 krug", "10000000003", None)]  # пусто — 1


@pytest.mark.parametrize("title", ["Количество", "кол-во", "Кол-во"])
def test_quantity_column_is_read(tmp_path, title):
    orders = b.read_cdek_orders(cdek_file(tmp_path / "сдэк.xlsx", ROWS, title))
    assert [o["qty"] for o in orders] == [3, 1, 1] and b.pieces(orders) == 5


def test_no_quantity_column_means_one(tmp_path):
    path = tmp_path / "сдэк.xlsx"
    wbk = openpyxl.Workbook()
    wbk.active.append(["№ задания", "Артикул продавца", "трэк"])
    wbk.active.append(["111", "Zerkalo_60x160", "10000000001"])
    wbk.save(path)
    assert b.read_cdek_orders(path)[0]["qty"] == 1


def test_sheet_and_act_count_mirrors_not_waybills(tmp_path, log):
    orders = b.read_cdek_orders(cdek_file(tmp_path / "сдэк.xlsx", ROWS))
    out = tmp_path / "out"
    b.make_cdek(orders, out, log)
    ws = openpyxl.load_workbook(next((out / b.KOMINTERNA).glob("СДЭК сборка *.xlsx"))).active
    rows = [r for r in ws.iter_rows(values_only=True)]
    assert rows[1][:4] == (111, "Zerkalo_60x160", "10000000001", 3)
    totals = {r[5]: r[6] for r in rows[1:] if r[5]}
    assert totals == {"Zerkalo_60x160": 3, "Zerkalo_50x130": 1, "Итого": 4}
    text = act_text(next((out / b.KOMINTERNA).glob("СДЭК акт *.docx")))
    assert re.search(r"Принято\s*:\s*2\s*накладн", text) and re.search(r"мест\s*:\s*4", text)
    assert "Коминтерна: 2 заказов, 4 шт." in log.text()


def test_barcode_printed_once_per_mirror(tmp_path, log, monkeypatch):
    orders = b.read_cdek_orders(cdek_file(tmp_path / "сдэк.xlsx", ROWS))
    monkeypatch.setattr(b, "cdek_on", lambda: True)
    monkeypatch.setattr(b, "cdek_barcodes_api",
                        lambda tracks, log: [label_pdf(["10000000001", "10000000002", "10000000003"])])
    out = tmp_path / "out"
    b.make_cdek(orders, out, log)
    pages = [p.extract_text() for p in PdfReader(str(next((out / b.KOMINTERNA).glob("СДЭК ШК *.pdf")))).pages]
    assert ["10000000001" in p for p in pages] == [True, True, True, False]   # 3 ШК на 3 зеркала, потом 1
    assert "ШК 4 шт." in log.text()


def test_warehouse_gets_one_label_line_per_mirror(tmp_path, log, monkeypatch):
    orders = b.read_cdek_orders(cdek_file(tmp_path / "сдэк.xlsx", ROWS))
    sent = {}
    monkeypatch.setattr(b, "sklad_on", lambda: True)
    monkeypatch.setattr(b, "morning_hold", lambda day, lines, *a, **k: sent.setdefault("lines", lines))
    monkeypatch.setattr(b, "sklad_send_batch",
                        lambda day, market, channel, truck, source, items, log: sent.setdefault("items", items))
    b.make_cdek(orders, tmp_path / "out", log)
    assert {x["posting"]: x["qty"] for x in sent["lines"]} == {"10000000001": 3, "10000000002": 1, "10000000003": 1}
    assert [x["posting"] for x in sent["items"]].count("10000000001") == 3


def test_sborka_total_counts_quantity(tmp_path):
    path = cdek_file(tmp_path / "сдэк.xlsx", ROWS, "кол-во")
    assert b.read_totals(path) == {"Zerkalo_60x160": 3, "Zerkalo_50x130": 1, "Ambient 80 krug": 1}


# ---------------------------------------------------------------- сверка количества с СДЭК (1.41)

def cdek_entity(places, amounts, status="CREATED"):
    return {"entity": {"statuses": [{"code": status, "date_time": "2026-11-12T08:00:00+0300"}],
                       "packages": [{"items": [{"amount": a} for a in amounts[i:i + 1]]} for i in range(places)]}}


@pytest.fixture
def cdek_api(isolated, monkeypatch):
    import json

    import responses
    monkeypatch.setattr(b, "cdek_token", lambda: "test-token")
    answers = {}

    def handle(request):
        track = request.url.split("cdek_number=")[1]
        return 200, {}, json.dumps(answers[track])

    isolated.add_callback(responses.GET, re.compile(r"https://api\.cdek\.ru/v2/orders.*"), callback=handle)
    return answers


def test_quantity_matches_cdek(cdek_api, log):
    cdek_api["10000000001"] = cdek_entity(1, [3])          # одно место, в нём 3 товара
    cdek_api["10000000002"] = cdek_entity(1, [1])
    orders = [{"task": "1", "article": "Z", "track": "10000000001", "qty": 3},
              {"task": "2", "article": "Z", "track": "10000000002", "qty": 1}]
    assert b.cdek_check(orders, log) == 0


def test_quantity_differs_from_cdek(cdek_api, log):
    cdek_api["10000000001"] = cdek_entity(3, [1, 1, 1])    # три места по одному
    orders = [{"task": "1", "article": "Zerkalo_60x160", "track": "10000000001", "qty": 2}]
    assert b.cdek_check(orders, log) == 1
    assert "10000000001 (Zerkalo_60x160): в таблице 2 шт., а в СДЭК 3 (мест 3, товаров 3)" in log.text()


def test_places_used_when_no_goods(cdek_api, log):
    cdek_api["10000000001"] = cdek_entity(2, [])
    orders = [{"task": "1", "article": "Z", "track": "10000000001", "qty": 1}]
    assert b.cdek_check(orders, log) == 1
    assert "в таблице 1 шт., а в СДЭК 2 (мест 2, товаров 0)" in log.text()


def test_no_amounts_in_cdek_no_warning(cdek_api, log):
    cdek_api["10000000001"] = {"entity": {"statuses": [{"code": "CREATED"}]}}
    assert b.cdek_check([{"task": "1", "article": "Z", "track": "10000000001", "qty": 3}], log) == 0
