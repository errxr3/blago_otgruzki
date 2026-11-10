"""Пункт 6: ленты «объединенные» — артикул сверху, сортировка; подпись времени; битые файлы; пути."""

import io
from datetime import datetime

import pytest
from pypdf import PdfReader

import blago_otgruzki as b
from fakes import label_pdf


def texts(path):
    return [p.extract_text() for p in PdfReader(str(path)).pages]


@pytest.fixture
def folder(tmp_path):
    f = tmp_path / "Ноябрь 2026" / "Озон утро (копия)" / b.KOMINTERNA
    f.mkdir(parents=True)
    (f / "Zerkalo_50x130.pdf").write_bytes(label_pdf(["a-1", "a-2"]))
    (f / "Zerkalo_190x90_X.pdf").write_bytes(label_pdf(["b-1"]))
    (f / "bp_Zerkalo_50x160.pdf").write_bytes(label_pdf(["c-1"]))
    (f / "Alum rama 200x100 black.pdf").write_bytes(label_pdf(["d-1"]))
    return f


def test_ribbon_sorted_and_labelled(folder):
    target = folder.parent / b.combined_name(b.KOMINTERNA)
    assert b.merge_folder(folder, target, 8, 8, footer="Озон 2") == (4, 5)
    assert target.name == "Коминтерна объединенные.pdf"
    pages = texts(target)
    order = [("Zerkalo 50x130", "a-1"), ("Zerkalo 50x130", "a-2"), ("Zerkalo 190x90 X", "b-1"),
             ("bp Zerkalo 50x160", "c-1"), ("Alum rama 200x100 black", "d-1")]
    for page, (label, number) in zip(pages, order):
        assert label in page and number in page and "Озон 2" in page


def test_label_is_horizontal_on_top_for_normal_page(folder):
    target = folder.parent / "лента.pdf"
    b.merge_folder(folder, target, 8, 8)
    page = PdfReader(str(target)).pages[0]
    assert (round(float(page.mediabox.width)), round(float(page.mediabox.height))) == (164, 113)


def test_very_low_page_gets_side_strip_instead(tmp_path):
    f = tmp_path / "low"
    f.mkdir()
    (f / "Zerkalo_1.pdf").write_bytes(label_pdf(["x-1"], size=(164, 15)))
    b.merge_folder(f, tmp_path / "out.pdf", 8, 8)
    page = PdfReader(str(tmp_path / "out.pdf")).pages[0]
    assert round(float(page.mediabox.width)) == round(164 + b.STRIP_WIDTH_MM * b.MM)
    assert "x-1" in page.extract_text() and "Zerkalo 1" in page.extract_text()


def test_empty_folder_removes_old_ribbon(tmp_path):
    f = tmp_path / "empty"
    f.mkdir()
    old = tmp_path / "лента.pdf"
    old.write_bytes(label_pdf(["old"]))
    assert b.merge_folder(f, old, 8, 8) == (0, 0)
    assert not old.exists()


def test_broken_pdf_in_folder_is_skipped_and_reported(folder, log):
    """Один битый файл в папке не должен ронять сборку ленты: остальные склеиваются, о битом — строка."""
    (folder / "Zerkalo_битый.pdf").write_bytes(b"%PDF-1.4 broken")
    out = folder.parent
    (out / b.UCHITELSKAYA).mkdir()
    b.merge_all(out, 8, 8, log)
    assert len(texts(out / "Коминтерна объединенные.pdf")) == 5
    text = log.text()
    assert "Zerkalo_битый.pdf" in text and "не читается" in text
    assert "Коминтерна объединенные.pdf: 5 этикеток, 4 артикулов" in text


def test_merge_all_reports_empty_production(folder, log):
    out = folder.parent
    (out / b.UCHITELSKAYA).mkdir()
    b.merge_all(out, 8, 8, log)
    assert "Учительская: этикеток нет, лента не создана" in log.text()


def test_label_from_filename():
    assert b.label_from_filename("Zerkalo_190x90_X.pdf") == "Zerkalo 190x90 X"
    assert b.label_from_filename("Ambient-80 krug.pdf") == "Ambient 80 krug"


# ---------------------------------------------------------------- время заказа на этикетке

def test_time_stamped_by_barcode_even_if_page_order_differs():
    pdf = label_pdf(["%B-2%", "%B-1%"])           # Ozon отдал страницы в другом порядке
    times = {"1": datetime(2026, 11, 9, 8, 0), "2": datetime(2026, 11, 9, 12, 30)}
    pages, stamped, matched = b.stamp_times(pdf, ["1", "2"], times, {"1": ["%B-1%"], "2": ["%B-2%"]})
    assert (stamped, matched) == (2, 2)
    assert "09.11 12:30" in pages[0].extract_text() and "09.11 08:00" in pages[1].extract_text()


def test_time_not_guessed_when_pages_do_not_match_numbers():
    """Страниц меньше, чем отправлений, штрихкода нет, время разное — чужое время не ставим."""
    pdf = label_pdf(["?"])
    times = {"1": datetime(2026, 11, 9, 8, 0), "2": datetime(2026, 11, 9, 12, 30)}
    pages, stamped, matched = b.stamp_times(pdf, ["1", "2"], times, {})
    assert (stamped, matched) == (0, 0) and "09.11" not in pages[0].extract_text()


def test_broken_label_bytes_raise_for_caller():
    with pytest.raises(Exception):
        b.stamp_times(b"%PDF-1.4 broken", ["1"], {}, {})


def test_kgt_sticker_code_is_collected():
    buf = io.BytesIO()
    from reportlab.pdfgen import canvas
    c = canvas.Canvas(buf, pagesize=(164, 113))
    c.drawString(10, 60, "ii5019607 3009")
    c.save()
    found = {}
    b.stamp_times(buf.getvalue(), ["9-1"], {}, {}, found=found)
    assert found == {"9-1": ["ii50196073009"]}


# ---------------------------------------------------------------- «Не уехало вчера»

def test_left_numbers_from_pdf_excel_csv(tmp_path):
    pdf = tmp_path / "ШК Озон 3.pdf"
    pdf.write_bytes(label_pdf(["12538645-0672-2", "0270161011-0010-1"]))
    import openpyxl
    xlsx = tmp_path / "Озон 3 08.11.xlsx"
    wbk = openpyxl.Workbook()
    wbk.active.append(["Номер отправления"])
    wbk.active.append(["05742108-1571-1"])
    wbk.save(xlsx)
    csv = tmp_path / "выгрузка.csv"
    csv.write_bytes('"Номер заказа";"Номер отправления"\n"44444444-0001";"44444444-0001-3"\n'.encode("cp1251"))
    postings, orders, notes = b.left_numbers([pdf, xlsx, csv])
    assert postings == {"12538645-0672-2", "0270161011-0010-1", "05742108-1571-1", "44444444-0001-3"}
    assert orders == {"05742108-1571", "44444444-0001"}      # номер заказа — только из Excel и CSV
    assert len(notes) == 3


def test_left_numbers_broken_file_is_reported_not_raised(tmp_path):
    bad = tmp_path / "битый.pdf"
    bad.write_bytes(b"not a pdf at all")
    postings, orders, notes = b.left_numbers([bad])
    assert not postings and "не прочитался" in notes[0]
