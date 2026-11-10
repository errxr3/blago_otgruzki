"""Пункт 2: сводная (G — артикул, H — сумма, I — «едет»), красные и жёлтые артикулы."""

import openpyxl
import pytest

import blago_otgruzki as b
from fakes import OZON_HEADER, ozon_report, t


def test_pivot_is_in_columns_g_h_i(tmp_path):
    path = ozon_report(tmp_path / "Озон.xlsx", [("1-1", t(8), "Zerkalo_50x130", 2)], yellow={"Zerkalo_50x130": 1})
    ws = openpyxl.load_workbook(path).worksheets[0]
    assert ws["G2"].value == "Zerkalo_50x130" and ws["H2"].value == 2 and ws["I2"].value == 1


def test_red_article_excludes_every_posting_with_it(tmp_path):
    path = ozon_report(tmp_path / "r.xlsx", [
        ("1-1", t(8), "Zerkalo_50x130", 1),
        ("2-1", t(9), "Zerkalo_60x160", 1),
        ("3-1", t(10), "Zerkalo_60x160", 2),
    ], red={"Zerkalo_60x160"})
    rpt = b.read_report(path)
    assert list(rpt.postings) == ["1-1"]
    assert sorted(n for _, n in rpt.red) == ["2-1", "3-1"]


def test_yellow_takes_oldest_orders_first(tmp_path):
    """Едут самые старые по времени, даже если в файле они стоят ниже."""
    path = ozon_report(tmp_path / "y.xlsx", [
        ("new-1", t(12), "Zerkalo_70x200", 1),
        ("old-1", t(7), "Zerkalo_70x200", 1),
        ("mid-1", t(9), "Zerkalo_70x200", 1),
    ], yellow={"Zerkalo_70x200": 2})
    rpt = b.read_report(path)
    assert set(rpt.postings) == {"old-1", "mid-1"}
    assert [n for _, n, _ in rpt.yellow] == ["new-1"]
    limit = rpt.limits[0]
    assert (limit["total"], limit["limit"], limit["taken"]) == (3, 2, 2)


def test_yellow_limit_zero_nothing_goes(tmp_path):
    path = ozon_report(tmp_path / "y0.xlsx", [("1-1", t(8), "Zerkalo_70x200", 1), ("2-1", t(9), "Ambient 80", 1)],
                       yellow={"Zerkalo_70x200": 0})
    rpt = b.read_report(path)
    assert list(rpt.postings) == ["2-1"]
    assert [n for _, n, _ in rpt.yellow] == ["1-1"]


def test_yellow_limit_above_ordered_everything_goes(tmp_path):
    path = ozon_report(tmp_path / "ybig.xlsx", [("1-1", t(8), "Zerkalo_70x200", 1), ("2-1", t(9), "Zerkalo_70x200", 2)],
                       yellow={"Zerkalo_70x200": 99})
    rpt = b.read_report(path)
    assert set(rpt.postings) == {"1-1", "2-1"} and not rpt.yellow
    assert rpt.limits[0]["taken"] == 3


def test_yellow_without_number_nothing_goes_and_note(tmp_path):
    path = ozon_report(tmp_path / "ynone.xlsx", [("1-1", t(8), "Zerkalo_70x200", 1)], yellow={"Zerkalo_70x200": None})
    rpt = b.read_report(path)
    assert not rpt.postings and len(rpt.yellow) == 1
    assert any("не указано" in note for note in rpt.notes)


def test_multi_piece_posting_never_partial(tmp_path):
    """В одном отправлении 3 шт., жёлтым разрешено 2: не едет целиком, место отдаётся следующему."""
    path = ozon_report(tmp_path / "multi.xlsx", [
        ("big-1", t(7), "Zerkalo_70x200", 3),
        ("small-1", t(9), "Zerkalo_70x200", 1),
        ("small-2", t(10), "Zerkalo_70x200", 1),
        ("late-1", t(11), "Zerkalo_70x200", 1),
    ], yellow={"Zerkalo_70x200": 2})
    rpt = b.read_report(path)
    assert set(rpt.postings) == {"small-1", "small-2"}
    reasons = {n: why for _, n, why in rpt.yellow}
    assert "частично собрать нельзя" in reasons["big-1"]
    assert "уже набраны" in reasons["late-1"]


def test_posting_with_red_position_skipped_whole_with_warning(tmp_path, log):
    """Два товара в одном отправлении, один красный: не едет ни один, в журнале понятное предупреждение."""
    path = ozon_report(tmp_path / "mixed.xlsx", [
        ("mix-1", t(8), "Zerkalo_60x160", 1),
        ("mix-1", t(8), "Ambient 80 krug", 1),
        ("ok-1", t(9), "Ambient 80 krug", 1),
    ], red={"Zerkalo_60x160"})
    rpt = b.read_report(path)
    assert list(rpt.postings) == ["ok-1"]
    assert rpt.mixed == [("mix-1", "красный, не едет", ["Zerkalo_60x160"], ["Ambient 80 krug"])]
    b.log_report(rpt, log)
    text = log.text()
    assert "НЕ ПОЕДЕТ ЦЕЛИКОМ" in text and "mix-1" in text and "Ambient 80 krug" in text


def test_posting_with_yellow_overflow_and_other_item_skipped_whole(tmp_path):
    path = ozon_report(tmp_path / "mixed_y.xlsx", [
        ("a-1", t(7), "Zerkalo_70x200", 1),
        ("mix-1", t(8), "Zerkalo_70x200", 1),
        ("mix-1", t(8), "Ambient 80 krug", 1),
    ], yellow={"Zerkalo_70x200": 1})
    rpt = b.read_report(path)
    assert list(rpt.postings) == ["a-1"]
    assert rpt.mixed[0][0] == "mix-1" and rpt.mixed[0][3] == ["Ambient 80 krug"]


def test_red_row_in_list_excludes_only_that_posting(tmp_path):
    path = ozon_report(tmp_path / "row.xlsx", [("1-1", t(8), "Zerkalo_50x130", 1), ("2-1", t(9), "Zerkalo_50x130", 1)],
                       red_rows={"2-1"})
    rpt = b.read_report(path)
    assert list(rpt.postings) == ["1-1"] and [n for _, n in rpt.red] == ["2-1"]


def test_red_wins_over_yellow_for_same_article(tmp_path):
    """Один артикул встречается в сводной дважды (красный и жёлтый) — красный главнее."""
    path = ozon_report(tmp_path / "both.xlsx", [("1-1", t(8), "Zerkalo_50x130", 1)], red={"Zerkalo_50x130"})
    wbk = openpyxl.load_workbook(path)
    ws = wbk.worksheets[0]
    from fakes import YELLOW
    ws.cell(5, 7, "Zerkalo_50x130").fill = YELLOW
    ws.cell(5, 9, 1)
    wbk.save(path)
    rpt = b.read_report(path)
    assert not rpt.postings and not rpt.limits and len(rpt.red) == 1


def test_no_time_column_uses_row_order_and_says_so(tmp_path):
    path = tmp_path / "notime.xlsx"
    b.write_report(path, "Озон", ["Номер отправления", "Артикул", "Количество"],
                   [["first-1", "Zerkalo_70x200", 1], ["second-1", "Zerkalo_70x200", 1]], 2, "Сумма")
    wbk = openpyxl.load_workbook(path)
    ws = wbk.worksheets[0]
    from fakes import YELLOW
    ws.cell(2, 6).fill = YELLOW
    ws.cell(2, 8, 1)
    wbk.save(path)
    rpt = b.read_report(path)
    assert list(rpt.postings) == ["first-1"]
    assert any("по порядку строк" in note for note in rpt.notes)


def test_report_with_header_only_is_empty(tmp_path):
    path = ozon_report(tmp_path / "empty.xlsx", [])
    rpt = b.read_report(path)
    assert rpt.total == 0 and not rpt.postings


def test_completely_empty_sheet_gives_clear_error(tmp_path):
    path = tmp_path / "blank.xlsx"
    openpyxl.Workbook().save(path)
    with pytest.raises(ValueError, match="Номер отправления"):
        b.read_report(path)


def test_report_without_article_column_gives_clear_error(tmp_path):
    path = tmp_path / "noart.xlsx"
    wbk = openpyxl.Workbook()
    wbk.active.append(["Номер отправления", "Что-то"])
    wbk.save(path)
    with pytest.raises(ValueError, match="Артикул"):
        b.read_report(path)


def test_bad_quantity_counts_as_one(tmp_path):
    path = ozon_report(tmp_path / "qty.xlsx", [("1-1", t(8), "Zerkalo_50x130", 1)])
    wbk = openpyxl.load_workbook(path)
    wbk.worksheets[0]["D2"] = "много"
    wbk.save(path)
    assert b.read_report(path).postings == {"1-1": {"Zerkalo_50x130": 1}}


def test_report_file_is_not_modified(tmp_path):
    path = ozon_report(tmp_path / "ro.xlsx", [("1-1", t(8), "Zerkalo_50x130", 1)], red={"Zerkalo_50x130"})
    before = path.read_bytes()
    b.read_report(path)
    assert path.read_bytes() == before


def test_cyrillic_and_spaces_in_path(tmp_path):
    folder = tmp_path / "Ноябрь 2026" / "10.11 Озон утро"
    folder.mkdir(parents=True)
    path = ozon_report(folder / "Озон 10.11 (копия).xlsx", [("1-1", t(8), "Zerkalo_50x130", 1)])
    assert list(b.read_report(path).postings) == ["1-1"]


@pytest.mark.parametrize("article, production", [
    ("Zerkalo_190x90", b.KOMINTERNA),
    ("bp_Zerkalo_50x160", b.KOMINTERNA),
    ("Alum rama 200x100 black", b.KOMINTERNA),
    ("  alum noga 170x70 white ", b.KOMINTERNA),
    ("Ambient 80 krug multi", b.UCHITELSKAYA),
    ("Ambient warm 100x80 gorizont multi", b.UCHITELSKAYA),
    ("Неизвестный артикул 123", b.UCHITELSKAYA),
])
def test_production_split(article, production):
    """Пункт 6: Zerkalo и Alum — Коминтерна, остальное — Учительская."""
    assert b.production_of(article) == production


def test_header_names_are_those_of_owner_report():
    assert OZON_HEADER[0] in b.COL_NUM and OZON_HEADER[2] in b.COL_ART and OZON_HEADER[3] in b.COL_QTY
