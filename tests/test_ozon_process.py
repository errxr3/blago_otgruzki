"""Пункты 1, 3, 4, 5, 6: кабинеты, статусы, «Проверить», этикетки, _X, раскладка и ленты."""

import pytest
import requests
from pypdf import PdfReader

import blago_otgruzki as b
from fakes import CABINETS, ozon_report, posting, t

KGT = "FBS КГТ Тверь"
READ_ONLY = {"/v4/posting/fbs/list"}
CHANGING = {"/v4/posting/fbs/ship", "/v3/posting/fbs/package-label/create", "/v1/carriage/create",
            "/v1/carriage/approve", "/v1/carriage/pass/create"}


@pytest.fixture
def day(ozon, tmp_path):
    """Общий отчёт: 3 кабинета + КГТ у Z и R, оба статуса, красный и жёлтый артикулы,
    многоштучный заказ, один артикул в кабинетах 1 и 3, неизвестный номер."""
    z, r, iz = ozon.cabinets["Z"], ozon.cabinets["R"], ozon.cabinets["IZ"]
    z["10000001-0001-1"] = posting(b.READY, ("Zerkalo_190x90", 1))
    z["10000002-0001-1"] = posting(b.PACKING, ("Zerkalo_50x130", 1))
    z["10000003-0001-1"] = posting(b.PACKING, ("Ambient 80 krug", 3))
    z["10000004-0001-1"] = posting(b.READY, ("Alum rama 200x100 black", 1), warehouse=KGT)
    z["10000005-0001-1"] = posting(b.PACKING, ("Zerkalo_60x160", 1))
    z["10000006-0001-1"] = posting(b.PACKING, ("Zerkalo_70x200", 1))
    z["10000007-0001-1"] = posting(b.PACKING, ("Zerkalo_70x200", 1))
    z["10000008-0001-1"] = posting(b.PACKING, ("Zerkalo_70x200", 1))
    r["20000001-0001-1"] = posting(b.PACKING, ("bp_Zerkalo_50x160", 1))
    r["20000002-0001-1"] = posting(b.READY, ("Ambient 100x80", 1), warehouse="FBS - СГТ Тверь")
    iz["30000001-0001-1"] = posting(b.READY, ("Zerkalo_190x90", 1))
    iz["30000002-0001-1"] = posting(b.PACKING, ("Ambient 80 krug", 1))
    folder = tmp_path / "Ноябрь 2026" / "10.11 Озон утро"   # кириллица и пробелы в пути
    folder.mkdir(parents=True)
    report = ozon_report(folder / "Озон 10.11.xlsx", [
        ("10000001-0001-1", t(8), "Zerkalo_190x90", 1),
        ("10000002-0001-1", t(8, 5), "Zerkalo_50x130", 1),
        ("10000003-0001-1", t(8, 10), "Ambient 80 krug", 3),
        ("10000004-0001-1", t(8, 15), "Alum rama 200x100 black", 1),
        ("10000005-0001-1", t(8, 20), "Zerkalo_60x160", 1),
        ("10000008-0001-1", t(12), "Zerkalo_70x200", 1),   # самый новый — не едет
        ("10000006-0001-1", t(7), "Zerkalo_70x200", 1),
        ("10000007-0001-1", t(9), "Zerkalo_70x200", 1),
        ("20000001-0001-1", t(9, 5), "bp_Zerkalo_50x160", 1),
        ("20000002-0001-1", t(9, 10), "Ambient 100x80", 1),
        ("30000001-0001-1", t(9, 15), "Zerkalo_190x90", 1),
        ("30000002-0001-1", t(9, 20), "Ambient 80 krug", 1),
        ("99999999-0001-1", t(9, 25), "Zerkalo_50x130", 1),  # нет ни в одном кабинете
    ], red={"Zerkalo_60x160"}, yellow={"Zerkalo_70x200": 2})
    return report, folder / "Этикетки"


def run(day, log, progress, **kw):
    report, out = day
    kw.setdefault("autoship", True)
    return b.process(str(report), None, CABINETS, 14, str(out), log, progress, **kw)


def pages(path):
    return [p.extract_text() for p in PdfReader(str(path)).pages]


# ---------------------------------------------------------------- «Проверить отчёт»

def test_check_counts_both_statuses_and_changes_nothing(day, ozon, log, progress):
    before = {cab: dict(items) for cab, items in ozon.cabinets.items()}
    run(day, log, progress, download=False)
    text = log.text()
    assert "ВСЕГО ШК БУДЕТ СКАЧАНО: 12" in text
    assert "из «Ожидают отгрузки»:  4" in text
    assert "из «Ожидают сборки»:    8" in text
    assert "Z — 8,  R — 2,  IZ — 2" in text
    assert "Коминтерна — 7,  Учительская — 5" in text
    assert "99999999-0001-1" in text and "кабинет не найден" in text
    # ни одного запроса, который что-то меняет: только чтение списка заказов
    assert set(ozon.paths()) == READ_ONLY
    assert not set(ozon.paths()) & CHANGING
    assert ozon.cabinets == before
    assert not day[1].exists()   # папка этикеток даже не создана


def test_check_with_autoship_off_warns(day, log, progress):
    run(day, log, progress, download=False, autoship=False)
    text = log.text()
    assert "ГАЛКА АВТОСБОРКИ СНЯТА" in text and "скачается только 4 ШК" in text


def test_cabinet_found_by_posting_number(day, log, progress):
    run(day, log, progress, download=False)
    text = log.text()
    assert "✔ Z: из отчёта найдено 6 — ожидают отгрузки 2, ожидают сборки 4" in text
    assert "✔ R: из отчёта найдено 2 — ожидают отгрузки 1, ожидают сборки 1" in text
    assert "✔ IZ: из отчёта найдено 2 — ожидают отгрузки 1, ожидают сборки 1" in text


# ---------------------------------------------------------------- «Скачать и собрать»

def test_full_run_ships_splits_and_sorts_into_folders(day, ozon, log, progress):
    form = run(day, log, progress)
    out = day[1]
    # собраны только едущие «ожидают сборки»: красный и третий жёлтый не тронуты
    shipped = [body["posting_number"] for _, path, body in ozon.calls if path == "/v4/posting/fbs/ship"]
    assert sorted(shipped) == ["10000002-0001-1", "10000003-0001-1", "10000006-0001-1", "10000007-0001-1",
                               "20000001-0001-1", "30000002-0001-1"]
    assert ozon.cabinets["Z"]["10000005-0001-1"]["status"] == b.PACKING   # красный
    assert ozon.cabinets["Z"]["10000008-0001-1"]["status"] == b.PACKING   # жёлтый сверх количества
    # многоштучный заказ: 3 штуки — 3 грузоместа по одной штуке
    multi = next(body for _, path, body in ozon.calls
                 if path == "/v4/posting/fbs/ship" and body["posting_number"] == "10000003-0001-1")
    assert len(multi["packages"]) == 3
    assert all(p["products"][0]["quantity"] == 1 for p in multi["packages"])
    # один PDF на артикул; второй кабинет с тем же артикулом — с _X
    kom = sorted(p.name for p in (out / b.KOMINTERNA).glob("*.pdf"))
    uch = sorted(p.name for p in (out / b.UCHITELSKAYA).glob("*.pdf"))
    assert kom == ["Alum rama 200x100 black.pdf", "Zerkalo_190x90.pdf", "Zerkalo_190x90_X.pdf",
                   "Zerkalo_50x130.pdf", "Zerkalo_70x200.pdf", "bp_Zerkalo_50x160.pdf"]
    assert uch == ["Ambient 100x80.pdf", "Ambient 80 krug.pdf", "Ambient 80 krug_X.pdf"]
    assert len(pages(out / b.KOMINTERNA / "Zerkalo_70x200.pdf")) == 2      # одинаковые артикулы склеены
    assert len(pages(out / b.UCHITELSKAYA / "Ambient 80 krug.pdf")) == 3   # три грузоместа — три этикетки
    assert "10000001-0001-1" in pages(out / b.KOMINTERNA / "Zerkalo_190x90.pdf")[0]     # кабинет Z
    assert "30000001-0001-1" in pages(out / b.KOMINTERNA / "Zerkalo_190x90_X.pdf")[0]   # кабинет IZ
    assert "Скачано этикеток: 12" in log.text()
    assert "СКАЧАНО ШК: 12   (Коминтерна 7 · Учительская 5)" in log.text()
    # что не поехало и что не скачалось — в файлах
    skipped = (out / "НЕ_ЕДУТ.txt").read_text(encoding="utf-8")
    assert "10000005-0001-1\tкрасный" in skipped and "10000008-0001-1" in skipped
    assert "99999999-0001-1" in (out / "НЕ_СКАЧАНО.txt").read_text(encoding="utf-8")
    # склады этого запуска — для отгрузки: КГТ попал
    assert KGT.lower() in form["Z"] and set(form) == {"Z", "R", "IZ"}


def test_combined_ribbons_have_article_on_top_and_sorted(day, log, progress):
    run(day, log, progress, footer="Озон утро")
    out = day[1]
    kom = pages(out / "Коминтерна объединенные.pdf")
    uch = pages(out / "Учительская объединенные.pdf")
    assert len(kom) == 7 and len(uch) == 5
    labels = ["Zerkalo 70x200", "Zerkalo 70x200", "Zerkalo 50x130", "Zerkalo 190x90 X", "Zerkalo 190x90",
              "bp Zerkalo 50x160", "Alum rama 200x100 black"]
    for text, label in zip(kom, labels):   # сортировка Я→А по имени файла, артикул написан на странице
        assert label in text
    assert [("X" in x.split("krug")[-1]) for x in uch[:4]] == [True, False, False, False]
    assert "Ambient 100x80" in uch[4]
    assert all("Озон утро" in text for text in kom + uch)   # надпись волны на каждой этикетке
    # надпись артикула горизонтальная: страница не расширена под боковую полосу
    first = PdfReader(str(out / "Коминтерна объединенные.pdf")).pages[0]
    assert round(float(first.mediabox.width)) == 164


def test_order_time_is_stamped_on_labels(day, log, progress):
    run(day, log, progress)
    text = pages(day[1] / b.KOMINTERNA / "Zerkalo_190x90.pdf")[0]
    assert "09.11 08:00" in text
    assert "Время заказа подписано на 12 из 12 ШК" in log.text()


def test_rerun_same_day_makes_no_duplicates(day, ozon, log, progress):
    run(day, log, progress)
    out = day[1]
    first = {p.relative_to(out).as_posix(): len(pages(p)) for p in out.rglob("*.pdf")}
    ships = ozon.count("/v4/posting/fbs/ship")
    run(day, log, progress)
    again = {p.relative_to(out).as_posix(): len(pages(p)) for p in out.rglob("*.pdf")}
    assert again == first                                   # те же файлы, те же страницы
    assert ozon.count("/v4/posting/fbs/ship") == ships      # второй раз ничего не собрано
    assert "добавлено ещё 2 грузоместа" in log.text()        # грузоместа многоштучного найдены по заказу


def test_autoship_off_does_not_ship(day, ozon, log, progress):
    run(day, log, progress, autoship=False)
    assert ozon.count("/v4/posting/fbs/ship") == 0
    assert "Скачано этикеток: 4" in log.text()
    assert "автосборка выключена" in (day[1] / "НЕ_СКАЧАНО.txt").read_text(encoding="utf-8")


def test_user_says_no_nothing_happens(day, ozon, log, progress):
    run(day, log, progress, confirm=lambda text: False)
    assert set(ozon.paths()) == READ_ONLY
    assert "Отменено — ничего не собрано и не скачано" in log.text()
    assert not day[1].exists()


def test_empty_report_does_nothing(ozon, tmp_path, log, progress):
    report = ozon_report(tmp_path / "пусто.xlsx", [])
    assert b.process(str(report), None, CABINETS, 14, str(tmp_path / "out"), log, progress, autoship=True) is None
    assert ozon.calls == []


def test_unknown_article_goes_to_uchitelskaya(ozon, tmp_path, log, progress):
    ozon.cabinets["Z"]["1-1"] = posting(b.READY, ("Новинка 2027", 1))
    report = ozon_report(tmp_path / "r.xlsx", [("1-1", t(8), "Новинка 2027", 1)])
    b.process(str(report), None, CABINETS, 14, str(tmp_path / "out"), log, progress)
    assert (tmp_path / "out" / b.UCHITELSKAYA / "Новинка 2027.pdf").exists()


def test_article_with_slash_gets_safe_file_name(ozon, tmp_path, log, progress):
    ozon.cabinets["Z"]["1-1"] = posting(b.READY, ("Ambient 80/60", 1))
    report = ozon_report(tmp_path / "r.xlsx", [("1-1", t(8), "Ambient 80/60", 1)])
    b.process(str(report), None, CABINETS, 14, str(tmp_path / "out"), log, progress)
    assert (tmp_path / "out" / b.UCHITELSKAYA / "Ambient 80_60.pdf").exists()


def test_wrong_status_is_reported(ozon, tmp_path, log, progress):
    ozon.cabinets["Z"]["1-1"] = posting("delivering", ("Zerkalo_50x130", 1))
    report = ozon_report(tmp_path / "r.xlsx", [("1-1", t(8), "Zerkalo_50x130", 1)])
    b.process(str(report), None, CABINETS, 14, str(tmp_path / "out"), log, progress, download=False)
    assert "статус «доставляется»" in log.text()


# ---------------------------------------------------------------- сбои Ozon

def test_429_on_list_is_retried(day, ozon, log, progress):
    ozon.fail[("R", "/v4/posting/fbs/list")] = [(429, {"message": "too many"}), (429, {"message": "too many"})]
    run(day, log, progress, download=False)
    assert "ВСЕГО ШК БУДЕТ СКАЧАНО: 12" in log.text()
    assert ozon.count("/v4/posting/fbs/list", "R") == 3


def test_500_on_one_cabinet_stops_everything(day, ozon, log, progress):
    ozon.fail[("R", "/v4/posting/fbs/list")] = (500, {"message": "internal"})
    assert run(day, log, progress) is None
    assert "ОСТАНОВЛЕНО: не загрузился кабинет R" in log.text()
    assert not set(ozon.paths()) & CHANGING        # неполный день не собираем
    assert not day[1].exists()


def test_timeout_on_list_stops_everything(day, ozon, log, progress):
    ozon.fail[("IZ", "/v4/posting/fbs/list")] = requests.Timeout("read timeout")
    assert run(day, log, progress) is None
    text = log.text()
    assert "IZ: не удалось получить список заказов — Ozon не отвечает" in text
    assert not set(ozon.paths()) & CHANGING


def test_ship_error_keeps_the_rest_going(day, ozon, log, progress):
    ozon.fail[("R", "/v4/posting/fbs/ship")] = (500, {"message": "boom"})
    run(day, log, progress)
    failed = (day[1] / "НЕ_СКАЧАНО.txt").read_text(encoding="utf-8")
    assert "20000001-0001-1" in failed and "не удалось собрать" in failed
    assert "Скачано этикеток: 11" in log.text()
    assert not (day[1] / b.KOMINTERNA / "bp_Zerkalo_50x160.pdf").exists()


def test_ship_timeout_is_never_repeated(day, ozon, log, progress):
    """Сборка могла пройти на стороне Ozon — повторять её нельзя."""
    ozon.fail[("R", "/v4/posting/fbs/ship")] = requests.Timeout("read timeout")
    run(day, log, progress)
    assert ozon.count("/v4/posting/fbs/ship", "R") == 1
    assert "нет связи" in (day[1] / "НЕ_СКАЧАНО.txt").read_text(encoding="utf-8")


def test_label_error_leaves_no_empty_pdf(day, ozon, log, progress):
    """Этикетки артикула не скачались — пустого PDF в папке быть не должно."""
    ozon.fail[("R", "/v3/posting/fbs/package-label/create")] = (500, {"message": "boom"})
    run(day, log, progress)
    out = day[1]
    assert not (out / b.KOMINTERNA / "bp_Zerkalo_50x160.pdf").exists()
    assert not (out / b.UCHITELSKAYA / "Ambient 100x80.pdf").exists()
    failed = (out / "НЕ_СКАЧАНО.txt").read_text(encoding="utf-8")
    assert "20000001-0001-1" in failed and "20000002-0001-1" in failed
    assert "Скачано этикеток: 10" in log.text()
    assert len(pages(out / "Коминтерна объединенные.pdf")) == 6


def test_unprinted_label_reported_others_downloaded(day, ozon, log, progress):
    ozon.unprinted.add("10000001-0001-1")
    run(day, log, progress)
    assert "10000001-0001-1" in (day[1] / "НЕ_СКАЧАНО.txt").read_text(encoding="utf-8")
    assert "Скачано этикеток: 11" in log.text()


def test_broken_label_pdf_does_not_crash_after_shipping(day, ozon, log, progress):
    """Ozon отдал битый файл этикеток уже после сборки: программа не падает, а пишет, что не скачано."""
    ozon.label_bytes = b"%PDF-1.4\nthis is not a real pdf"
    run(day, log, progress)
    failed = (day[1] / "НЕ_СКАЧАНО.txt").read_text(encoding="utf-8")
    assert "10000001-0001-1" in failed and "файл этикеток" in failed
    assert "Скачано этикеток: 0" in log.text()
    assert not list(day[1].rglob("*.pdf"))


def test_not_a_pdf_answer_is_reported(day, ozon, log, progress):
    ozon.label_bytes = b"<html>error</html>"
    run(day, log, progress)
    assert "не удалось скачать файл этикеток" in (day[1] / "НЕ_СКАЧАНО.txt").read_text(encoding="utf-8")


def test_same_article_in_three_cabinets_gets_x_and_xx(ozon, tmp_path, log, progress):
    for cab, num in (("Z", "1-1"), ("R", "2-1"), ("IZ", "3-1")):
        ozon.cabinets[cab][num] = posting(b.READY, ("Zerkalo_190x90", 1))
    report = ozon_report(tmp_path / "r.xlsx", [(n, t(8), "Zerkalo_190x90", 1) for n in ("1-1", "2-1", "3-1")])
    out = tmp_path / "out"
    b.process(str(report), None, CABINETS, 14, str(out), log, progress)
    assert sorted(p.name for p in (out / b.KOMINTERNA).glob("*.pdf")) == [
        "Zerkalo_190x90.pdf", "Zerkalo_190x90_X.pdf", "Zerkalo_190x90_XX.pdf"]
    assert "Zerkalo_190x90_X   → Коминтерна  [R]" in log.text()


def test_article_only_in_second_and_third_cabinet(ozon, tmp_path, log, progress):
    """Суффикс считается среди кабинетов, где артикул есть: первый из них — без суффикса."""
    ozon.cabinets["R"]["2-1"] = posting(b.READY, ("Ambient 80 krug", 1))
    ozon.cabinets["IZ"]["3-1"] = posting(b.READY, ("Ambient 80 krug", 1))
    report = ozon_report(tmp_path / "r.xlsx", [("2-1", t(8), "Ambient 80 krug", 1), ("3-1", t(9), "Ambient 80 krug", 1)])
    out = tmp_path / "out"
    b.process(str(report), None, CABINETS, 14, str(out), log, progress)
    assert sorted(p.name for p in (out / b.UCHITELSKAYA).glob("*.pdf")) == ["Ambient 80 krug.pdf",
                                                                            "Ambient 80 krug_X.pdf"]


def test_already_packed_multi_piece_posting_warns_one_label(ozon, tmp_path, log, progress):
    ozon.cabinets["Z"]["1-1"] = posting(b.READY, ("Zerkalo_50x130", 3))
    report = ozon_report(tmp_path / "r.xlsx", [("1-1", t(8), "Zerkalo_50x130", 3)])
    b.process(str(report), None, CABINETS, 14, str(tmp_path / "out"), log, progress, download=False)
    text = log.text()
    assert "НЕСКОЛЬКО ШТУК, НО ТОЛЬКО 1 ШК" in text and "3 шт. → 1 ШК" in text
    assert "ВСЕГО ШК БУДЕТ СКАЧАНО: 1" in text


def test_fewer_label_pages_than_postings_is_counted_honestly(ozon, tmp_path, log, progress):
    """Ozon отдал файл, где этикеток меньше, чем отправлений: считаем страницы и предупреждаем."""
    from fakes import label_pdf
    for i in (1, 2, 3):
        ozon.cabinets["Z"][f"{i}-1"] = posting(b.READY, ("Zerkalo_50x130", 1))
    report = ozon_report(tmp_path / "r.xlsx", [(f"{i}-1", t(8), "Zerkalo_50x130", 1) for i in (1, 2, 3)])
    ozon.label_bytes = label_pdf(["1-1", "2-1"])          # двух этикеток вместо трёх
    out = tmp_path / "out"
    b.process(str(report), None, CABINETS, 14, str(out), log, progress)
    text = log.text()
    assert "Скачано этикеток: 2" in text
    assert "Zerkalo_50x130" in text and "этикеток 2, а отправлений 3" in text
