"""Утренние отчёты Ozon (Озон и КГТ из трёх кабинетов) и связь со складом BLAGO — всё на подделках."""

import json
import re

import openpyxl
import pytest
import requests
import responses

import blago_otgruzki as b
from fakes import CABINETS, ozon_report, posting, t

SKLAD = "https://sklad.test"


@pytest.fixture
def cabinets(ozon):
    z, r, iz = ozon.cabinets["Z"], ozon.cabinets["R"], ozon.cabinets["IZ"]
    z["1-1"] = posting(b.PACKING, ("Zerkalo_50x130", 2))
    z["2-1"] = posting(b.READY, ("Alum rama 200x100", 1), warehouse="FBS КГТ Тверь")
    z["3-1"] = posting("delivering", ("Zerkalo_60x160", 1))
    r["4-1"] = posting(b.READY, ("Ambient 80 krug", 1))
    r["5-1"] = posting(b.PACKING, ("Ambient 200x100", 1), warehouse="FBS - СГТ Тверь")
    iz["6-1"] = posting(b.PACKING, ("Zerkalo_50x130", 1))
    return ozon


def rows(path):
    ws = openpyxl.load_workbook(path).worksheets[0]
    return [(r[0], r[2], r[3]) for r in ws.iter_rows(min_row=2, max_col=4, values_only=True) if r[0]]


def test_morning_reports_split_ozon_and_kgt(cabinets, tmp_path, log):
    out = tmp_path / "Сборка 10.11"
    out.mkdir()
    data, files = b.ozon_reports(CABINETS, 14, out, log, b.OZON_STATUSES,
                                 {"ozon": "Озон 10.11", "kgt": "КГТ 10.11"})
    assert data == {"ozon": {"Zerkalo_50x130": 3, "Ambient 80 krug": 1},
                    "kgt": {"Alum rama 200x100": 1, "Ambient 200x100": 1}}
    # сначала «ожидают сборки» по кабинетам Z, R, IZ, потом «ожидают отгрузки»; чужие статусы не берём
    assert rows(files["ozon"]) == [("1-1", "Zerkalo_50x130", 2), ("6-1", "Zerkalo_50x130", 1),
                                   ("4-1", "Ambient 80 krug", 1)]
    assert sorted(rows(files["kgt"])) == [("2-1", "Alum rama 200x100", 1), ("5-1", "Ambient 200x100", 1)]
    assert set(cabinets.paths()) == {"/v4/posting/fbs/list"}       # отчёты ничего не меняют
    # отчёт читается обратно программой: сводная стоит в G/H
    assert b.read_report(files["ozon"]).postings == {"1-1": {"Zerkalo_50x130": 2}, "6-1": {"Zerkalo_50x130": 1},
                                                     "4-1": {"Ambient 80 krug": 1}}


def test_morning_report_not_made_if_cabinet_failed(cabinets, tmp_path, log):
    cabinets.fail[("R", "/v4/posting/fbs/list")] = (500, {"message": "boom"})
    with pytest.raises(RuntimeError, match="отчёт НЕ сделан — не загрузился кабинет R"):
        b.ozon_reports(CABINETS, 14, tmp_path, log, b.OZON_STATUSES, {"ozon": "Озон 10.11"})
    assert not list(tmp_path.glob("*.xlsx"))


def test_left_yesterday_is_not_taken(cabinets, tmp_path, log):
    data, _ = b.ozon_reports(CABINETS, 14, tmp_path, log, b.OZON_STATUSES, {"ozon": "Озон 10.11"},
                             skip=({"4-1"}, set()))
    assert data == {"ozon": {"Zerkalo_50x130": 3}}
    assert "Не уехало вчера: 1 отправление в отчёт не беру (R 1)" in log.text()


def test_wave_takes_only_awaiting_packaging(cabinets, tmp_path, log):
    data, _ = b.ozon_reports(CABINETS, 14, tmp_path, log, (b.PACKING,), {"ozon": "Озон 2 10.11"})
    assert data == {"ozon": {"Zerkalo_50x130": 3}}


# ---------------------------------------------------------------- склад BLAGO

@pytest.fixture
def sklad(isolated, monkeypatch):
    monkeypatch.setitem(b.SKLAD, "url", SKLAD)
    monkeypatch.setitem(b.SKLAD, "key", "test-sklad-key")
    calls, answers = [], {"decisions": {}, "fail_batches": False}

    def handle(request):
        path = request.url.replace(SKLAD, "")
        body = json.loads(request.body)
        calls.append((path, body))
        if path == "/api/v1/demand":
            postings = {line["posting"]: {"decision": answers["decisions"].get(line["posting"], "go"), "reason": "тест"}
                        for line in body["lines"]}
            return 200, {}, json.dumps({"ok": True, "postings": postings, "questions": []})
        if answers["fail_batches"]:
            return 500, {}, json.dumps({"ok": False, "error": "упал"})
        return 200, {}, json.dumps({"ok": True, "added": len(body["items"]), "updated": 0})

    isolated.add_callback(responses.POST, re.compile(SKLAD + "/.*"), callback=handle)
    answers["calls"] = calls
    return answers


def test_sklad_decides_but_owner_colors_win(sklad, log):
    sklad["decisions"] = {"no-1": "no", "later-1": "later", "yellow-1": "no"}
    postings = {"go-1": {"A": 1}, "no-1": {"B": 2}, "later-1": {"C": 1}, "yellow-1": {"Y": 1}}
    kept = b.sklad_decide("ozon", "Ozon", "Озон1", postings, {}, {"Y"}, log, lambda text: True, True,
                          drop=["red-1"])
    assert set(kept) == {"go-1", "yellow-1"}          # жёлтый покрашен владельцем — склад не переспрашивает
    path, body = sklad["calls"][0]
    assert path == "/api/v1/demand" and body["drop"] == [{"market": "ozon", "posting": "red-1"}]
    assert "едет 2 шт." in log.text() and "не едет 2 шт." in log.text()


def test_sklad_down_falls_back_to_report_colors(sklad, isolated, log, monkeypatch):
    monkeypatch.setitem(b.SKLAD, "url", "https://sklad-down.test")
    isolated.add(responses.POST, re.compile(r"https://sklad-down\.test/.*"), body=requests.ConnectionError("down"))
    postings = {"1-1": {"A": 1}}
    assert b.sklad_decide("ozon", "Ozon", "Озон1", postings, {}, set(), log, lambda text: True, True) == postings
    assert "Склад не ответил" in log.text()


def test_check_mode_asks_sklad_but_sends_no_batch(sklad, ozon, tmp_path, log, progress):
    ozon.cabinets["Z"]["1-1"] = posting(b.READY, ("Zerkalo_50x130", 1))
    report = ozon_report(tmp_path / "r.xlsx", [("1-1", t(8), "Zerkalo_50x130", 1)])
    b.process(str(report), None, CABINETS, 14, str(tmp_path / "out"), log, progress, download=False)
    assert [path for path, _ in sklad["calls"]] == ["/api/v1/demand"]
    assert set(ozon.paths()) == {"/v4/posting/fbs/list"}


def test_batch_goes_to_sklad_with_codes_and_channel(sklad, ozon, tmp_path, log, progress):
    ozon.cabinets["Z"]["1-1"] = posting(b.READY, ("Zerkalo_50x130", 1))
    report = ozon_report(tmp_path / "r.xlsx", [("1-1", t(8), "Zerkalo_50x130", 1)])
    b.process(str(report), None, CABINETS, 14, str(tmp_path / "out"), log, progress, footer="Озон 3")
    path, body = sklad["calls"][-1]
    assert path == "/api/v1/batches"
    assert (body["market"], body["channel"], body["truck"]) == ("ozon", "Ozon", "Озон3")
    assert body["items"] == [{"posting": "1-1", "origin": "1-1", "article": "Zerkalo_50x130",
                              "production": b.KOMINTERNA, "codes": ["%1-1%"],
                              "ordered_at": "2026-11-09T08:00:00"}]


def test_batch_not_sent_is_saved_and_resent(sklad, ozon, tmp_path, log, progress):
    ozon.cabinets["Z"]["1-1"] = posting(b.READY, ("Zerkalo_50x130", 1))
    report = ozon_report(tmp_path / "r.xlsx", [("1-1", t(8), "Zerkalo_50x130", 1)])
    sklad["fail_batches"] = True
    b.process(str(report), None, CABINETS, 14, str(tmp_path / "out"), log, progress)
    assert "НЕ ушла на склад" in log.text()
    assert len(json.loads(b.SKLAD_PENDING.read_text(encoding="utf-8"))) == 1
    sklad["fail_batches"] = False
    b.sklad_flush(log)
    assert not b.SKLAD_PENDING.exists()          # очередь пуста — файл убран
    assert "На склад дошло отложенное: 1" in log.text()


@pytest.mark.parametrize("wave, expected", [("Озон утро", ("Ozon", "Озон1")), ("Озон 2", ("Ozon", "Озон2")),
                                            ("озон 3", ("Ozon", "Озон3")), ("КГТ 2", ("КГТ", "")),
                                            ("", ("Ozon", "Озон1"))])
def test_wave_to_sklad_channel(wave, expected):
    assert b.ozon_channel(wave) == expected
