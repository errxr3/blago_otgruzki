"""Окно 2.0 (страница в pywebview): всё, что делают кнопки, — без настоящего окна и без сети."""

import re
import threading
import time
from pathlib import Path

import pytest

import blago_otgruzki as b
from test_ozon_process import day  # noqa: F401 — общий отчёт Ozon с кабинетами-подделками

UI = (Path(b.__file__).parent / b.UI_FILE).read_text(encoding="utf-8")


def events(app, kind=None):
    got = []
    while not app.events.empty():
        ev = app.events.get()
        if kind is None or ev["t"] == kind:
            got.append(ev)
    return got


def wait_idle(app, timeout=20):
    end = time.monotonic() + timeout
    while app.busy and time.monotonic() < end:
        time.sleep(0.02)
    assert not app.busy


def test_page_has_every_button():
    for label in ("Скачать отчёты и сборку", "Добавить СДЭК…", "Не уехало вчера…", "Скачать отчёт и сборку",
                  "Проверить", "Скачать и собрать", "Пересобрать ленты", "Пропуска", "Создать поставки и скачать",
                  "Скачать ШК открытых поставок", "Сделать листы, ШК и акты", "Сохранить настройки",
                  "Проверить ключи", "Проверить токен", "Проверить связь", "Отправить на склад"):
        assert f">{label}<" in UI, label
    # каждая кнопка зовёт действие, которое программа разрешает
    for name in re.findall(r"act\('([a-z_]+)'", UI):
        assert name in b.Api.ACTIONS, name
        assert callable(getattr(b.BlagoApp, name)), name
    for key in re.findall(r'id="s_([a-z_]+)"', UI):
        assert key in b.BlagoApp.SETTINGS_KEYS, key


def test_settings_saved_and_loaded(tmp_path):
    app = b.BlagoApp()
    assert [c["name"] for c in app.cab] == ["Озон 1", "Озон 2", "Озон 3"]
    app.set_cabinet(0, "cid", "111")
    app.set_cabinet(0, "key", "k-1")
    app.set_cabinet(0, "name", "Z")
    for key, value in (("wb_token", "wb-t"), ("pass_phone", "+7999"), ("days", "21"), ("autoship", False),
                       ("root_dir", str(tmp_path))):
        assert app.set(key, value)
    assert not app.set("chto-to", 1)   # чужие поля не принимаются
    app.save_now()
    assert events(app, "status")[-1]["color"] == "ok"
    again = b.BlagoApp()
    assert again.cab[0] == {"name": "Z", "cid": "111", "key": "k-1"}
    assert again.cabinets() == {"Z": {"client_id": "111", "api_key": "k-1"}}
    assert again.s["wb_token"] == "wb-t" and again.s["pass_phone"] == "+7999"
    assert again.s["days"] == 21 and again.s["autoship"] is False and again.s["root_dir"] == str(tmp_path)


def test_morning_and_first_wave_get_own_folders(tmp_path):
    app = b.BlagoApp()
    app.set("root_dir", str(tmp_path))
    app.set("footer", "Озон утро")
    assert app.ozon_sub() == "Озон утро"
    app.set("footer", "Озон 1")
    assert app.ozon_sub() == "Озон 1"
    (tmp_path / "Озон утро").mkdir()
    assert b.find_dir(tmp_path, "Озон 1") is None   # «Озон утро» не подхватывается для «Озон 1»
    assert b.find_dir(tmp_path, "озон  утро") == tmp_path / "Озон утро"
    assert app.day_dir("Озон 1") == tmp_path / "Озон 1"
    links = events(app, "links")[-1]["links"]
    assert links["ozon"].endswith("Озон 1") and links["sborka"].endswith("Сборка")


def test_question_waits_for_answer():
    app = b.BlagoApp()
    got = {}
    worker = threading.Thread(target=lambda: got.update(ok=app.ask("Отправить?")))
    worker.start()
    for _ in range(100):
        asks = events(app, "ask")
        if asks:
            break
        time.sleep(0.01)
    assert asks and asks[0]["text"] == "Отправить?" and worker.is_alive()
    b.Api(app).answer(asks[0]["id"], True)
    worker.join(2)
    assert got == {"ok": True}


def test_api_runs_only_allowed_actions():
    app = b.BlagoApp()
    api = b.Api(app)
    assert api.action("__init__") is False and api.action("save_settings") is False
    assert api.action("start_morning") is True   # без шаблона — предупреждение, а не падение
    alert = events(app, "alert")
    assert alert and "шаблон" in alert[0]["text"]
    # pywebview отдаёт странице только то, что видит в объекте: никаких полей, кроме методов
    assert [k for k in vars(api) if not k.startswith("_")] == []


def test_poll_gives_log_with_colors_and_sklad_line():
    app = b.BlagoApp()
    app.log("✔ готово")
    app.log("=== ПРОВЕРКА ===")
    app.log("✖ сломалось")
    r = app.poll()
    assert [(e["text"], e["tag"]) for e in r["events"]] == [("✔ готово", "ok"), ("ПРОВЕРКА", "head"),
                                                          ("✖ сломалось", "err")]
    assert r["sklad"]["text"].startswith("Склад: не подключён") and r["busy"] is False


def test_ozon_check_from_window(day, tmp_path):  # noqa: F811
    """«Проверить» на вкладке Ozon: отчёт, кабинеты, журнал, кнопки гаснут и загораются."""
    report, _ = day
    app = b.BlagoApp()
    app.cab = [{"name": n, "cid": c["client_id"], "key": c["api_key"]} for n, c in b_cabinets().items()]
    app.set("root_dir", str(tmp_path))
    app.set_excel(str(report))
    assert events(app, "excel")[0]["name"] == report.name
    app.start_ozon("check")
    assert app.busy
    wait_idle(app)
    text = "\n".join(e["text"] for e in events(app, "log"))
    assert "OZON · ПРОВЕРКА" in text and "Ошибка" not in text
    assert not list(Path(tmp_path).rglob("*.pdf"))   # «Проверить» ничего не качает


def b_cabinets():
    from fakes import CABINETS
    return CABINETS
