"""Окно открывается и кнопки на месте. Нет дисплея (сервер без экрана) — тест пропускается."""

import tkinter as tk

import pytest

import blago_otgruzki as b


def test_window_builds_with_all_buttons(monkeypatch, tmp_path):
    monkeypatch.setattr(b, "MORNING_CONFIG", tmp_path / "morning.json")
    try:
        app = b.App()
    except tk.TclError:
        pytest.skip("нет дисплея")
    try:
        app.update()
        names = {w.cget("text") for w in app.buttons}
        for button in ("Скачать отчёты и сборку", "Не уехало вчера…", "Проверить", "Скачать и собрать",
                       "Создать поставки и скачать", "Пересобрать ленты", "Пропуска"):
            assert button in names, button
    finally:
        app.destroy()
