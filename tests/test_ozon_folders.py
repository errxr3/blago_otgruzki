"""Папки этикеток Ozon: у «Озон утро» и «Озон 1» — свои (1.45)."""

import tkinter as tk

import pytest

import blago_otgruzki as b


def test_morning_and_first_wave_get_own_folders(monkeypatch, tmp_path):
    monkeypatch.setattr(b, "MORNING_CONFIG", tmp_path / "morning.json")
    try:
        app = b.App()
    except tk.TclError:
        pytest.skip("нет дисплея")
    try:
        app.root_dir.set(str(tmp_path))
        app.footer.set("Озон утро")
        assert app.ozon_sub() == "Озон утро"
        app.footer.set("Озон 1")
        assert app.ozon_sub() == "Озон 1"
        (tmp_path / "Озон утро").mkdir()
        # готовая папка «Озон утро» не подхватывается для «Озон 1»
        assert app.find_dir(tmp_path, "Озон 1") is None
        assert app.find_dir(tmp_path, "озон  утро") == tmp_path / "Озон утро"
        assert app.day_dir("Озон 1") == tmp_path / "Озон 1"
    finally:
        app.destroy()


def test_rounded_button_acts_like_ttk(monkeypatch, tmp_path):
    monkeypatch.setattr(b, "MORNING_CONFIG", tmp_path / "morning.json")
    try:
        app = b.App()
    except tk.TclError:
        pytest.skip("нет дисплея")
    try:
        hits = []
        btn = b.RButton(app, text="Жми", command=lambda: hits.append(1), style="Accent.TButton")
        btn.pack()
        app.update()
        btn.invoke()
        btn.configure(state="disabled")
        btn.invoke()
        assert hits == [1] and btn.cget("state") == "disabled" and btn["text"] == "Жми"
        btn.configure(text="Другое", state="normal")
        assert btn.cget("text") == "Другое"
        entry = b.REntry(app, show="•")
        entry.pack()
        entry.configure(show="")
        assert entry.entry.cget("show") == ""
    finally:
        app.destroy()
