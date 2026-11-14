"""Настройки в отдельной папке (AppData) и перенос старых файлов из папки программы (1.51)."""

import blago_otgruzki as b


def test_old_settings_moved_once(monkeypatch, tmp_path):
    prog, data = tmp_path / "prog", tmp_path / "data"
    prog.mkdir()
    data.mkdir()
    monkeypatch.setattr(b, "BASE_DIR", prog)
    monkeypatch.setattr(b, "DATA_DIR", data)
    (prog / "cabinets.json").write_text('{"x": 1}', encoding="utf-8")
    (prog / "wb.json").write_text('{"old": 1}', encoding="utf-8")
    (data / "wb.json").write_text('{"new": 1}', encoding="utf-8")   # уже есть — не трогаем
    (prog / "BLAGO.xlsx").write_bytes(b"x")                          # не настройки — остаётся

    assert b.move_old_data() == ["cabinets.json"]
    assert (data / "cabinets.json").read_text(encoding="utf-8") == '{"x": 1}'
    assert not (prog / "cabinets.json").exists()
    assert (data / "wb.json").read_text(encoding="utf-8") == '{"new": 1}'
    assert (prog / "BLAGO.xlsx").exists()
    assert b.move_old_data() == []


def test_same_folder_nothing_moves(monkeypatch, tmp_path):
    monkeypatch.setattr(b, "BASE_DIR", tmp_path)
    monkeypatch.setattr(b, "DATA_DIR", tmp_path)
    (tmp_path / "cabinets.json").write_text("{}", encoding="utf-8")
    assert b.move_old_data() == []


def test_windows_uses_appdata(monkeypatch, tmp_path):
    monkeypatch.setattr(b.sys, "platform", "win32")
    monkeypatch.setenv("APPDATA", str(tmp_path))
    assert b._data_dir() == tmp_path / "BLAGO OTGRUZKI"
    assert (tmp_path / "BLAGO OTGRUZKI").is_dir()
