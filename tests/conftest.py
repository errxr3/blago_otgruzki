"""Общие настройки тестов BLAGO OTGRUZKI.

ЖЁСТКОЕ ПРАВИЛО: ни одного настоящего запроса к Ozon, WB и складу. Вся сеть перехвачена
библиотекой `responses`: запрос на адрес, которого нет среди подделок, падает с ConnectionError.
Ключи в тестах выдуманные; файлы настроек программы (cabinets.json, wb.json, sklad.json …)
тесты не читают и не пишут — пути уведены во временную папку.
"""

import sys
from pathlib import Path

import pytest
import responses

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import blago_otgruzki as b  # noqa: E402

from fakes import FakeOzon, FakeWB  # noqa: E402


@pytest.fixture(autouse=True)
def isolated(monkeypatch, tmp_path):
    """Сеть закрыта, паузы убраны, склад выключен, служебные файлы — во временной папке."""
    monkeypatch.setattr(b.time, "sleep", lambda seconds: None)
    monkeypatch.setitem(b.SKLAD, "url", "")
    monkeypatch.setitem(b.SKLAD, "key", "")
    state = tmp_path / "_state"
    state.mkdir()
    for name, file in (("STICKERS_FILE", "stickers.json"), ("SKLAD_PENDING", "sklad_pending.json"),
                       ("SKLAD_MORNING", "sklad_morning.json"), ("SKLAD_CONFIG", "sklad.json"),
                       ("CONFIG", "cabinets.json"), ("WB_CONFIG", "wb.json"), ("UI_CONFIG", "ui.json"),
                       ("CDEK_CONFIG", "cdek.json"), ("MORNING_CONFIG", "morning.json")):
        monkeypatch.setattr(b, name, state / file)
    monkeypatch.setattr(b, "LABEL_TASK_WAIT", 1)
    with responses.RequestsMock(assert_all_requests_are_fired=False) as mock:
        yield mock


@pytest.fixture
def ozon(isolated):
    return FakeOzon(isolated)


@pytest.fixture
def wb(isolated):
    return FakeWB(isolated)


@pytest.fixture
def log():
    lines = []

    def write(text=""):
        lines.append(str(text))

    write.lines = lines
    write.text = lambda: "\n".join(lines)
    return write


@pytest.fixture
def progress():
    return lambda done, total: None
