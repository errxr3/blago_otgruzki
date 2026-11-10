"""Страховка: в тестах настоящая сеть закрыта, боевых ключей нет."""

import pytest
import requests

import blago_otgruzki as b


@pytest.mark.parametrize("url", [f"{b.API}/v4/posting/fbs/ship", f"{b.WB_API}/api/v3/supplies",
                                 "https://blagosklad.ru/api/v1/batches"])
def test_unmocked_request_never_leaves_the_machine(url):
    with pytest.raises(requests.ConnectionError):
        requests.post(url, json={}, timeout=1)


def test_tests_do_not_use_real_settings_files(tmp_path):
    for path in (b.CONFIG, b.WB_CONFIG, b.SKLAD_CONFIG, b.SKLAD_PENDING, b.SKLAD_MORNING, b.STICKERS_FILE):
        assert tmp_path in path.parents, path
    assert not b.sklad_on()
