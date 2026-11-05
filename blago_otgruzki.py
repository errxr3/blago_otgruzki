#!/usr/bin/env python3
"""
BLAGO OTGRUZKI — этикетки Ozon FBS и поставки Wildberries FBS по отчётам.

Отчёт: .xlsx, колонки «Номер отправления», «Артикул», «Количество».
Заказы всех кабинетов вперемешку, в статусах «Ожидают сборки» и
«Ожидают отгрузки». Правее списка — сводная (артикул, сумма):
  красный артикул — не едет, все его отправления пропускаются;
  жёлтый артикул  — едет часть: в колонке сразу после суммы пишется,
                    сколько штук едет; берутся самые старые заказы.

Что делает кнопка «Скачать и собрать»:
  1. читает отчёт и выбрасывает «красные» отправления;
  2. находит, в каком кабинете Ozon лежит каждый заказ;
  3. заказы из «Ожидают сборки» собирает через API — каждая штука товара
     отдельным грузоместом, чтобы на каждое зеркало была своя этикетка;
  4. скачивает этикетки, один PDF на артикул (имя файла = артикул;
     если артикул есть в двух кабинетах — пачка второго с _X в имени);
  5. раскладывает PDF по производствам:
        Коминтерна   — Zerkalo… и Alum…
        Учительская  — всё остальное (Ambient…)
  6. в каждой папке подписывает артикул на этикетках и склеивает их
     в одну ленту (сортировка Я->А):
        «Коминтерна объединенные.pdf», «Учительская объединенные.pdf».

Установка (один раз):  pip install requests openpyxl pypdf reportlab
Запуск:                python blago_otgruzki.py

Сборка в .exe:
    pip install pyinstaller
    pyinstaller --onefile --windowed --name "BLAGO OTGRUZKI" blago_otgruzki.py

Вкладка Wildberries (кнопка «Создать поставки и скачать»):
  1. читает отчёт(ы) WB — «№ задания», «Артикул продавца»; склад и дату
     берёт из фильтров сводной; красные/жёлтые — как у Ozon (для жёлтых
     старшинство заказа берётся из WB);
  2. раскладывает задания на 4 поставки дня:
        «Поставка от ДД.ММ.ГГГГ МГТ коминтерна», «… МГТ учительская»,
        «… СГТ коминтерна», «… СГТ учительская»;
     если есть вчерашняя неотгруженная поставка того же вида — показывает,
     что в ней лежит, и спрашивает, добавлять ли в неё;
  3. скачивает стикеры 58×40 (без подписей) и лист подбора в формате WB,
     в одинаковом порядке, в папку результата WB/<производство>.
  В доставку поставки не передаются.

Ключи кабинетов сохраняются в cabinets.json, токен WB — в wb.json
рядом с программой. Эти файлы никому не пересылать.
"""
import io
import json
import os
import queue
import subprocess
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import openpyxl
import requests
from pypdf import PdfReader, PdfWriter, Transformation
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

APP_NAME = "BLAGO OTGRUZKI"
APP_VERSION = "1.20"  # при каждом обновлении +0.1, запись в CHANGELOG.md
API = "https://api-seller.ozon.ru"
BATCH = 20                   # максимум отправлений в одном запросе этикеток
READY = "awaiting_deliver"   # статус, в котором Ozon отдаёт этикетку
BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
CONFIG = BASE_DIR / "cabinets.json"

# --- производства
KOMINTERNA, UCHITELSKAYA = "Коминтерна", "Учительская"
PRODUCTIONS = (KOMINTERNA, UCHITELSKAYA)

# --- подпись артикула на этикетке
DEFAULT_TOP_MARGIN_PT = 8      # высота полосы сверху под текст
DEFAULT_FONT_SIZE = 8          # размер шрифта
FONT_NAME = "Helvetica-Bold"
STRIP_WIDTH_MM = 14            # запасной вариант, если страница слишком низкая
MM = 72 / 25.4

STATUS_RU = {
    "awaiting_packaging": "ожидает сборки",
    "awaiting_deliver": "ожидает отгрузки",
    "delivering": "доставляется",
    "delivered": "доставлено",
    "cancelled": "отменено",
    "awaiting_registration": "ожидает регистрации",
    "acceptance_in_progress": "идёт приёмка",
    "arbitration": "арбитраж",
}


# ================================================================ логика

def production_of(article):
    """Zerkalo… (в т.ч. bp_Zerkalo…) и Alum… -> Коминтерна, остальное -> Учительская."""
    a = article.strip().lower()
    if "zerkalo" in a or a.startswith(("alum", "алюм")):
        return KOMINTERNA
    return UCHITELSKAYA


def combined_name(production):
    return f"{production} объединенные.pdf"


def sheet_names(path):
    return openpyxl.load_workbook(path, read_only=True).sheetnames


def cell_colors(cell):
    """RGB заливки и шрифта ячейки (то, что задано явным цветом)."""
    fill, font = getattr(cell, "fill", None), getattr(cell, "font", None)
    found = []
    for kind, color in (("fill", fill.fgColor if fill is not None and fill.fill_type else None),
                        ("font", font.color if font is not None else None)):
        rgb = getattr(color, "rgb", None)
        if isinstance(rgb, str) and len(rgb) == 8:
            found.append((kind, int(rgb[2:4], 16), int(rgb[4:6], 16), int(rgb[6:8], 16)))
    return found


def is_red(cell):
    """Красная заливка или красный шрифт."""
    return any(r >= 180 and g <= 110 and b <= 110 for _, r, g, b in cell_colors(cell))


def is_yellow(cell):
    """Жёлтая или оранжево-жёлтая заливка (стандартные жёлтый и «золотой» из палитры Excel)."""
    return any(kind == "fill" and r >= 200 and g >= 150 and b <= 120 for kind, r, g, b in cell_colors(cell))


def parse_time(value):
    if isinstance(value, datetime):
        return value
    if isinstance(value, str):
        for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
            try:
                return datetime.strptime(value.strip(), fmt)
            except ValueError:
                pass
    return None


class Report:
    """Разобранный отчёт."""

    def __init__(self):
        self.postings = {}  # едут: {номер: {артикул: количество}}
        self.red = []       # не едут, красные: (артикул, номер)
        self.yellow = []    # не едут, не вошли в жёлтое количество: (артикул, номер, причина)
        self.mixed = []     # (номер, причина, [не едущие артикулы], [остальные артикулы])
        self.limits = []    # по жёлтым артикулам: dict(art, total, limit, taken, skipped)
        self.notes = []     # замечания к самому отчёту
        self.times = {}     # {номер: время «Принят в обработку»}

    @property
    def total(self):
        return len(self.postings) + len(self.red) + len(self.yellow)


COL_NUM = ("Номер отправления", "№ задания")
COL_ART = ("Артикул", "Артикул продавца")
COL_QTY = ("Количество", "к")
COL_TIME = ("Принят в обработку",)


def _find_col(header, names, required=True):
    for n in names:
        if n in header:
            return header.index(n)
    if required:
        raise ValueError(f"В первой строке листа нет колонки «{names[0]}»")
    return None


def read_report(path, sheet=None, keep=None, times=None):
    """Читает отчёт. Сам файл программа не меняет.

    Список отправлений — слева: «Номер отправления», «Принят в обработку»,
    «Артикул», «Количество». Правее — сводная: артикул, сумма, а в следующей
    колонке при необходимости число «сколько штук едет».

    КРАСНЫЙ артикул в сводной  -> не едет ни одно отправление с ним.
    ЖЁЛТЫЙ артикул в сводной   -> едет только указанное рядом количество штук;
                                  берутся самые старые заказы, новые остаются.
    Красная строка в самом списке -> не едет это отправление.
    """
    rpt = Report()
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = list(ws.iter_rows())
    header = [str(c.value).strip() if c.value is not None else "" for c in rows[0]] if rows else []
    i_num, i_art = _find_col(header, COL_NUM), _find_col(header, COL_ART)
    i_qty = _find_col(header, COL_QTY, required=False)
    i_time = _find_col(header, COL_TIME, required=False)
    last_raw = max(i_num, i_art)  # правая граница списка отправлений
    while last_raw + 1 < len(header) and header[last_raw + 1]:
        last_raw += 1

    # --- цвета в сводной
    red_articles, limits = set(), {}  # limits: жёлтый артикул -> сколько штук едет (None — число не указано)
    for row in rows:
        for i in range(last_raw + 1, len(row)):
            c = row[i]
            if not isinstance(c.value, str) or not c.value.strip():
                continue
            art = c.value.strip()
            if is_red(c):
                red_articles.add(art)
            elif is_yellow(c):
                val = row[i + 2].value if i + 2 < len(row) else None  # артикул, сумма, «едет»
                try:
                    limits[art] = max(int(val), 0) if val is not None and not isinstance(val, bool) else None
                except (TypeError, ValueError):
                    limits[art] = None
    for art in red_articles:
        limits.pop(art, None)

    # --- список отправлений
    allp = {}
    for order, row in enumerate(rows[1:]):
        if len(row) <= max(i_num, i_art) or not row[i_num].value or not row[i_art].value:
            continue
        num, art = str(row[i_num].value).strip(), str(row[i_art].value).strip()
        if keep and not keep({header[i]: row[i].value for i in range(last_raw + 1) if header[i]}):
            continue  # строка не проходит фильтр (например, другой склад WB)
        try:
            qty = max(int(row[i_qty].value), 1) if i_qty is not None else 1
        except (TypeError, ValueError):
            qty = 1
        p = allp.setdefault(num, {"arts": {}, "time": None, "order": order, "red_row": False})
        p["arts"][art] = p["arts"].get(art, 0) + qty
        if i_time is not None and p["time"] is None:
            p["time"] = parse_time(row[i_time].value)
        if p["time"] is None and times:
            p["time"] = times.get(num)
        if any(is_red(c) for c in row[:last_raw + 1]):
            p["red_row"] = True

    used = {a for p in allp.values() for a in p["arts"]}
    limits = {a: v for a, v in limits.items() if a in used}
    for art, v in limits.items():
        if v is None:
            rpt.notes.append(f"Жёлтый артикул {art}: рядом со сводной не указано, сколько штук едет. "
                             f"Пока числа нет — НЕ едет ни одна штука этого артикула.")
    if limits and i_time is None and not times:
        rpt.notes.append("В отчёте нет колонки «Принят в обработку»: старые и новые заказы "
                         "для жёлтых артикулов определяются по порядку строк (сверху — старые).")

    # --- красные
    rest = {}
    for num, p in allp.items():
        arts = p["arts"]
        red = [a for a in arts if a in red_articles]
        if red or p["red_row"]:
            rpt.red.append((red[0] if red else next(iter(arts)), num))
            other = [a for a in arts if a not in red_articles]
            if red and other and not p["red_row"]:
                rpt.mixed.append((num, "красный, не едет", red, other))
        else:
            rest[num] = p

    # --- жёлтые: от старых к новым, пока не наберётся разрешённое количество
    left = {a: (v or 0) for a, v in limits.items()}
    stat = {a: {"art": a, "limit": v, "total": 0, "taken": 0, "skipped": []} for a, v in limits.items()}
    for num, p in sorted(rest.items(), key=lambda kv: (kv[1]["time"] or datetime.max, kv[1]["order"])):
        arts = p["arts"]
        mine = [a for a in arts if a in limits]
        for a in mine:
            stat[a]["total"] += arts[a]
        blocked = [a for a in mine if arts[a] > left[a]]
        if blocked:
            when = p["time"].strftime("%d.%m %H:%M") if p["time"] else ""
            for a in blocked:
                stat[a]["skipped"].append((num, when, arts[a], left[a]))
            a = blocked[0]
            if limits[a] is None:
                why = "жёлтый — не указано, сколько штук едет"
            elif left[a] == 0:
                why = f"жёлтый — новый заказ, разрешённые {limits[a]} шт. уже набраны более старыми"
            else:
                why = (f"жёлтый — в одном отправлении {arts[a]} шт., а жёлтым выбрано {limits[a]} "
                       f"(свободно {left[a]}); это один заказ, частично собрать нельзя")
            rpt.yellow.append((a, num, why))
            other = [a for a in arts if a not in blocked]
            if other:
                rpt.mixed.append((num, "жёлтый, не вошёл в количество", blocked, other))
        else:
            for a in mine:
                left[a] -= arts[a]
                stat[a]["taken"] += arts[a]
            rpt.postings[num] = arts
    # отправления возвращаем в порядке отчёта
    rpt.postings = {n: rpt.postings[n] for n in allp if n in rpt.postings}
    rpt.limits = [stat[a] for a in sorted(stat)]
    rpt.times = {n: p["time"] for n, p in allp.items() if p["time"]}
    return rpt


WORKERS = 6   # сколько запросов к Ozon идёт одновременно (лимит Ozon — 50 в секунду на кабинет)
RETRY_WAITS = (1, 2, 4, 8, 15, 30)   # паузы при ответе 429 «слишком много запросов», сек.


NET_WAITS = (3, 8, 15)   # повторы, если Ozon не ответил вовремя или оборвал связь


def ozon_post(url, net_retry=True, tries429=None, **kw):
    """POST в Ozon с повтором, если Ozon ответил 429 (превышен лимит запросов в секунду),
    и — для чтения — если Ozon не ответил за отведённое время или оборвал связь.
    Сборку (net_retry=False) при обрыве не повторяем: она могла пройти, повтор опасен."""
    waits = list(NET_WAITS) if net_retry else []
    for wait in (RETRY_WAITS if tries429 is None else RETRY_WAITS[:tries429]):
        try:
            r = requests.post(url, **kw)
        except (requests.Timeout, requests.ConnectionError):
            if not waits:
                raise
            time.sleep(waits.pop(0))
            continue
        if r.status_code != 429:
            return r
        time.sleep(wait)
    if tries429 is not None:
        return r if "r" in locals() else requests.post(url, **kw)
    return requests.post(url, **kw)


def headers_for(cfg):
    return {"Client-Id": str(cfg["client_id"]).strip(), "Api-Key": cfg["api_key"].strip(),
            "Content-Type": "application/json"}


OBSOLETE_WORDS = ("obsolete", "deprecated", "disabled", "no longer", "not supported", "not implemented",
                  "устарел", "отключ", "больше не поддерж", "не поддерживается")


def obsolete_error(r, market, msg):
    """Понятный текст, если маркетплейс отключил метод, которым пользуется программа."""
    text = str(msg).lower()
    if r.status_code in (405, 410, 501) or any(w in text for w in OBSOLETE_WORDS):
        path = str(getattr(r, "url", "") or "").split("?")[0]
        path = path.split(".ru", 1)[-1] if ".ru" in path else path
        return (f"{market} ОТКЛЮЧИЛ МЕТОД {path or ''} — программу нужно обновить. "
                f"Пришли скриншот журнала. ({str(msg)[:150]})")
    return None


def api_error(r):
    try:
        msg = r.json().get("message") or r.text
    except ValueError:
        msg = r.text
    obs = obsolete_error(r, "OZON", msg)
    if obs:
        return obs
    hint = {401: "неверный Client-Id или Api-Key",
            403: "ключ неверный или у него нет прав (нужна роль Posting FBS)",
            404: "неверный Client-Id или Api-Key"}.get(r.status_code, "")
    return f"ошибка {r.status_code}" + (f" — {hint}" if hint else "") + f" ({str(msg)[:200]})"


def list_postings(headers, days):
    """FBS-отправления кабинета за последние дни:
    {номер: {"status": ..., "products": [{"offer_id", "sku", "quantity"}], "barcodes": [...]}}.

    С октября 2026 Ozon перекрыл /v3/posting/fbs/list (лимит 0 — всегда 429),
    поэтому используется /v4/posting/fbs/list: до 100 отправлений за запрос,
    следующая страница — по cursor вместо offset.
    """
    now = datetime.now(timezone.utc)
    fmt = "%Y-%m-%dT%H:%M:%SZ"
    flt = {"since": (now - timedelta(days=days)).strftime(fmt),
           "to": (now + timedelta(days=1)).strftime(fmt)}
    result, cursor, extras = {}, "", {"with": {"barcodes": True}}
    for _ in range(1000):  # защита от бесконечного цикла
        body = {"cursor": cursor, "filter": flt, "limit": 100, "sort_dir": "asc", **extras}
        r = ozon_post(f"{API}/v4/posting/fbs/list", headers=headers, json=body, timeout=30)
        if r.status_code == 400 and extras:
            extras = {}  # если Ozon не принял запрос штрихкодов — работаем без них
            continue
        if r.status_code != 200:
            raise RuntimeError(api_error(r))
        data = r.json()
        data = data.get("result", data) if isinstance(data.get("result"), dict) else data
        batch = data.get("postings") or []
        for p in batch:
            bc = p.get("barcodes") or {}
            result[p["posting_number"]] = {
                "status": p.get("status", ""),
                "products": [{"offer_id": str(x.get("offer_id", "")), "sku": x.get("sku"),
                              "quantity": int(x.get("quantity") or 1)} for x in p.get("products") or []],
                "barcodes": [str(v) for v in (bc.values() if isinstance(bc, dict) else bc) if v],
                "warehouse": str((p.get("delivery_method") or {}).get("warehouse")
                                 or (p.get("warehouse") or {}).get("name") or ""),
                "time": p.get("in_process_at") or "",
                "wh_id": (p.get("delivery_method") or {}).get("warehouse_id"),
                "dm_id": (p.get("delivery_method") or {}).get("id")}
        cursor = data.get("cursor") or ""
        if not data.get("has_next") or not cursor or not batch:
            return result
    return result


def check_key(headers):
    """Быстрая проверка ключа: один запрос списка заказов на 1 отправление."""
    now = datetime.now(timezone.utc)
    fmt = "%Y-%m-%dT%H:%M:%SZ"
    body = {"cursor": "", "limit": 1, "sort_dir": "asc",
            "filter": {"since": (now - timedelta(days=3)).strftime(fmt), "to": now.strftime(fmt)}}
    r = ozon_post(f"{API}/v4/posting/fbs/list", headers=headers, json=body, timeout=30)
    if r.status_code != 200:
        raise RuntimeError(api_error(r))


def plural(n, one, few, many):
    """1 отправление, 2 отправления, 5 отправлений."""
    n10, n100 = n % 10, n % 100
    word = one if n10 == 1 and n100 != 11 else few if 2 <= n10 <= 4 and not 12 <= n100 <= 14 else many
    return f"{n} {word}"


def units_of(info):
    return sum(p["quantity"] for p in info["products"]) or 1


def ship_posting(headers, num, products):
    """Собирает заказ: каждая штука товара — отдельное грузоместо (своя этикетка).
    Возвращает (номера получившихся отправлений | None, {номер: артикул}, ошибка)."""
    packages = [{"products": [{"product_id": p["sku"], "quantity": 1}]}
                for p in products for _ in range(p["quantity"])]
    try:
        r = ozon_post(f"{API}/v4/posting/fbs/ship", net_retry=False, headers=headers, timeout=60,
                          json={"posting_number": num, "packages": packages,
                                "with": {"additional_data": True}})
    except requests.RequestException as e:
        return None, {}, f"нет связи: {e}"
    if r.status_code != 200:
        return None, {}, api_error(r)
    data = r.json()
    numbers = [str(n) for n in data.get("result") or []]
    arts = {}
    for item in data.get("additional_data") or []:
        prods = item.get("products") or []
        if item.get("posting_number") and prods and prods[0].get("offer_id"):
            arts[str(item["posting_number"])] = str(prods[0]["offer_id"])
    return numbers, arts, ""


# С 2 ноября 2026 Ozon отключил /v2/posting/fbs/package-label. Этикетки теперь — заданием:
#   /v3/posting/fbs/package-label/create {"posting_numbers": [...до 1000]} -> {"tasks": [{"task_id", "task_type"}]}
#   /v2/posting/fbs/package-label/get {"task_id"} -> {"file_url", "status": {...}, "error": {...}}
# task_type: big_label — обычная этикетка, small_label — маленькая 58×40 (её и печатаем).
LABEL_MODE = {"async": True}
LABEL_TYPE = "small_label"
LABEL_TASK_WAIT = 150   # сколько секунд ждём, пока Ozon подготовит файл
LABEL_BATCH = 500       # отправлений в одном задании (Ozon разрешает до 1000)


def _find(obj, key):
    """Первое значение ключа key в ответе Ozon на любой глубине."""
    if isinstance(obj, dict):
        if key in obj:
            return obj[key]
        for v in obj.values():
            found = _find(v, key)
            if found is not None:
                return found
    elif isinstance(obj, list):
        for v in obj:
            found = _find(v, key)
            if found is not None:
                return found
    return None


def _download(url, headers):
    for h in (None, headers):   # ссылка обычно открытая; если нет — с ключами кабинета
        try:
            f = requests.get(url, headers=h, timeout=60)
        except requests.RequestException:
            continue
        if f.status_code == 200 and f.content[:4] == b"%PDF":
            return f.content
    return None


def label_task(headers, numbers):
    """Одно задание на этикетки. Возвращает (pdf | None, {номер: почему не напечатан}, общая ошибка)."""
    r = ozon_post(f"{API}/v3/posting/fbs/package-label/create", headers=headers,
                  json={"posting_numbers": numbers}, timeout=60)
    if r.status_code != 200:
        return None, {}, api_error(r)
    data = r.json()
    tasks = _find(data, "tasks") or []
    task = (next((t for t in tasks if t.get("task_type") == LABEL_TYPE), None)
            or (tasks[0] if tasks else None))
    if not task or not task.get("task_id"):
        return None, {}, f"Ozon не дал задание на этикетки ({r.text[:150]})"
    deadline = time.time() + LABEL_TASK_WAIT
    while time.time() < deadline:
        time.sleep(2)
        g = ozon_post(f"{API}/v2/posting/fbs/package-label/get", headers=headers,
                      json={"task_id": task["task_id"]}, timeout=60)
        if g.status_code != 200:
            return None, {}, api_error(g)
        res = g.json()
        res = res.get("result", res) if isinstance(res.get("result"), dict) else res
        status = res.get("status") or {}
        unprinted = {str(u.get("posting_number")): u.get("message") or "Ozon не напечатал этикетку"
                     for u in (status.get("unprinted_postings") or []) if isinstance(status, dict)}
        err = res.get("error") or {}
        if isinstance(err, dict) and (err.get("code") or err.get("message")):
            return None, unprinted, f"Ozon: {err.get('message') or err.get('code')}"
        if res.get("file_url"):
            pdf = _download(res["file_url"], headers)
            if pdf is None:
                return None, unprinted, "не удалось скачать файл этикеток по ссылке Ozon"
            return pdf, unprinted, ""
        code = str(status.get("code", "") if isinstance(status, dict) else status).lower()
        if code in ("error", "failed", "fail"):
            return None, unprinted, f"Ozon не сделал этикетки ({code})"
        if isinstance(status, dict) and status.get("postings_count") and len(unprinted) >= status["postings_count"]:
            return None, unprinted, ""
    return None, {}, "Ozon долго готовит этикетки — попробуй ещё раз"


def fetch_labels(headers, numbers):
    """PDF с этикетками (для совместимости). Возвращает (bytes | None, текст ошибки)."""
    try:
        pdf, unprinted, err = label_task(headers, numbers)
    except requests.RequestException as e:
        return None, f"нет связи: {e}"
    if pdf is not None and not unprinted:
        return pdf, ""
    return None, err or "; ".join(f"{n}: {m}" for n, m in unprinted.items())


def grab_labels(headers, numbers):
    """Качает этикетки заданиями Ozon. Возвращает ([(номера, pdf)], [(номер, ошибка)])."""
    good, bad = [], []
    for i in range(0, len(numbers), LABEL_BATCH):
        chunk = numbers[i:i + LABEL_BATCH]
        try:
            pdf, unprinted, err = label_task(headers, chunk)
        except requests.RequestException as e:
            pdf, unprinted, err = None, {}, f"нет связи: {e}"
        printed = [n for n in chunk if n not in unprinted]
        if pdf is not None and printed:
            good.append((printed, pdf))
            bad += [(n, unprinted[n]) for n in chunk if n in unprinted]
        else:
            bad += [(n, unprinted.get(n) or err or "этикетка не получена") for n in chunk]
    return good, bad


# ---------------------------------------------------------------- подпись и склейка

def label_from_filename(filename):
    name = os.path.splitext(filename)[0]
    return name.replace("_", " ").replace("-", " ").strip()


def make_top_overlay(page_w, page_h, text, top_margin, font_size):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))
    y_center = page_h - top_margin / 2
    c.setFont(FONT_NAME, font_size)
    c.drawCentredString(page_w / 2, y_center - font_size * 0.32, text)
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def make_strip_overlay(total_width_pt, height_pt, strip_width_pt, text, font_size):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(total_width_pt, height_pt))
    c.saveState()
    c.translate(strip_width_pt / 2, height_pt / 2)
    c.rotate(90)
    size = font_size
    max_len = height_pt - 10
    while c.stringWidth(text, FONT_NAME, size) > max_len and size > 4:
        size -= 0.5
    c.setFont(FONT_NAME, size)
    c.drawCentredString(0, -size / 3, text)
    c.restoreState()
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


TIME_FONT_SIZE = 9      # размер шрифта времени заказа
TIME_STRIP_PT = 12      # нижняя полоса этикетки Ozon, где нет ни QR-кода, ни адреса


def time_text(dt):
    return dt.strftime("%d.%m %H:%M")


_FOOTER_FONT = None


def footer_font():
    """Жирный шрифт с русскими буквами (Helvetica их не умеет)."""
    global _FOOTER_FONT
    if _FOOTER_FONT is None:
        _FOOTER_FONT = FONT_NAME
        windir = os.environ.get("WINDIR", r"C:\Windows")
        for path in (os.path.join(windir, "Fonts", "arialbd.ttf"),
                     "/Library/Fonts/Arial Bold.ttf",
                     "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"):
            if os.path.exists(path):
                try:
                    pdfmetrics.registerFont(TTFont("FooterBold", path))
                    _FOOTER_FONT = "FooterBold"
                    break
                except Exception:  # noqa: BLE001 — нет шрифта, остаётся Helvetica
                    pass
    return _FOOTER_FONT


def make_footer_overlay(page_w, page_h, left_text, right_text):
    """Нижняя полоса этикетки: своя надпись слева, время заказа справа."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))
    font, y = footer_font(), 3
    right_w = c.stringWidth(right_text, font, TIME_FONT_SIZE) if right_text else 0
    if right_text:
        x = page_w - 5
        c.setFillGray(1)
        c.rect(x - right_w - 2, y - 2, right_w + 4, TIME_FONT_SIZE + 2, stroke=0, fill=1)
        c.setFillGray(0)
        c.setFont(font, TIME_FONT_SIZE)
        c.drawRightString(x, y, right_text)
    if left_text:
        x0 = 5
        if not right_text:   # надпись ставится отдельно от времени — оставляем место под время справа
            right_w = c.stringWidth("00.00 00:00", font, TIME_FONT_SIZE)
        room = page_w - x0 - 5 - (right_w + 8)
        size = TIME_FONT_SIZE
        while size > 5 and c.stringWidth(left_text, font, size) > room:
            size -= 0.5
        c.setFillGray(1)   # белым закрываем всю левую часть — старая надпись не просвечивает
        c.rect(x0 - 2, y - 2, room + 8, TIME_FONT_SIZE + 2, stroke=0, fill=1)
        c.setFillGray(0)
        c.setFont(font, size)
        c.drawString(x0, y, left_text)
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def make_time_overlay(page_w, page_h, text):
    return make_footer_overlay(page_w, page_h, "", text)


def stamp_times(pdf_bytes, nums, times, barcodes, footer=""):
    """Ставит время «Принят в обработку» на этикетки одной пачки.

    Какая страница к какому отправлению: сначала по штрихкоду отправления в тексте
    страницы, иначе по порядку (Ozon отдаёт этикетки в порядке запроса).
    Возвращает (страницы, сколько подписано, сколько сопоставлено по штрихкоду).
    """
    pages = list(PdfReader(io.BytesIO(pdf_bytes)).pages)
    by_order = len(pages) == len(nums)
    stamped = matched = 0
    for i, page in enumerate(pages):
        text = "".join(page.extract_text().split())
        hit = [n for n in nums if any("".join(b.split()) in text for b in barcodes.get(n, []))]
        if len(hit) == 1:
            num, matched = hit[0], matched + 1
        elif by_order:
            num = nums[i]
        elif len({times.get(n) for n in nums}) == 1:
            num = nums[0]          # у всех заказов пачки одно время — порядок не важен
        else:
            num = None
        dt = times.get(num) if num else None
        if dt or footer:
            w, h = float(page.mediabox.width), float(page.mediabox.height)
            page.merge_page(make_footer_overlay(w, h, footer, time_text(dt) if dt else ""))
        if dt:
            stamped += 1
    return pages, stamped, matched


def merge_folder(folder, output_path, font_size, top_margin, footer=""):
    """Подписывает артикул (имя файла) на каждой этикетке и склеивает все PDF
    папки в одну ленту, сортировка Я->А. Возвращает (файлов, этикеток)."""
    folder, output_path = Path(folder), Path(output_path)
    files = sorted((p for p in folder.glob("*.pdf")), key=lambda p: p.name.lower(), reverse=True)
    if not files:
        if output_path.exists():
            output_path.unlink()
        return 0, 0

    writer = PdfWriter()
    strip_width_pt = STRIP_WIDTH_MM * MM
    for f in files:
        label = label_from_filename(f.name)
        for page in PdfReader(str(f)).pages:
            w, h = float(page.mediabox.width), float(page.mediabox.height)
            if footer:
                page.merge_page(make_footer_overlay(w, h, footer, ""))
            if h > top_margin + 10:
                page.merge_page(make_top_overlay(w, h, label, top_margin, font_size))
                writer.add_page(page)
            else:
                overlay = make_strip_overlay(w + strip_width_pt, h, strip_width_pt, label, font_size)
                page.add_transformation(Transformation().translate(strip_width_pt, 0))
                overlay.merge_page(page)
                writer.add_page(page)
    pages = len(writer.pages)
    with open(output_path, "wb") as fh:
        writer.write(fh)
    return len(files), pages


def merge_all(out, font_size, top_margin, log, footer=""):
    """Собирает ленты по обеим папкам производств. footer — надпись внизу каждой ШК."""
    footer = " ".join(str(footer).split())
    out = Path(out)
    for prod in PRODUCTIONS:
        target = out / combined_name(prod)
        try:
            files, pages = merge_folder(out / prod, target, font_size, top_margin, footer)
        except PermissionError:
            log(f"✖ {target.name}: файл открыт в другой программе — закрой его и нажми «Пересобрать ленты»")
            continue
        if files:
            log(f"✔ {target.name}: {pages} этикеток, {files} артикулов")
        else:
            log(f"   {prod}: этикеток нет, лента не создана")


# ---------------------------------------------------------------- основной процесс

def log_report(rpt, log, market="Ozon", unit=("отправление", "отправления", "отправлений")):
    """Пишет в журнал разбор отчёта: красные, жёлтые, смешанные, замечания."""
    red, yellow, mixed = rpt.red, rpt.yellow, rpt.mixed
    log(f"В отчёте {unit[2]}: {rpt.total}"
        + (f", красных (не едут): {len(red)}" if red else "")
        + (f", жёлтых сверх количества (не едут): {len(yellow)}" if yellow else ""))
    for note in rpt.notes:
        log("")
        log("⚠ ВНИМАНИЕ: " + note)
    if rpt.limits:
        log("")
        log("● ЖЁЛТЫЕ АРТИКУЛЫ — едет только часть, самые старые заказы:")
        for L in rpt.limits:
            lim = "не указано" if L["limit"] is None else L["limit"]
            log(f"●   {L['art']}:  в отчёте {L['total']} шт.,  разрешено {lim},  "
                f"едет {L['taken']},  не едет {L['total'] - L['taken']}")
            if L["limit"] is None:
                log("        число рядом со сводной не указано — не едет ни одно отправление")
                continue
            whole = [x for x in L["skipped"] if x[3] > 0]    # не помещается целиком
            late = [x for x in L["skipped"] if x[3] == 0]    # количество уже набрано
            for num, when, qty, free in whole:
                log(f"⚠       Отправление {num} ({when}) — {qty} шт. в одном отправлении, "
                    f"а жёлтым выбрано {L['limit']}"
                    + (f" (свободно осталось {free})" if free != L["limit"] else "") + ".")
                log("│       Собрать не можем: это один целый заказ, частично его отправить нельзя.")
                log(f"│       Не едет целиком — все {qty} шт. Чтобы отправить часть, раздели его в кабинете {market}.")
            for num, when, qty, _ in late[:15]:
                log(f"        не едет — новый заказ, {L['limit']} шт. уже набраны более старыми:  "
                    f"{num}   {when}   {qty} шт.")
            if len(late) > 15:
                log(f"        … и ещё {len(late) - 15} новых (после скачивания полный список — в НЕ_ЕДУТ.txt)")
    if mixed:
        log("")
        log(f"⚠ ВНИМАНИЕ: {plural(len(mixed), 'отправление', 'отправления', 'отправлений')} "
            f"НЕ ПОЕДЕТ ЦЕЛИКОМ")
        log("│ Причина: в отправлении несколько товаров, и один из них не едет")
        log("│ (красный или не вошёл в жёлтое количество).")
        log("│ Отправление одно, поэтому не собирается и не едет ни один товар из него.")
        for num, why, r, o in mixed:
            log("│")
            log(f"│ Отправление {num}")
            log(f"│     {why + ':':<31}{', '.join(r)}")
            log(f"│     {'не поедет вместе с ним:':<31}{', '.join(o)}")
        log("│")
        log("│ Что делать: если остальной товар нужно отправить — раздели это отправление")
        log(f"│ в кабинете {market} вручную и собери нужную часть.")
        log("")


PACKING = "awaiting_packaging"
LABEL_RETRIES, LABEL_WAIT = 8, 10  # ждём этикетки только что собранных заказов


def process(excel, sheet, cabinets, days, out, log, progress, download=True,
            font_size=DEFAULT_FONT_SIZE, top_margin=DEFAULT_TOP_MARGIN_PT,
            autoship=False, confirm=lambda text: True, footer=""):
    """download=False — только проверка отчёта: ничего не собирает и не качает."""
    rpt = read_report(excel, sheet)
    postings, red, yellow, mixed = rpt.postings, rpt.red, rpt.yellow, rpt.mixed
    log_report(rpt, log)
    # Склад: что едет по остаткам (покрашенное тобой — как покрашено)
    sk_channel, sk_truck = ozon_channel(footer)
    colored = {a for a, _ in red} | {L["art"] for L in rpt.limits}
    postings = sklad_decide("ozon", sk_channel, sk_truck, postings, rpt.times, colored, log, confirm, download)
    if not postings:
        return

    # 1. В каком кабинете каждое отправление
    snap, owner = {}, {}  # кабинет -> его заказы;  номер -> кабинет
    with ThreadPoolExecutor(max_workers=max(1, len(cabinets))) as pool:
        lists = {name: pool.submit(list_postings, headers_for(cfg), days) for name, cfg in cabinets.items()}
    broken = []
    for name in cabinets:  # разбираем в порядке кабинетов в окне
        try:
            snap[name] = lists[name].result()
        except (RuntimeError, requests.RequestException) as e:
            why = "Ozon не отвечает (долго нет ответа)" if isinstance(e, requests.Timeout) else str(e)
            log(f"✖ {name}: не удалось получить список заказов — {why}")
            broken.append(name)
            continue
        mine = [n for n in postings if n in snap[name] and n not in owner]
        owner.update((n, name) for n in mine)
        st = [snap[name][n]["status"] for n in mine]
        log(f"✔ {name}: из отчёта найдено {len(mine)} — ожидают отгрузки {st.count(READY)}, "
            f"ожидают сборки {st.count(PACKING)}")
    if broken:
        log("")
        log(f"✖ ОСТАНОВЛЕНО: не загрузился кабинет {', '.join(broken)} — ничего не собрано и не скачано.")
        log("│ Подожди минуту и нажми ещё раз.")
        return

    # 2. Что с каждым отправлением делать
    entries, packing, failed, bulky = [], [], [], []  # entries: (номер, артикул, кабинет)
    for num, arts in postings.items():
        first = next(iter(arts))
        if num not in owner:
            failed.append((first, num, "-", f"не найден ни в одном кабинете за {days} дн."))
            continue
        cab = owner[num]
        info = snap[cab][num]
        if info["status"] == READY:
            entries.append((num, first, cab))
            if units_of(info) > 1:
                bulky.append((first, num, cab, units_of(info)))
        elif info["status"] == PACKING:
            packing.append((num, first, cab))
        else:
            st = STATUS_RU.get(info["status"], info["status"])
            failed.append((first, num, cab, f"статус «{st}» — этикетка доступна только в «ожидает отгрузки»"))
    to_ship = packing if autoship else []
    places = sum(units_of(snap[c][n]) for n, _, c in to_ship)
    if to_ship and download:
        log(f"Будет собрано: {plural(len(to_ship), 'заказ', 'заказа', 'заказов')} → "
            f"{plural(places, 'грузоместо', 'грузоместа', 'грузомест')} (каждая штука отдельно)")
    if bulky:
        log("")
        log(f"⚠ ВНИМАНИЕ: {plural(len(bulky), 'отправление', 'отправления', 'отправлений')} "
            f"— НЕСКОЛЬКО ШТУК, НО ТОЛЬКО 1 ШК")
        log("│ Причина: они уже собраны в кабинете одной упаковкой. Разделить собранное нельзя,")
        log("│ поэтому на каждое такое отправление скачается одна этикетка.")
        log("│")
        for art, num, cab, units in bulky[:10]:
            log(f"│ {num}   {art}   [{cab}]   {units} шт. → 1 ШК")
        if len(bulky) > 10:
            log(f"│ … и ещё {len(bulky) - 10}")
        log("")

    if not download:
        # Итог проверки: считаем и «ожидают отгрузки», и «ожидают сборки».
        # В кабинетах при этом ничего не меняется.
        pack_places = sum(units_of(snap[c][n]) for n, _, c in packing)
        total = len(entries) + pack_places
        by_cab, by_prod = defaultdict(int), defaultdict(int)
        for _, a, c in entries:
            by_cab[c] += 1
            by_prod[production_of(a)] += 1
        for n, a, c in packing:
            by_cab[c] += units_of(snap[c][n])
            by_prod[production_of(a)] += units_of(snap[c][n])
        arts = {a for _, a, _ in entries} | {a for _, a, _ in packing}
        log("\n=== ИТОГ ПРОВЕРКИ (в кабинетах ничего не изменено) ===")
        log(f"✔ ВСЕГО ШК БУДЕТ СКАЧАНО: {total}   (артикулов: {len(arts)})")
        log(f"     из «Ожидают отгрузки»:  {len(entries)}")
        log(f"     из «Ожидают сборки»:    {pack_places}   "
            f"({plural(len(packing), 'заказ', 'заказа', 'заказов')}, каждая штука — отдельный ШК)")
        log("   По кабинетам:      " + ",  ".join(f"{c} — {by_cab[c]}" for c in cabinets if c in snap))
        log("   По производствам:  " + ",  ".join(f"{p} — {by_prod[p]}" for p in PRODUCTIONS))
        if packing and not autoship:
            log("")
            log("⚠ ВНИМАНИЕ: ГАЛКА АВТОСБОРКИ СНЯТА")
            log(f"│ Сейчас скачается только {len(entries)} ШК — то, что уже в «Ожидают отгрузки».")
            log(f"│ {plural(len(packing), 'заказ', 'заказа', 'заказов')} из «Ожидают сборки» "
                f"({pack_places} ШК) программа не тронет.")
        if red:
            log("")
            log(f"НЕ ЕДУТ — КРАСНЫЕ: {plural(len(red), 'отправление', 'отправления', 'отправлений')}")
            by_art = defaultdict(int)
            for art, _ in red:
                by_art[art] += 1
            for art, n in sorted(by_art.items()):
                log(f"     {art} — {n}")
        if rpt.limits:
            log("")
            log(f"● ЖЁЛТЫЕ — ЕДЕТ ЧАСТЬ (не едут новые: "
                f"{plural(len(yellow), 'отправление', 'отправления', 'отправлений')})")
            for L in rpt.limits:
                log(f"     {L['art']} — едет {L['taken']} из {L['total']} шт.")
        if mixed:
            log("")
            log(f"⚠ Из не едущих {len(mixed)} — отправления с несколькими товарами (см. предупреждение выше)")
        if failed:
            log("")
            log(f"✖ НЕ ПОЛУЧИТСЯ СКАЧАТЬ: {plural(len(failed), 'отправление', 'отправления', 'отправлений')}")
            for art, num, cab, why in failed[:20]:
                where = f"кабинет {cab}" if cab != "-" else "кабинет не найден"
                log(f"     {num}   {art}   ({where})")
                log(f"          причина: {why}")
            if len(failed) > 20:
                log(f"     … и ещё {len(failed) - 20}")
        return

    if not autoship:
        failed += [(a, n, c, "статус «ожидает сборки» — автосборка выключена") for n, a, c in packing]

    # 3. Сборка заказов «ожидают сборки»
    fresh = {}   # номер нового отправления -> (артикул, кабинет)
    origin = {}  # номер нового отправления -> исходное отправление из отчёта
    if to_ship:
        if not confirm(f"Собрать {plural(len(to_ship), 'заказ', 'заказа', 'заказов')} "
                       f"({plural(places, 'грузоместо', 'грузоместа', 'грузомест')}) в кабинетах Ozon?\n\n"
                       f"Каждая штука станет отдельным отправлением со своей этикеткой. "
                       f"Отменить сборку потом нельзя."):
            log("Отменено — ничего не собрано и не скачано.")
            return
        log("\nСобираю заказы…")
        shipped = defaultdict(set)  # кабинет -> номера заказов (без последней цифры)
        src = {}                    # номер заказа без последней цифры -> исходное отправление
        api_art = {}
        with ThreadPoolExecutor(max_workers=WORKERS) as pool:
            jobs = {pool.submit(ship_posting, headers_for(cabinets[cab]), num, snap[cab][num]["products"]):
                    (num, first, cab) for num, first, cab in to_ship}
            results = {}
            for k, fut in enumerate(as_completed(jobs), 1):
                results[jobs[fut]] = fut.result()
                progress(k, len(to_ship))
        for num, first, cab in to_ship:
            numbers, arts, err = results[(num, first, cab)]
            if numbers is None:
                failed.append((first, num, cab, f"не удалось собрать: {err}"))
                log(f"✖ {num} {first} [{cab}]: {err}")
            else:
                api_art.update(arts)
                for n in numbers or [num]:
                    fresh[n] = (first, cab)
                    origin[n] = num
                shipped[cab].add(num.rsplit("-", 1)[0])
                src[num.rsplit("-", 1)[0]] = num
        # уточняем номера и артикулы получившихся отправлений по свежему списку кабинета
        time.sleep(3)
        with ThreadPoolExecutor(max_workers=max(1, len(shipped))) as pool:
            fresh_lists = {cab: pool.submit(list_postings, headers_for(cabinets[cab]), days) for cab in shipped}
        for cab, orders in shipped.items():
            try:
                new = fresh_lists[cab].result()
            except (RuntimeError, requests.RequestException):
                continue
            for n, info in new.items():
                was = snap[cab].get(n, {}).get("status")
                if (n in fresh or n.rsplit("-", 1)[0] in orders) and info["status"] == READY and was != READY:
                    art = info["products"][0]["offer_id"] if info["products"] else ""
                    fresh[n] = (api_art.get(n) or art or fresh.get(n, ("", cab))[0], cab)
                    origin.setdefault(n, src.get(n.rsplit("-", 1)[0], n))
                    snap[cab].setdefault(n, info)
        for n, (art, cab) in fresh.items():
            entries.append((n, api_art.get(n) or art, cab))
        log(f"✔ Собрано, получилось отправлений: {len(fresh)}")

    # 4. Скачивание. Папки производств очищаются от PDF прошлого запуска,
    #    чтобы старые этикетки не попали в сегодняшнюю ленту.
    out = Path(out)
    for prod in PRODUCTIONS:
        (out / prod).mkdir(parents=True, exist_ok=True)
        for old in (out / prod).glob("*.pdf"):
            old.unlink()

    plan = defaultdict(lambda: defaultdict(list))  # артикул -> кабинет -> номера
    for num, art, cab in entries:
        plan[art][cab].append(num)
    order = list(cabinets)  # порядок кабинетов как в окне
    todo, done = len(entries), 0
    writers, pending = {}, []  # имя файла -> [папка, PdfWriter, штук];  pending: (title, cab, номер, ошибка)

    times = {n: rpt.times[origin.get(n, n)] for n, _, _ in entries if origin.get(n, n) in rpt.times}
    barcodes = {n: info.get("barcodes", []) for cab in snap.values() for n, info in cab.items()}
    stamp = {"done": 0, "matched": 0}

    def add(title, parts):
        w = writers.setdefault(title, [out / production_of(title_art[title]), PdfWriter(), 0])
        for nums, pdf in parts:
            pages, done_, matched_ = stamp_times(pdf, nums, times, barcodes)
            stamp["done"] += done_
            stamp["matched"] += matched_
            for page in pages:
                w[1].add_page(page)
            w[2] += len(nums)

    title_art, title_cab, titles, tasks = {}, {}, [], []
    for art in sorted(plan):
        names = [n for n in order if n in plan[art]]
        for idx, name in enumerate(names):
            # артикул в двух кабинетах: пачка следующего кабинета — с _X в имени
            title = art + ("_" + "X" * idx if idx else "")
            title_art[title], title_cab[title] = art, (name, len(names) > 1)
            titles.append(title)
            tasks.append((title, name, plan[art][name]))

    def fetch_groups(groups):
        """Качает этикетки нескольких пачек параллельно, результат — в порядке пачек."""
        nonlocal done
        with ThreadPoolExecutor(max_workers=WORKERS) as pool:
            futs = [pool.submit(grab_labels, headers_for(cabinets[name]), nums) for _, name, nums in groups]
            for (title, name, nums), fut in zip(groups, futs):
                good, bad = fut.result()
                add(title, good)
                yield title, name, bad
                done += len(nums)
                progress(min(done, todo), todo)

    log(f"… скачиваю этикетки: {todo} шт.")
    for title, name, bad in fetch_groups(tasks):
        pending += [(title, name, n, err) for n, err in bad]
        if done and todo and done % 40 < len(plan[title_art[title]][name]):
            log(f"… скачано {min(done, todo)} из {todo}")

    # только что собранным заказам Ozon делает этикетку не сразу — ждём и повторяем
    for attempt in range(LABEL_RETRIES):
        retry = [x for x in pending if x[2] in fresh]
        if not retry:
            break
        log(f"Жду этикетки от Ozon: осталось {len(retry)} (попытка {attempt + 1} из {LABEL_RETRIES})…")
        time.sleep(LABEL_WAIT)
        pending = [x for x in pending if x[2] not in fresh]
        groups = defaultdict(list)
        for title, name, n, _ in retry:
            groups[(title, name)].append(n)
        for title, name, bad in fetch_groups([(t, n, nums) for (t, n), nums in groups.items()]):
            pending += [(title, name, n, err) for n, err in bad]
    failed += [(title_art[t], n, name, err) for t, name, n, err in pending]

    # Партия → склад: ровно те этикетки, что скачались (по ним грузчики сверяют зеркала)
    not_got = {n for _, n, _, _ in failed}
    sklad_send_batch(today_msk(), "ozon", sk_channel, sk_truck, f"Ozon · {footer.strip() or 'Озон утро'}", [
        {"posting": n, "origin": origin.get(n, n), "article": a, "production": production_of(a),
         "codes": barcodes.get(n, []), "ordered_at": iso(times.get(n))}
        for n, a, _ in entries if n not in not_got], log)

    mixed_note = {n: "; вместе с ним не поехал: " + ", ".join(o) for n, _, _, o in mixed}
    skipped = ([(a, n, "красный") for a, n in red]
               + list(yellow))
    total = 0
    for title in titles:
        if title not in writers:
            log(f"   0  {title}")
            continue
        folder, writer, count = writers[title]
        safe = "".join(c if c.isalnum() or c in " _-." else "_" for c in title)
        writer.write(folder / f"{safe}.pdf")
        total += count
        name, shared = title_cab[title]
        log(f"{count:>4}  {title}   → {folder.name}" + (f"  [{name}]" if shared else ""))

    log(f"Время заказа подписано на {stamp['done']} из {total} ШК"
        + (f" (по штрихкоду сверено {stamp['matched']})" if stamp["matched"] else ""))
    if not times:
        log("⚠ В отчёте нет колонки «Принят в обработку» — время на ШК не подписано")

    for fname, lines in (("НЕ_СКАЧАНО.txt", ["Артикул\tНомер отправления\tКабинет\tПричина"]
                          + ["\t".join(f) for f in failed] if failed else None),
                         ("НЕ_ЕДУТ.txt", ["Артикул\tНомер отправления\tПричина"]
                          + [f"{a}\t{n}\t{why}{mixed_note.get(n, '')}" for a, n, why in skipped]
                          if skipped else None),
                         ("ПРОПУЩЕНО_КРАСНЫЕ.txt", None)):  # старое имя файла — убираем
        if lines:
            (out / fname).write_text("\n".join(lines), encoding="utf-8")
        elif (out / fname).exists():
            (out / fname).unlink()
    log("\n=== ИТОГ ===")
    log(f"Скачано этикеток: {total}")
    if failed:
        log(f"✖ НЕ СКАЧАНО: {len(failed)} — номера и причины в файле НЕ_СКАЧАНО.txt")
    if skipped:
        log(f"   Не едут по отчёту: красных {len(red)}, жёлтых сверх количества {len(yellow)} "
            f"— список в файле НЕ_ЕДУТ.txt")

    # 5. Подпись и склейка по производствам
    log("\nСобираю ленты…")
    merge_all(out, font_size, top_margin, log, footer)
    by_prod = defaultdict(int)
    for folder, _, count in writers.values():
        by_prod[folder.name] += count
    log("")
    log(f"✔ СКАЧАНО ШК: {total}   (" + " · ".join(f"{p} {by_prod[p]}" for p in PRODUCTIONS) + ")")
    if failed:
        log(f"✖ НЕ СКАЧАНО: {len(failed)} — список в НЕ_СКАЧАНО.txt")
    # склады, по которым после скачивания можно сформировать отгрузку: {кабинет: {склад}}
    form = defaultdict(set)
    for num, _, cab in entries:
        info = snap.get(cab, {}).get(num, {})
        for key in ("warehouse", "wh_id", "dm_id"):
            if info.get(key):
                form[cab].add(str(info[key]).lower())
    return dict(form)


# ================================================================ Wildberries

WB_API = "https://marketplace-api.wildberries.ru"
WB_CONFIG = BASE_DIR / "wb.json"
WB_BATCH = 100            # заданий в одном запросе добавления в поставку и стикеров
WB_GAP = 0.21             # WB: не чаще одного запроса в 200 мс
WB_CARGO = {1: "МГТ", 2: "СГТ", 3: "КГТ"}
WB_STATUS_RU = {"new": "новое", "confirm": "на сборке (уже в поставке)",
                "complete": "в доставке", "cancel": "отменено продавцом"}
WB_WBSTATUS_RU = {"canceled_by_client": "покупатель отменил", "declined_by_client": "покупатель отменил",
                  "canceled": "отменено WB", "defect": "брак", "waiting": "в работе у WB",
                  "sorted": "отсортировано WB", "sold": "выкуплено", "ready_for_pickup": "в ПВЗ",
                  "postponed_delivery": "доставка перенесена"}
STICKER_W_MM, STICKER_H_MM = 58, 40

_wb_lock = threading.Lock()
_wb_last = [0.0]


def wb_call(method, path, token, retry429=True, **kw):
    """Запрос к WB: не чаще раза в 200 мс, при 429 — пауза и повтор."""
    headers = {"Authorization": token.strip()}
    if "json" in kw:
        headers["Content-Type"] = "application/json"
    for wait in RETRY_WAITS + (None,):
        with _wb_lock:
            gap = _wb_last[0] + WB_GAP - time.monotonic()
            if gap > 0:
                time.sleep(gap)
            _wb_last[0] = time.monotonic()
        r = requests.request(method, WB_API + path, headers=headers, timeout=60, **kw)
        if r.status_code != 429 or wait is None or not retry429:
            return r
        try:
            wait = max(wait, float(r.headers.get("X-Ratelimit-Retry") or 0))
        except ValueError:
            pass
        time.sleep(wait)
    return r


def wb_error(r):
    try:
        data = r.json()
        msg = data.get("message") or data.get("detail") or data.get("title") or r.text
    except ValueError:
        msg = r.text
    obs = obsolete_error(r, "WB", msg)
    if obs:
        return obs
    hint = {401: "токен неверный или просрочен",
            403: "у токена нет доступа к категории «Маркетплейс»",
            409: "WB не принял задание (статус или габарит не подходит)"}.get(r.status_code, "")
    return f"ошибка {r.status_code}" + (f" — {hint}" if hint else "") + f" ({str(msg)[:200]})"


def wb_ok(r, *codes):
    if r.status_code not in (codes or (200,)):
        raise RuntimeError(wb_error(r))
    return r


def wb_time(value):
    """RFC3339 от WB -> местное время без зоны (как время в отчётах Ozon)."""
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00")).astimezone().replace(tzinfo=None)
    except ValueError:
        return None


def wb_check(token):
    wb_ok(wb_call("GET", "/api/v3/supplies", token, params={"limit": 1, "next": 0}))


def wb_new_orders(token):
    """Новые сборочные задания: {id: задание}."""
    r = wb_ok(wb_call("GET", "/api/v3/orders/new", token))
    return {str(o["id"]): o for o in r.json().get("orders") or []}


def wb_orders(token, days_from, days_to=0):
    """Все задания за период (дни назад): {id: задание}. WB отдаёт максимум 30 дней за раз."""
    now = int(time.time())
    result = {}
    lo_day = days_to
    while lo_day < days_from:
        hi = now - lo_day * 86400
        lo = now - min(lo_day + 30, days_from) * 86400
        nxt = 0
        for _ in range(1000):
            r = wb_ok(wb_call("GET", "/api/v3/orders", token,
                              params={"limit": 1000, "next": nxt, "dateFrom": lo, "dateTo": hi}))
            data = r.json()
            batch = data.get("orders") or []
            for o in batch:
                result[str(o["id"])] = o
            nxt = data.get("next") or 0
            if len(batch) < 1000 or not nxt:
                break
        lo_day += 30
    return result


def wb_statuses(token, ids):
    """{id: (статус продавца, статус WB)}."""
    result = {}
    ids = [int(i) for i in ids]
    for i in range(0, len(ids), 1000):
        r = wb_ok(wb_call("POST", "/api/v3/orders/status", token, json={"orders": ids[i:i + 1000]}))
        for o in r.json().get("orders") or []:
            result[str(o["id"])] = (o.get("supplierStatus", ""), o.get("wbStatus", ""))
    return result


def wb_open_supplies(token):
    """Незакрытые (не переданные в доставку) поставки."""
    result, nxt = [], 0
    for _ in range(1000):
        r = wb_ok(wb_call("GET", "/api/v3/supplies", token, params={"limit": 1000, "next": nxt}))
        data = r.json()
        batch = data.get("supplies") or []
        result += [s for s in batch if not s.get("done")]
        nxt = data.get("next") or 0
        if len(batch) < 1000 or not nxt:
            break
    return result


def wb_supply_orders(token, supply_id):
    r = wb_ok(wb_call("GET", f"/api/marketplace/v3/supplies/{supply_id}/order-ids", token))
    data = r.json()
    return [str(i) for i in (data.get("orderIds") or data.get("orders") or [])]


def wb_create_supply(token, name):
    r = wb_ok(wb_call("POST", "/api/v3/supplies", token, json={"name": name}), 200, 201)
    return r.json()["id"]


def wb_add_orders(token, supply_id, ids):
    """Добавляет задания в поставку по 100. Возвращает [(id, ошибка)] для не добавленных."""
    bad = []
    path = f"/api/marketplace/v3/supplies/{supply_id}/orders"
    for i in range(0, len(ids), WB_BATCH):
        chunk = ids[i:i + WB_BATCH]
        r = wb_call("PATCH", path, token, json={"orders": [int(x) for x in chunk]})
        if r.status_code in (200, 204):
            continue
        if len(chunk) == 1 or r.status_code in (401, 403):
            bad += [(x, wb_error(r)) for x in chunk]
            continue
        for x in chunk:  # одно «плохое» задание валит пачку — добавляем по одному
            r1 = wb_call("PATCH", path, token, json={"orders": [int(x)]})
            if r1.status_code not in (200, 204):
                bad.append((x, wb_error(r1)))
    return bad


def wb_stickers(token, ids):
    """Стикеры PNG 58×40: ({id: стикер}, [(id, ошибка)])."""
    good, bad = {}, []
    for i in range(0, len(ids), WB_BATCH):
        chunk = ids[i:i + WB_BATCH]
        r = wb_call("POST", "/api/v3/orders/stickers", token,
                    params={"type": "png", "width": STICKER_W_MM, "height": STICKER_H_MM},
                    json={"orders": [int(x) for x in chunk]})
        if r.status_code != 200:
            bad += [(x, wb_error(r)) for x in chunk]
            continue
        for st in r.json().get("stickers") or []:
            if st.get("file"):
                good[str(st["orderId"])] = st
        bad += [(x, "WB не отдал стикер") for x in chunk if x not in good]
    return good, bad


def pick_order(article):
    """Порядок почти как в листе подбора WB: по алфавиту, bp_Zerkalo_160x50 — сразу после Zerkalo_160x50."""
    a = article.lower()
    return (a[3:] if a.startswith("bp_") else a), a.startswith("bp_"), a


def sticker_text(st):
    return f"{st.get('partA', '')} {st.get('partB', '')}".strip()


def stickers_pdf(stickers):
    """PDF 58×40 мм, один стикер на страницу, в переданном порядке."""
    import base64
    from reportlab.lib.utils import ImageReader
    buf = io.BytesIO()
    w, h = STICKER_W_MM * MM, STICKER_H_MM * MM
    c = canvas.Canvas(buf, pagesize=(w, h))
    for st in stickers:
        img = ImageReader(io.BytesIO(base64.b64decode(st["file"])))
        c.drawImage(img, 0, 0, width=w, height=h)
        c.showPage()
    c.save()
    return buf.getvalue()


def pick_list_xlsx(path, supply_id, rows, day):
    """Лист подбора в формате WB: rows — [(артикул, стикер)] в порядке стикеров."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Лист подбора"
    left = Alignment(horizontal="left", vertical="center")
    wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.merge_cells("A1:C1")
    ws.merge_cells("A2:I2")
    ws.merge_cells("A4:D4")
    ws.merge_cells("E4:G4")
    ws["A1"] = f"Дата: {day:%d.%m.%Y}"
    ws["A1"].font, ws["A1"].alignment = Font(size=11), left
    ws["A2"] = f"Лист подбора {supply_id}"
    ws["A2"].font, ws["A2"].alignment = Font(size=14, bold=True), left
    ws["A2"].fill = PatternFill("solid", fgColor="FF9A41FE")
    ws["A4"] = f"Количество товаров: {len(rows)}"
    ws["A4"].font, ws["A4"].alignment = Font(size=11), left
    for col, text in (("A", "Артикул продавца"), ("B", "Стикер")):
        c = ws[f"{col}5"]
        c.value, c.font, c.alignment, c.border = text, Font(size=11, bold=True), wrap, box
        c.fill = PatternFill("solid", fgColor="FFE3E4EA")
    for i, (art, sticker) in enumerate(rows, 6):
        a, b = ws.cell(i, 1, art), ws.cell(i, 2, sticker)
        a.font, b.font = Font(size=11), Font(size=11, bold=True)
        for c in (a, b):
            c.alignment, c.border = wrap, box
        ws.row_dimensions[i].height = 68
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 9
    for r, hgt in ((1, 20), (2, 40), (4, 24), (5, 36)):
        ws.row_dimensions[r].height = hgt
    wb.save(path)


B2B_SUFFIX = "юр.лицо"


def is_b2b(order):
    """Заказ юрлица (B2B): WB кладёт его только в отдельную поставку."""
    opts = order.get("options") or {}
    return bool(opts.get("isB2B") or opts.get("isB2b") or order.get("isB2b") or order.get("isB2B"))


def wb_supply_name(day, cargo, prod, b2b=False):
    return f"Поставка от {day:%d.%m.%Y} {cargo} {prod.lower()}" + (f" {B2B_SUFFIX}" if b2b else "")


def wb_parse_name(name):
    """«Поставка от 27.10.2026 МГТ коминтерна [юр.лицо]» -> (дата, габарит, производство, юрлицо) или None."""
    import re
    m = re.match(r"\s*Поставка от (\d\d\.\d\d\.\d{4})\s+(\S+)\s+(\S+)(\s+юр\.?\s*лиц[оа])?\s*$",
                 name or "", re.I)
    if not m:
        return None
    try:
        day = datetime.strptime(m.group(1), "%d.%m.%Y").date()
    except ValueError:
        return None
    prod = next((p for p in PRODUCTIONS if p.lower() == m.group(3).lower()), None)
    return (day, m.group(2).upper(), prod, bool(m.group(4))) if prod else None


def wb_report_filter(path, sheet=None):
    """Фильтры сводной WB («Склад продавца», «Доставить на склад») -> (функция-фильтр, описание).
    В сводной фильтр стоит в ячейке справа от названия поля; «(Все)» — без фильтра."""
    wbk = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wbk[sheet] if sheet else wbk.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=12, values_only=True)]
    wbk.close()
    header = [str(v).strip() if v is not None else "" for v in rows[0]] if rows else []
    first_empty = header.index("") if "" in header else len(header)
    want = {}
    for r_i, row in enumerate(rows):
        for i in range(first_empty, len(row) - 1):
            label = str(row[i]).strip() if row[i] is not None else ""
            if label in ("Склад продавца", "Доставить на склад") and row[i + 1] is not None:
                val = row[i + 1]
                val = val.strftime("%d.%m.%Y") if isinstance(val, datetime) else str(val).strip()
                if val and not val.startswith("(Все") and not val.startswith("(Несколько"):
                    want[label] = val
    if not want:
        return None, "фильтров в сводной нет — берутся все задания"

    def norm(v):
        return v.strftime("%d.%m.%Y") if isinstance(v, datetime) else str(v if v is not None else "").strip()

    def keep(row):
        return all(norm(row.get(k)) == v for k, v in want.items() if k in row)

    return keep, ", ".join(f"{k}: {v}" for k, v in want.items())


def wb_read_reports(files, times):
    """Читает один или несколько отчётов WB. Возвращает [(файл, описание фильтра, Report)]."""
    result = []
    for path in files:
        keep, desc = wb_report_filter(path)
        result.append((path, desc, read_report(path, keep=keep, times=times)))
    return result


WB_SKLAD_CHANNEL = {"МГТ": "WB", "СГТ": "СГТ", "КГТ": "КГТ"}   # габарит поставки → канал склада


def wb_save_supply(token, supply, ids, articles, out, day, log, collect=None):
    """Скачивает стикеры поставки и делает лист подбора. Порядок — как в листе подбора WB:
    по артикулу, внутри артикула — по номеру стикера. Возвращает (сколько ШК, [(id, ошибка)])."""
    if not ids:
        log(f"   {supply['name']} — пустая")
        return 0, []
    good, bad = wb_stickers(token, ids)
    rows = sorted(((articles.get(i, "?"), sticker_text(st), i) for i, st in good.items()),
                  key=lambda x: (pick_order(x[0]), x[1]))
    parsed = wb_parse_name(supply["name"])
    prod = parsed[2] if parsed else production_of(rows[0][0]) if rows else UCHITELSKAYA
    folder = Path(out) / prod
    folder.mkdir(parents=True, exist_ok=True)
    safe = "".join(c if c.isalnum() or c in " _-.()" else "_" for c in supply["name"])
    if rows:
        (folder / f"ШК {safe}.pdf").write_bytes(stickers_pdf([good[i] for _, _, i in rows]))
        pick_list_xlsx(folder / f"Лист подбора {safe}.xlsx", supply["id"], [(a, s) for a, s, _ in rows], day)
    log(f"✔ {supply['name']} — {len(rows)} ШК")
    if collect is not None:  # для партии на склад: какие стикеры скачаны
        cargo = parsed[1] if parsed else "МГТ"
        for art, _, i in rows:
            st = good[i]
            collect.append((cargo, {"posting": str(i), "article": art, "production": prod,
                                    "codes": [str(st.get("barcode") or ""),
                                              f"{st.get('partA', '')}{st.get('partB', '')}", str(i)]}))
    unknown = sum(1 for a, _, _ in rows if a == "?")
    if unknown:
        log(f"⚠    у {unknown} заданий WB не отдал артикул — в листе подбора стоит «?»")
    return len(rows), bad


def wb_articles(token, ids, known):
    """Артикулы заданий: сначала из уже полученных, остальное — из списка заданий за 30, затем 90 дней."""
    arts = {i: known[i]["article"] for i in ids if i in known}
    for lo, hi in ((30, 0), (90, 30)):
        missing = [i for i in ids if i not in arts]
        if not missing:
            break
        found = wb_orders(token, lo, hi)
        arts.update((i, found[i].get("article", "?")) for i in missing if i in found)
    return arts


def process_wb(files, token, out, log, progress, mode="check", confirm=lambda text: True,
               today=None):
    """mode: check — только проверка (в WB ничего не меняется);
             run — создать поставки, добавить задания, скачать стикеры и листы подбора;
             stickers — только скачать стикеры и листы открытых поставок программы."""
    day = today or datetime.now(timezone(timedelta(hours=3))).date()  # дата по Москве

    if mode == "stickers":
        supplies = [s for s in wb_open_supplies(token) if wb_parse_name(s.get("name"))]
        if not supplies:
            log("Открытых поставок «Поставка от …» нет.")
            return
        _wb_download(token, supplies, {}, out, day, log, progress)
        return

    # 1. Новые задания WB: нужны время заказа (для жёлтых) и габарит
    log("… получаю задания WB")
    new = wb_new_orders(token)
    times = {i: wb_time(o.get("createdAt")) for i, o in new.items() if o.get("createdAt")}

    # 2. Отчёты — по строке на файл, жёлтые — по строке на артикул
    reports = wb_read_reports(files, times)
    postings, red, yellow, colored = {}, [], [], set()
    for path, desc, rpt in reports:
        log(f"{Path(path).stem}: едут {len(rpt.postings)}"
            + (f" · красных {len(rpt.red)}" if rpt.red else "")
            + (f" · жёлтых сверх количества {len(rpt.yellow)}" if rpt.yellow else ""))
        for L in rpt.limits:
            lim = "число не указано" if L["limit"] is None else f"едет {L['taken']} из {L['total']}"
            log(f"●   {L['art']} — {lim}")
        for note in rpt.notes:
            log("⚠ " + note)
        postings.update(rpt.postings)
        red += rpt.red
        yellow += rpt.yellow
        colored |= {a for a, _ in rpt.red} | {L["art"] for L in rpt.limits}
    if not postings:
        log("В отчётах нет заданий, которые едут.")
        return

    # 3. Что с каждым заданием
    groups = defaultdict(list)  # (габарит, производство, склад) -> [id]
    failed = []                 # (артикул, id, причина)
    not_new = [n for n in postings if n not in new]
    status = wb_statuses(token, not_new) if not_new else {}
    for num, arts in postings.items():
        art = next(iter(arts))
        o = new.get(num)
        if o is None:
            st = status.get(num)
            if not st:
                why = "не найдено в WB"
            elif st[0] == "new" and st[1] and st[1] != "waiting":
                why = f"{WB_WBSTATUS_RU.get(st[1], st[1])} (статус WB «{st[1]}»)"
            elif st[0] == "new":
                why = "статус «новое», но WB не отдаёт его в новых заданиях"
            else:
                why = f"статус «{WB_STATUS_RU.get(st[0], st[0])}»"
            failed.append((art, num, why))
            continue
        cargo = WB_CARGO.get(o.get("cargoType"), f"габарит {o.get('cargoType')}")
        groups[(cargo, production_of(o.get("article") or art), o.get("warehouseId"), is_b2b(o))].append(num)

    # Склад: что едет по остаткам (покрашенное тобой — как покрашено)
    if sklad_on():
        by_channel = defaultdict(dict)
        for (cargo, _, _), ids in groups.items():
            for i in ids:
                by_channel[WB_SKLAD_CHANNEL.get(cargo, "WB")][i] = postings[i]
        allowed = set()
        for channel, part in by_channel.items():
            got = sklad_decide("wb", channel, "", part, times, colored, log, confirm, mode == "run")
            if got is None:
                return
            allowed |= set(got)
        for key in list(groups):
            groups[key] = [i for i in groups[key] if i in allowed]
            if not groups[key]:
                del groups[key]

    # 4. Поставки: сегодняшние с таким именем берём как есть, вчерашние неотгруженные — спрашиваем
    open_sup = wb_open_supplies(token)
    by_kind = defaultdict(list)
    for s in open_sup:
        p = wb_parse_name(s.get("name"))
        if p:
            by_kind[(p[1], p[2], p[3])].append((p[0], s))
    multi_wh = defaultdict(set)
    for cargo, prod, wh, b2b in groups:
        multi_wh[(cargo, prod, b2b)].add(wh)

    plan = []  # (ключ группы, имя, поставка | None, [id])
    known_arts = dict(new)
    for key in sorted(groups, key=lambda k: (k[3], k[0], PRODUCTIONS.index(k[1]) if k[1] in PRODUCTIONS else 9)):
        cargo, prod, wh, b2b = key
        kind = (cargo, prod, b2b)
        ids = groups[key]
        name = wb_supply_name(day, cargo, prod, b2b)
        if len(multi_wh[kind]) > 1:
            name += f" склад {wh}"
        same = [s for d, s in by_kind.get(kind, []) if d == day and s.get("name") == name]
        older = sorted(((d, s) for d, s in by_kind.get(kind, []) if d < day),
                       key=lambda x: x[0], reverse=True)
        target = same[0] if same else None
        if target is None and len(multi_wh[kind]) == 1:
            for d, s in older:
                inside = wb_supply_orders(token, s["id"])
                arts = wb_articles(token, inside, known_arts)
                cnt = defaultdict(int)
                for i in inside:
                    cnt[arts.get(i, "?")] += 1
                log("")
                log(f"⚠ Не отгружена вчерашняя «{s['name']}» — в ней {len(inside)} шт.:")
                for a, n in sorted(cnt.items()):
                    log(f"│     {a} — {n}")
                if mode != "run":
                    target = None
                    break
                lines = "\n".join(f"   {a} — {n}" for a, n in sorted(cnt.items())) or "   (пустая)"
                if confirm(f"Есть неотгруженная поставка:\n«{s['name']}»\n\nВ ней уже лежит:\n{lines}\n\n"
                           f"Проверь, что эти зеркала есть.\n\n"
                           f"Добавить в неё сегодняшние {len(ids)} заданий ({name[len('Поставка от 00.00.0000 '):]})?\n\n"
                           f"Да — добавить в эту поставку\nНет — создать новую «{name}»"):
                    target = s
                    log("│ → сегодняшние задания добавятся в неё")
                    break
                log("│ → будет новая поставка за сегодня")
        plan.append((key, name, target, ids))

    log("")
    for (cargo, prod, _, _), name, target, ids in plan:
        where = f" (в «{target['name']}»)" if target and target.get("name") != name else ""
        log(f"{'→' if mode != 'run' else '•'} {name} — {plural(len(ids), 'задание', 'задания', 'заданий')}{where}")
    total = sum(len(ids) for *_, ids in plan)
    if failed:
        by_why = defaultdict(list)
        for art, num, why in failed:
            by_why[why].append((art, num))
        for why, items in by_why.items():
            log(f"✖ Не поедут — {why}: {len(items)}")
            if "уже в поставке" not in why:
                for art, num in items[:10]:
                    log(f"│     {num}  {art}")
    if mode != "run":
        log(f"✔ ВСЕГО ШК БУДЕТ СКАЧАНО: {total}")
        return
    if not plan:
        return

    # 5. Создание поставок и добавление заданий
    news = [name for _, name, target, _ in plan if target is None]
    if not confirm(f"Добавить {plural(total, 'задание', 'задания', 'заданий')} в поставки WB"
                   + (f" (новых поставок: {len(news)})" if news else "") + "?\n\n"
                   "Задания перейдут в статус «На сборке». Вернуть их в «Новые» потом нельзя."):
        log("Отменено — в WB ничего не изменено.")
        return
    touched, done = [], 0
    for (cargo, prod, _, _), name, target, ids in plan:
        try:
            if target is None:
                target = {"id": wb_create_supply(token, name), "name": name}
            bad = wb_add_orders(token, target["id"], ids)
        except (RuntimeError, requests.RequestException) as e:
            log(f"✖ {name}: {e}")
            failed += [(next(iter(postings[i])), i, f"не добавлено: {e}") for i in ids]
            continue
        failed += [(next(iter(postings[i])), i, f"не добавлено: {err}") for i, err in bad]
        for i, err in bad:
            log(f"✖ Не добавлено {i}  {next(iter(postings[i]))} — {err}")
        touched.append(target)
        done += len(ids)
        progress(done, total)

    # 6. Стикеры и листы подбора
    sticker_bad = _wb_download(token, touched, known_arts, out, day, log, progress)
    out_dir = Path(out)
    out_dir.mkdir(parents=True, exist_ok=True)
    fails = [(a, n, why) for a, n, why in failed] + [("", n, err) for n, err in sticker_bad]
    skipped = [(a, n, "красный") for a, n in red] + list(yellow)
    for fname, lines in (("НЕ_СКАЧАНО.txt", ["Артикул\tНомер задания\tПричина"]
                          + ["\t".join(f) for f in fails] if fails else None),
                         ("НЕ_ЕДУТ.txt", ["Артикул\tНомер задания\tПричина"]
                          + ["\t".join(x) for x in skipped] if skipped else None)):
        if lines:
            (out_dir / fname).write_text("\n".join(lines), encoding="utf-8")
        elif (out_dir / fname).exists():
            (out_dir / fname).unlink()


def _wb_download(token, supplies, known, out, day, log, progress):
    log("")
    log("… скачиваю стикеры и листы подбора")
    ids = {s["id"]: wb_supply_orders(token, s["id"]) for s in supplies}
    every = [i for v in ids.values() for i in v]
    arts = wb_articles(token, every, known)
    total, bad, collect = 0, [], []
    for k, s in enumerate(supplies, 1):
        n, b = wb_save_supply(token, s, ids[s["id"]], arts, out, day, log, collect)
        total += n
        bad += b
        progress(k, len(supplies))
    for i, err in bad[:20]:
        log(f"✖ стикер {i}: {err}")
    log(f"✔ СКАЧАНО ШК: {total}")
    # Партии → склад: WB (МГТ), СГТ — по скачанным стикерам
    by_channel = defaultdict(list)
    for cargo, item in collect:
        by_channel[WB_SKLAD_CHANNEL.get(cargo, "WB")].append(item)
    for channel, items in by_channel.items():
        sklad_send_batch(day, "wb", channel, "", f"WB · {channel}", items, log)
    return bad


# ================================================================ Утро: отчёты и сборка

MORNING_CONFIG = BASE_DIR / "morning.json"
MSK = timezone(timedelta(hours=3))   # время в отчётах — московское, как в выгрузках Ozon
OZON_STATUSES = (PACKING, READY)     # «ожидают сборки» и «ожидают отгрузки»
# склад с таким словом в названии — отчёт КГТ (в кабинете R он называется «FBS - СГТ Тверь»)
KGT_WORDS = ("кгт", "kgt", "сгт", "sgt", "крупногабар", "сверхгабар")


_TO_CYR = str.maketrans("cskgtaeopxmhby", "сскгтаеорхмнву")  # латинские буквы-двойники -> русские


def is_kgt(warehouse):
    """Ozon иногда пишет «СГТ» латинской C — поэтому сравниваем после замены двойников."""
    w = " ".join((warehouse or "").lower().split()).translate(_TO_CYR)
    return any(k.translate(_TO_CYR) in w for k in KGT_WORDS)
# столбцы «Сборки»: ключ -> как столбец подписан во второй строке листа «Полотна»
SBORKA_COLS = {"ozon": "OZON", "kgt": "КГТ", "cdek": "СДЭК", "wb": "WB Тверь", "sgt": "ТВ СГТ"}
SBORKA_SHEETS = ("Полотна", "Алюминий", "Безрамные")
_LAT = str.maketrans("хХаАоОсСеЕуУрРкКмМтТ", "xXaAoOcCeEyYpPkKmMtT")


def norm_article(article):
    """Артикул к одному виду для сравнения с шаблоном: русские буквы -> латинские, лишние пробелы,
    bp_Zerkalo_50x160 = Zerkalo_50x160_bp, Zerkalo_200x80 = Zerkalo_80x200, регистр не важен."""
    import re
    a = " ".join(str(article).split()).translate(_LAT).replace("Zerkalo.", "Zerkalo_")
    if a.lower().startswith("bp_"):
        a = a[3:] + "_bp"
    m = re.match(r"^(zerkalo)_(\d+)x(\d+)(_.*)?$", a, re.I)
    if m:
        x, y = sorted((int(m.group(2)), int(m.group(3))))
        a = f"Zerkalo_{x}x{y}{m.group(4) or ''}"
    a = "_".join(p for p in a.split("_") if p)
    return a.lower()


def ozon_time(value):
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00")).astimezone(MSK).replace(tzinfo=None)
    except ValueError:
        return None


def write_report(path, sheet, header, rows, qty_col, pivot_title, blocks=None):
    """Отчёт как у тебя: слева список, справа (через 2 пустых столбца) итог по артикулам.
    rows — списки значений под header; qty_col — индекс количества; артикул — колонка «Артикул…».
    blocks — вместо rows: список блоков (Z, R, IZ, Z otgr…), между блоками — линия."""
    from openpyxl.styles import Border, Font, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(header)
    if blocks is not None:
        rows = [r for b in blocks for r in b]
    line = Side(style="thin")
    for b in blocks or [rows]:
        first = ws.max_row + 1
        for r in b:
            ws.append(r)
        if blocks is not None and b:
            for c in range(1, len(header) + 1):
                ws.cell(first, c).border = Border(top=line, bottom=Side())
                last = ws.cell(ws.max_row, c)
                last.border = Border(top=line if ws.max_row == first else Side(), bottom=line)
    for row in ws.iter_rows(min_row=2, max_col=len(header)):
        for c in row:
            if isinstance(c.value, datetime):
                c.number_format = "dd.mm.yyyy h:mm"
    i_art = next(i for i, h in enumerate(header) if h.startswith("Артикул"))
    total = defaultdict(int)
    for r in rows:
        total[r[i_art]] += r[qty_col]
    col = len(header) + 3
    bold = Font(bold=True)
    ws.cell(1, col, "Названия строк").font = bold
    ws.cell(1, col + 1, pivot_title).font = bold
    for k, art in enumerate(sorted(total, key=str.lower), 2):
        ws.cell(k, col, art)
        ws.cell(k, col + 1, total[art])
    end = len(total) + 2
    ws.cell(end, col, "Общий итог").font = bold
    ws.cell(end, col + 1, sum(total.values())).font = bold
    for i, h in enumerate(header, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = max(12, len(h) + 2)
    ws.column_dimensions[openpyxl.utils.get_column_letter(i_art + 1)].width = 34
    ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 34
    ws.column_dimensions[openpyxl.utils.get_column_letter(col + 1)].width = 12
    ws.freeze_panes = "A2"
    wb.save(path)
    return dict(total)


def read_totals(path):
    """Сумма по артикулам из списка заказов (первый лист): колонка «Артикул…» и «к»/«Количество».
    Красный цвет не учитывается — в сборку идёт всё."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    header = [str(c).strip() if c is not None else "" for c in next(rows)]
    i_art = _find_col(header, COL_ART)
    i_qty = _find_col(header, COL_QTY, required=False)
    total = defaultdict(int)
    for r in rows:
        if len(r) <= i_art or r[i_art] is None or not str(r[i_art]).strip():
            continue
        try:
            q = int(r[i_qty]) if i_qty is not None and r[i_qty] is not None else 1
        except (TypeError, ValueError):
            q = 1
        total[str(r[i_art]).strip()] += q
    wb.close()
    return dict(total)


def put_grand_totals(ws, cols):
    """Строка над итогом «полотна(без креплений)//алюминий//безрамные»: общая сумма столбца
    по всем трём листам. Диапазоны берутся из той же итоговой формулы, поэтому шаблон можно менять.
    Если в ячейке уже что-то есть (своя формула) — не трогаем."""
    import re
    for key, c in cols.items():
        for r in range(3, ws.max_row + 1):
            f = ws.cell(r, c).value
            if isinstance(f, str) and f.startswith("=") and '"//"' in f:
                sums = re.findall(r"SUM\([^)]*\)", f, re.I)
                target = ws.cell(r - 1, c)
                if sums and r > 3 and target.value in (None, ""):
                    target.value = "=" + "+".join(sums)
                break


def fill_sborka(template, out_path, day, data, log, base=None):
    """Заполняет «Сборку». data: {ключ столбца: {артикул: количество}}.
    base — уже заполненный файл (тогда меняются только столбцы из data)."""
    wb = openpyxl.load_workbook(base or template)
    head = wb[SBORKA_SHEETS[0]] if SBORKA_SHEETS[0] in wb.sheetnames else wb.worksheets[0]
    cols = {}
    for c in head[2]:
        v = " ".join(str(c.value or "").split()).lower()
        for key, title in SBORKA_COLS.items():
            if v == title.lower():
                cols[key] = c.column
    missing = [SBORKA_COLS[k] for k in data if k not in cols]
    if missing:
        raise ValueError("в шаблоне во 2-й строке листа «Полотна» нет столбцов: " + ", ".join(missing))
    if base is None:
        head.cell(1, 3, datetime.combine(day, datetime.min.time()))
    rows = []   # (лист, строка, нормальный артикул)
    for name in SBORKA_SHEETS:
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        for r in range(3, ws.max_row + 1):
            a = ws.cell(r, 1).value
            if not isinstance(a, str) or not a.strip() or a.startswith("="):
                continue
            rows.append((ws, r, norm_article(a)))
    where = defaultdict(list)
    for ws, r, key in rows:
        where[key].append((ws, r))
    result = {}
    for key, totals in data.items():
        c = cols[key]
        for ws, r, _ in rows:  # чистим столбец: в шаблоне могли остаться старые числа
            cell = ws.cell(r, c)
            if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                cell.value = None
        placed, lost = 0, []
        merged = defaultdict(int)
        names = defaultdict(list)
        for art, q in totals.items():
            merged[norm_article(art)] += q
            names[norm_article(art)].append(art)
        for k, q in merged.items():
            if k in where:
                ws, r = where[k][0]
                ws.cell(r, c).value = q
                placed += q
            else:
                lost.append((", ".join(names[k]), q))
        result[key] = (placed, sum(merged.values()), lost)
    put_grand_totals(head, cols)
    wb.save(out_path)
    lost = [(SBORKA_COLS[k], art, q) for k, (_, _, ls) in result.items() for art, q in ls]
    if lost:
        log("⚠ Нет в шаблоне сборки — добавь строками в шаблон, если нужны:")
        for col, art, q in sorted(lost):
            log(f"│     {art} — {q} шт. ({col.strip()})")
    return result


def wb_warehouses(token):
    try:
        r = wb_call("GET", "/api/v3/warehouses", token)
        if r.status_code == 200:
            return {w.get("id"): w.get("name", "") for w in r.json() or []}
    except requests.RequestException:
        pass
    return {}


def ozon_reports(cabinets, days, out, log, statuses, names, shk=None, lines=None, truck="Озон1", cabs=None):
    """Отчёты Ozon из всех кабинетов. statuses — какие статусы брать (в этом порядке блоков);
    names — {"ozon" | "kgt": имя файла без .xlsx}: какие отчёты делать.
    cabs — если передан, сюда пишется по каждому кабинету: (сколько отправлений | None, ошибка).
    Возвращает ({ключ: {артикул: шт}}, {ключ: путь})."""
    data, files = {}, {}
    log(f"… получаю заказы Ozon из {plural(len(cabinets), 'кабинета', 'кабинетов', 'кабинетов')}")
    with ThreadPoolExecutor(max_workers=max(1, len(cabinets))) as pool:
        lists = {n: pool.submit(list_postings, headers_for(c), days) for n, c in cabinets.items()}
    # блоки как в ручном отчёте: сначала «ожидают сборки» по кабинетам (Z, R, IZ),
    # потом «ожидают отгрузки» в том же порядке кабинетов
    order = list(cabinets)
    blocks = {k: {(st, n): [] for st in statuses for n in order} for k in ("ozon", "kgt")}
    nowh, broken = 0, []
    for name in cabinets:
        try:
            snap = lists[name].result()
        except (RuntimeError, requests.RequestException) as e:
            why = "Ozon не отвечает (долго нет ответа)" if isinstance(e, requests.Timeout) else str(e)
            log(f"✖ {name}: не удалось получить заказы — {why}")
            broken.append(name)
            if cabs is not None:
                cabs[name] = (None, why)
            continue
        if cabs is not None:
            cabs[name] = (sum(1 for i in snap.values() if i["status"] in statuses), "")
        for num, info in snap.items():
            if info["status"] not in statuses:
                continue
            t = ozon_time(info.get("time"))
            kgt = is_kgt(info.get("warehouse"))
            if ("kgt" if kgt else "ozon") not in names:
                continue  # выбран только МГТ или только КГТ — другой вид не трогаем и не пишем
            nowh += not info.get("warehouse")
            for pr in info["products"]:
                blocks["kgt" if kgt else "ozon"][(info["status"], name)].append(
                    [num, t, pr["offer_id"], pr["quantity"]])
                if lines is not None:  # заказы дня для склада
                    lines.append({"market": "ozon", "channel": "КГТ" if kgt else "Ozon",
                                  "truck": "" if kgt else truck, "posting": num, "article": pr["offer_id"],
                                  "qty": pr["quantity"], "production": production_of(pr["offer_id"]),
                                  "ordered_at": iso(t)})
            # ШК: собранное отправление — одна этикетка, несобранное — по этикетке на каждую штуку
            if shk is not None:
                k = "kgt" if kgt else "ozon"
                shk[k] = shk.get(k, 0) + (1 if info["status"] == READY else units_of(info))
    if broken:
        # без одного кабинета отчёт был бы неполным — такой не делаем, чтобы не отгрузить не всё
        raise RuntimeError(f"отчёт НЕ сделан — не загрузился кабинет {', '.join(broken)}. "
                           f"Подожди минуту и нажми ещё раз.")
    if nowh:
        log(f"⚠ У {nowh} заказов Ozon не указан склад — они попали в отчёт Озон, проверь КГТ вручную")
    hdr = ["Номер отправления", "Принят в обработку", "Артикул", "Количество"]
    for key, title in names.items():
        parts = []
        for st in statuses:
            for n in order:
                b = blocks[key][(st, n)]
                b.sort(key=lambda r: (r[1] or datetime.max, r[0]))
                parts.append(b)
        rows = [r for b in parts for r in b]
        path = Path(out) / f"{title}.xlsx"
        data[key] = write_report(path, title.split()[0], hdr, rows, 3, "Сумма по полю Количество", blocks=parts)
        files[key] = path
        per_cab = []
        for n in order:
            k = len({r[0] for st in statuses for r in blocks[key][(st, n)]})
            if k:
                per_cab.append(f"{n} {k}")
        label = "КГТ" if key == "kgt" else "Озон"
        log(f"✔ {label:<5} {' · '.join(per_cab) or 'заказов нет'}"
            + (f"  →  {len({r[0] for r in rows})} отпр., {sum(r[3] for r in rows)} шт." if rows else ""))
    return data, files


def sklad_summary(answer, log):
    """Итог склада по заказам дня одной строкой."""
    if not answer:
        return
    count = defaultdict(int)
    for decision, _ in answer.values():
        count[decision] += 1
    log(f"✔ Склад: едет {count['go']}"
        + (f" · вечером в Озон 3 {count['later']}" if count["later"] else "")
        + (f" · не едет {count['no']}" if count["no"] else "")
        + (f" · вопрос начальнику {count['ask']}" if count["ask"] else ""))


def ozon_wave(cabinets, template, out, days, wave, kgt, log, progress, today=None):
    """«Озон 2», «Озон 3» в течение дня: только «ожидают сборки» (МГТ или КГТ) и «Сборка»
    с одним столбцом. «Ожидают отгрузки» не трогаем — на них ШК уже скачаны утром."""
    day = today or datetime.now(MSK).date()
    stamp = f"{day:%d.%m}"
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    key = "kgt" if kgt else "ozon"
    title = f"{wave}{' КГТ' if kgt and 'кгт' not in wave.lower() else ''} {stamp}"
    shk, lines = {}, []
    cabs = {}
    data, files = ozon_reports(cabinets, days, out, log, (PACKING,), {key: title}, shk,
                               lines=lines, truck=ozon_channel(wave)[1], cabs=cabs)
    if sklad_on():
        sources = [(f"Ozon {n}", *cabs.get(n, (None, "нет ответа"))) for n in cabinets]
        morning_hold(day, lines, [], sources, log, title=title.rsplit(" ", 1)[0])
    progress(1, 2)
    if template and key in data:
        path = out / f"Сборка {title}.xlsx"
        try:
            fill_sborka(template, path, day, {key: data[key]}, log)
            files["sborka"] = path
        except PermissionError:
            log(f"✖ {path.name} открыт в Excel — закрой его и запусти ещё раз")
        except Exception as e:  # noqa: BLE001
            log(f"✖ Сборка: {e}")
    progress(2, 2)
    log(f"✔ ГОТОВО — {title}.xlsx" + (" и сборка" if "sborka" in files else ""))
    return files


def morning(cabinets, token, template, cdek, out, days, log, progress, today=None):
    """Утренние отчёты: Озон, КГТ, ВБ, СГТ (+ СДЭК из файла) и заполненная «Сборка».
    В кабинетах ничего не меняется — только чтение. Возвращает пути к отчётам."""
    day = today or datetime.now(MSK).date()
    stamp = f"{day:%d.%m}"
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    files, data, shk = {}, {}, {}
    lines, markets = [], []   # заказы дня для склада: он решает по остаткам, что едет
    sources = []              # что откуда пришло: (название, отправлений | None, ошибка) — для проверки

    # --- Ozon: все кабинеты сразу
    ozon_failed = False
    if cabinets:
        cabs = {}
        try:
            d, f = ozon_reports(cabinets, days, out, log, OZON_STATUSES,
                                {"ozon": f"Озон {stamp}", "kgt": f"КГТ {stamp}"}, shk, lines=lines, cabs=cabs)
            data.update(d)
            files.update(f)
            markets.append("ozon")  # все кабинеты пришли — можно заменить список Ozon на складе
        except RuntimeError as e:
            ozon_failed = True
            log(f"✖ Озон и КГТ: {e}")
            lines[:] = [x for x in lines if x.get("market") != "ozon"]  # неполный Ozon на склад не шлём
        for name in cabinets:
            count, err = cabs.get(name, (None, "нет ответа"))
            sources.append((f"Ozon {name}", count, err))
    else:
        log("⚠ Ключи Ozon не заполнены — отчёты Озон и КГТ пропущены")
    progress(1, 3)

    # --- WB: новые задания, МГТ и СГТ (СГТ — на сегодня и просроченные)
    if token:
        log("… получаю задания WB")
        try:
            new = wb_new_orders(token)
            names = wb_warehouses(token)
        except (RuntimeError, requests.RequestException) as e:
            log(f"✖ WB: {e}")
            sources.append(("WB", None, str(e)))
            new = None
        if new is not None:
            mgt, sgt, later = [], [], defaultdict(int)
            for o in sorted(new.values(), key=lambda o: o.get("createdAt") or ""):
                wh = names.get(o.get("warehouseId")) or ""
                row = [int(o["id"]), o.get("article", ""), wh or str(o.get("warehouseId") or "")]
                # как в ручном отчёте: СГТ — это склад продавца «… СГТ» (фильтр «Склад продавца»);
                # если WB не отдал название склада — по габариту задания
                is_sgt = is_kgt(wh) if wh else o.get("cargoType") == 2
                if is_sgt:
                    d = o.get("sellerDate") or ""
                    try:
                        due = datetime.strptime(d, "%d.%m.%Y").date()
                    except ValueError:
                        due = None
                    if due and due > day:
                        later[d] += 1
                        continue
                    sgt.append(row + [d or "-", 1])
                else:
                    mgt.append(row + ["-", 1])
                lines.append({"market": "wb", "channel": "СГТ" if is_sgt else "WB", "posting": str(o["id"]),
                              "article": o.get("article", ""), "qty": 1,
                              "production": production_of(o.get("article", "")),
                              "ordered_at": iso(wb_time(o.get("createdAt")))})
            markets.append("wb")
            sources += [("WB", len(mgt), ""), ("СГТ", len(sgt), "")]
            hdr = ["№ задания", "Артикул продавца", "Склад продавца", "Доставить на склад", "к"]
            for key, title, rows in (("wb", "ВБ", mgt), ("sgt", "СГТ", sgt)):
                path = out / f"{title} {stamp}.xlsx"
                data[key] = write_report(path, "Сборочные задания", hdr, rows, 4, "Сумма по полю к")
                files[key] = path
            shk["wb"], shk["sgt"] = len(mgt), len(sgt)
            late = sum(1 for r in sgt if r[3] != f"{day:%d.%m.%Y}")
            log(f"✔ {'ВБ':<5} {plural(len(mgt), 'задание', 'задания', 'заданий')}")
            log(f"✔ {'СГТ':<5} {plural(len(sgt), 'задание', 'задания', 'заданий')}"
                + (f"  (из них просроченных {late})" if late else "")
                + (f"  · на другие дни ещё {sum(later.values())}" if later else ""))
    else:
        log("⚠ Токен WB не заполнен — отчёты ВБ и СГТ пропущены")
    progress(2, 3)

    # --- СДЭК из файла менеджеров
    if cdek:
        try:
            data["cdek"] = read_totals(cdek)
            log(f"✔ {'СДЭК':<5} {sum(data['cdek'].values())} шт.")
            markets.append("cdek")
            sources.append(("СДЭК", len(data["cdek"]), ""))
            for art, q in data["cdek"].items():  # у СДЭК в файле нет номеров — заказ «на артикул»
                lines.append({"market": "cdek", "channel": "СДЭК", "posting": f"СДЭК {day:%d.%m} {art}"[:64],
                              "article": art, "qty": q, "production": production_of(art)})
        except Exception as e:  # noqa: BLE001
            log(f"✖ СДЭК: не удалось прочитать {Path(cdek).name} — {e}")

    # --- Склад: заказы дня НЕ уходят сами — сначала проверка, потом кнопка «Отправить на склад»
    if sklad_on() and (lines or sources):
        morning_hold(day, lines, markets, sources, log, title="Утро")

    # --- Сборка
    if template and data and not ozon_failed:
        path = out / f"Сборка {stamp}.xlsx"
        log("… заполняю сборку")
        try:
            fill_sborka(template, path, day, data, log)
            files["sborka"] = path
        except PermissionError:
            log(f"✖ {path.name} открыт в Excel — закрой его и запусти ещё раз")
        except Exception as e:  # noqa: BLE001
            log(f"✖ Сборка: {e}")
    elif not template:
        log("⚠ Шаблон сборки не выбран — сборка не заполнена")
    progress(3, 3)
    if ozon_failed:
        log("✖ НЕ ГОТОВО: Озон и КГТ не скачались — нажми «Скачать отчёты и сборку» ещё раз")
    else:
        log(f"✔ ГОТОВО — отчёты и Сборка {stamp}" if "sborka" in files else "Готово.")
    return files


def add_cdek(template, cdek, out, log, today=None, send_sklad=True):
    """Дописывает СДЭК в сегодняшнюю «Сборку» (или делает её из шаблона, если её ещё нет)."""
    day = today or datetime.now(MSK).date()
    path = Path(out) / f"Сборка {day:%d.%m}.xlsx"
    data = {"cdek": read_totals(cdek)}
    fill_sborka(template, path, day, data, log, base=path if path.exists() else None)
    log(f"✔ СДЭК {sum(data['cdek'].values())} шт. добавлен в {path.name}")
    if sklad_on() and send_sklad:  # на склад — в ту же очередь «Отправить на склад», что и утро
        lines = [{"market": "cdek", "channel": "СДЭК", "posting": f"СДЭК {day:%d.%m} {art}"[:64],
                  "article": art, "qty": q, "production": production_of(art)} for art, q in data["cdek"].items()]
        morning_hold(day, lines, ["cdek"], [("СДЭК", len(lines), "")], log, title="СДЭК")
    return path


# ---------------------------------------------------------------- СДЭК по заказам (WB DBS)
# Файл заказов: «№ задания», «Артикул продавца», «трэк» (номер заказа СДЭК, статус «Создан»).
# На каждое производство: лист сборки, ШК СДЭК в том же порядке и акт приёма-передачи.
CDEK_CONFIG = BASE_DIR / "cdek.json"
CDEK_API = "https://api.cdek.ru/v2"
CDEK = {"account": "", "secure": ""}
COL_TASK = ("№ задания", "Номер задания", "Номер отправления ИМ")
COL_TRACK = ("трэк", "трек", "Трэк", "Трек", "Номер заказа", "Номер заказа СДЭК")
ACT_MONTHS = ("Января", "Февраля", "Марта", "Апреля", "Мая", "Июня", "Июля", "Августа", "Сентября",
              "Октября", "Ноября", "Декабря")


def _num(value):
    """Номер из ячейки: 10343931083 / '10343931083\xa0' / 1.0343931083E10 -> '10343931083'."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return "".join(str(value).split()).replace("\xa0", "")


def read_cdek_orders(path):
    """Заказы СДЭК из файла: [{task, article, track}] в порядке файла. Нет колонки «трэк» — пусто."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        rows = wb.worksheets[0].iter_rows(values_only=True)
        header = [str(c).strip() if c is not None else "" for c in next(rows, ())]
        i_track = _find_col(header, COL_TRACK, required=False)
        if i_track is None:
            return []
        i_art = _find_col(header, COL_ART)
        i_task = _find_col(header, COL_TASK, required=False)
        out = []
        for r in rows:
            art = str(r[i_art]).strip() if i_art < len(r) and r[i_art] is not None else ""
            track = _num(r[i_track]) if i_track < len(r) else ""
            if not art or not track:
                continue
            task = _num(r[i_task]) if i_task is not None and i_task < len(r) else ""
            out.append({"task": task, "article": art, "track": track})
        return out
    finally:
        wb.close()


def cdek_order(orders):
    """Порядок листа сборки: по артикулу от Я к А без учёта регистра (Zerkalo, bp_Zerkalo, Alum rama, Alum Noga),
    одинаковые артикулы — в порядке файла."""
    return sorted(orders, key=lambda o: o["article"].casefold(), reverse=True)


def cdek_on():
    return bool(CDEK.get("account") and CDEK.get("secure"))


def cdek_token():
    r = requests.post(f"{CDEK_API}/oauth/token", timeout=30, data={
        "grant_type": "client_credentials", "client_id": CDEK["account"], "client_secret": CDEK["secure"]})
    if r.status_code != 200 or "access_token" not in r.text:
        raise RuntimeError(f"СДЭК не пустил по ключам ({r.status_code}): {r.text[:150]}")
    return r.json()["access_token"]


def cdek_barcodes_api(tracks, log, wait=90):
    """ШК мест заказов СДЭК через API: заказ на печать → ждём готовности → PDF. Возвращает [pdf]."""
    token = cdek_token()
    h = {"Authorization": f"Bearer {token}"}
    files = []
    for i in range(0, len(tracks), 100):
        chunk = tracks[i:i + 100]
        r = requests.post(f"{CDEK_API}/print/barcodes", headers=h, timeout=60, json={
            "orders": [{"cdek_number": int(t) if t.isdigit() else t} for t in chunk],
            "copy_count": 1, "format": "A6", "lang": "RUS"})
        if r.status_code not in (200, 202):
            raise RuntimeError(f"СДЭК не принял печать ШК ({r.status_code}): {r.text[:200]}")
        uuid = (r.json().get("entity") or {}).get("uuid")
        if not uuid:
            raise RuntimeError(f"СДЭК не вернул номер печати: {r.text[:200]}")
        deadline = time.time() + wait
        while True:
            info = requests.get(f"{CDEK_API}/print/barcodes/{uuid}", headers=h, timeout=30).json()
            codes = [st.get("code") for st in (info.get("entity") or {}).get("statuses") or []]
            if "READY" in codes:
                break
            if "INVALID" in codes or time.time() > deadline:
                errs = "; ".join(e.get("message", "") for e in info.get("errors") or [])
                raise RuntimeError(f"СДЭК не подготовил ШК ({', '.join(codes) or 'нет ответа'}) {errs}")
            time.sleep(2)
        pdf = requests.get(f"{CDEK_API}/print/barcodes/{uuid}.pdf", headers=h, timeout=60)
        if pdf.status_code != 200 or pdf.content[:4] != b"%PDF":
            raise RuntimeError(f"СДЭК не отдал файл ШК ({pdf.status_code})")
        files.append(pdf.content)
        log(f"… ШК СДЭК: скачано {min(i + 100, len(tracks))} из {len(tracks)}")
    return files


def cdek_pages(pdfs):
    """Страницы ШК по номеру заказа СДЭК (номер — первый текст на этикетке) и код [ITM] с этикетки."""
    import re
    pages, itm = {}, {}
    for data in pdfs:
        for page in PdfReader(io.BytesIO(data)).pages:
            text = page.extract_text() or ""
            m = re.search(r"\b(\d{9,12})\b", text)
            if not m:
                continue
            pages.setdefault(m.group(1), []).append(page)
            i = re.search(r"\[ITM\]\s*(\d+)", text)
            if i:
                itm[m.group(1)] = i.group(1)
    return pages, itm


def write_cdek_sheet(path, orders):
    """Лист сборки: задания сверху вниз (как лягут ШК) и справа — сколько каждого артикула."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Сборочные задания"
    ws.append(["№ задания", "Артикул продавца", "трэк"])
    for o in orders:
        ws.append([int(o["task"]) if o["task"].isdigit() else o["task"], o["article"], o["track"]])
    counts = defaultdict(int)
    for o in orders:
        counts[o["article"]] += 1
    ws["E1"], ws["F1"] = "Артикул", "Количество"
    for i, (art, n) in enumerate(sorted(counts.items(), key=lambda x: x[0].casefold(), reverse=True), start=2):
        ws.cell(row=i, column=5, value=art)
        ws.cell(row=i, column=6, value=n)
    ws.cell(row=len(counts) + 2, column=5, value="Итого")
    ws.cell(row=len(counts) + 2, column=6, value=len(orders))
    for cell in (*ws[1], ws.cell(row=len(counts) + 2, column=5), ws.cell(row=len(counts) + 2, column=6)):
        cell.font = openpyxl.styles.Font(bold=True)
    for col, width in (("A", 14), ("B", 32), ("C", 14), ("E", 32), ("F", 12)):
        ws.column_dimensions[col].width = width
    wb.save(path)


ACT_TEMPLATE = (  # акт приёма-передачи заказов СДЭК — твой образец, строки заполняет программа
    "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"
    "Ji5lULOiTKwozUENjrDzopmocASCF1MYQklUM3NdogIpCI7Ao7w1Uc3WaWqiQuzrTTCz1yWqFWViRWkmKhxh50UzUeEIBC+mMISS"
    "qGbmukQFUhAcgUd5a6LOLBMVYl9vgpm9LlGtKBMrSjNR4Qg7L5qJCkcgeDGFIZRENTPXJSqQguAIPMpbE3XfMlEh9vUmmNnrEtWK"
    "MrGiNBMVjrDzopmocASCF1MYQklUM3NdogIpCI7Ao7w1UeeWiQqxrzfBzF6XqFaUiRWlmahwhJ0XzUSFIxC8mMIQSqKamesSFUhB"
    "cAQe5cWly/oX++u2Y/NdTx1qMu6NKpz6VP/RcB017Y8qvdKzJr1ZJ4zdOtWP9xqQaX+Ivwx8JreoNZfb69wD4wuiH0+7f0FUp2u2"
    "2Ps2pfg9hrwEq8BnfS2VPZXZpKelUuTNpj0tlVXnbNbTUpkGZ/sdljIvyy+r8OlIMR71Mx5rzOf9zNUuPuhnqPbwy36Gagd3DUI1"
    "w31HDM5Prfd79tO8+t6pQugKxxrhQE/oCktVK+3efm/R9IS+6ukJfWXUE4z01GLMhdWjjBXWo2BSq2lmKjU8UfUEU6lVAkhqBQOX"
    "WkWBpVZRMKnVgdFUapVgKjV8cNYTQFIrGLjUKgostYqCSa1OZaZSqwRTqVWCqdSWE7IWA5daRYGlVlEwqdXFnanUKsFUapVgKrVK"
    "AEmtYOBSqyiw1CoKJrVSJRtLrRJMpVYJplKrBJDUCgYutYoCS62iuqSWuyjwaqlmbrYIqxmaTcg1Q7PBuWYIqJZq1sBqqUYAVkuq"
    "VrBqqS4arFqqqwerluoywqolRU9YtdQqLKxaalUYVi3ppTarltqkhicqrFpqk9qsWtJKbVYtdUptVi11Sm1WLemlNquW2qQ2q5ba"
    "pIYPzrBqSSu1WbXUKbVZtdQptVm1pJfarFpqk9qsWmqT2qxaapPackKGVUudUptVS51Sm1VLeqnNqqU2qc2qpTapzaqlNqnNqiWt"
    "1GbVUqfUZtVSp9Rm1ZJearNqqU1qs2qpTWqzaqlNarNqSSu1WbXUKbVZtdQptaZaGt41nkUj2PJ5S/zg9GFDxd2eaz+Y8fI7exYX"
    "AeWBF97RgMjnyQgnnOIROsVjZKSvxbVC+X+c8IKuOGY0Gu3PpvPT/Kji0T/J1/LjSSFf8vU0ab5Xe+iOdEF12l1zr93iPk86p0eK"
    "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"
    "LnhtbFBLBQYAAAAADAAMAAIDAABuNgAAAAA="
)


def _x(text):
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


_WT = r"(<w:t(?: [^>]*)?>)([^<]*)(</w:t>)"


def _set_text(xml, value):
    """Весь текст куска Word — в первый фрагмент (с его оформлением), остальные пустые."""
    import re
    first = [True]

    def put(m):
        text = _x(value) if first[0] else ""
        first[0] = False
        return m.group(1).replace("<w:t>", '<w:t xml:space="preserve">') + text + m.group(3)
    return re.sub(_WT, put, xml)


def _set_numbers(xml, value):
    """«Принято: 16 накладная; мест: 16.» — числа (жирные, крупные) меняются, слова остаются."""
    import re
    state = {"prev_digit": False}

    def put(m):
        digits = m.group(2).strip().isdigit()
        if not digits:
            state["prev_digit"] = False
            return m.group(0)
        text = "" if state["prev_digit"] else str(value)
        state["prev_digit"] = True
        return m.group(1) + text + m.group(3)
    return re.sub(_WT, put, xml)


def write_cdek_act(path, orders, day):
    """Акт приёма-передачи заказов по образцу: дата, строки заказов, «Принято: N накладная; мест: N»."""
    import base64
    import re
    import zipfile
    src = io.BytesIO(base64.b64decode("".join(ACT_TEMPLATE)))
    with zipfile.ZipFile(src) as zin:
        parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    x = parts["word/document.xml"].decode("utf-8")
    a = x.index("<w:tbl>")
    b = x.index("</w:tbl>", a)
    tbl = x[a:b]
    rows = re.findall(r"<w:tr[ >].*?</w:tr>", tbl, flags=re.S)
    tmpl = rows[1]
    # много заказов — таблица продолжается на следующих листах, шапка повторяется на каждом
    head = rows[0].replace("<w:trPr>", "<w:trPr><w:tblHeader/>", 1)
    tbl = tbl.replace(rows[0], head, 1)
    rows[0] = head

    def row(n, o):
        cells = re.findall(r"<w:tc>.*?</w:tc>", tmpl, flags=re.S)
        out = tmpl
        for cell, value in zip(cells, (n, o["track"], o["task"], 1)):
            out = out.replace(cell, _set_text(cell, value), 1)
        return out
    body = "".join(row(n, o) for n, o in enumerate(orders, start=1))
    start = tbl.index(rows[1])
    end = tbl.index(rows[-1]) + len(rows[-1])
    tbl = tbl[:start] + body + tbl[end:]
    # колонка «№» шире, чтобы «10», «100» не переносились
    tbl = re.sub(r'<w:gridCol w:w="\d+"/>', '<w:gridCol w:w="700"/>', tbl, count=1)
    tbl = re.sub(r"<w:tr[ >].*?</w:tr>",
                 lambda r: re.sub(r'<w:tcW w:w="\d+" w:type="dxa"/>', '<w:tcW w:w="700" w:type="dxa"/>',
                                  r.group(0), count=1), tbl, flags=re.S)
    x = x[:a] + tbl + x[b:]

    def fix_para(p):
        text = "".join(t[1] for t in re.findall(_WT, p))
        if "«" in text and "г." in text:
            return _set_text(p, f"от «{day.day}» {ACT_MONTHS[day.month - 1]} {day.year} г.")
        if text.strip().startswith("Принято"):
            return _set_numbers(p, len(orders))
        return p
    x = re.sub(r"<w:p[ >].*?</w:p>", lambda m: fix_para(m.group(0)), x, flags=re.S)
    parts["word/document.xml"] = x.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def cdek_filter(path, orders, log=None):
    """Покраска в файле СДЭК — как в отчётах Ozon: справа сводная «артикул — количество».
    Красный артикул — не едет ни один заказ; жёлтый — едет число из колонки рядом (первые по файлу);
    красная строка в списке — не едет этот заказ. Возвращает заказы, которые едут."""
    rpt = read_report(path)
    going = set(rpt.postings)
    keep = [o for o in orders if not o["task"] or o["task"] in going]
    why = {num: "красный" for _, num in rpt.red}
    why.update({num: reason for _, num, reason in rpt.yellow})
    stopped = [{**o, "why": why.get(o["task"], "не едет")} for o in orders if o["task"] and o["task"] not in going]
    if log:
        for o in stopped:
            log(f"⛔ СДЭК не едет: {o['task']} · {o['article']} — {o['why']}")
        for note in rpt.notes:
            if "Принят в обработку" not in note:  # у СДЭК времени нет — жёлтые идут по порядку файла
                log(f"⚠ {note}")
    return keep, stopped


def write_cdek_stopped(path, stopped, source=None):
    """Общий список «не уехало» для менеджеров: все колонки из файла заказов + производство и почему."""
    header, rows = ["№ задания", "Артикул продавца", "трэк"], {}
    if source:  # строки как в исходном файле — со всеми колонками, что там есть
        wb_in = openpyxl.load_workbook(source, data_only=True, read_only=True)
        try:
            it = wb_in.worksheets[0].iter_rows(values_only=True)
            head = [str(c).strip() if c is not None else "" for c in next(it, ())]
            width = 0
            while width < len(head) and head[width]:
                width += 1
            i_task = _find_col(head, COL_TASK, required=False)
            if width and i_task is not None and i_task < width:
                header = head[:width]
                for r in it:
                    task = _num(r[i_task]) if i_task < len(r) else ""
                    if task:
                        rows.setdefault(task, [_num(v) if isinstance(v, (int, float))
                                               else v.replace("\xa0", " ").strip() if isinstance(v, str) else v
                                               for v in r[:width]])
        finally:
            wb_in.close()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Не уехало"
    ws.append(header + ["Номер WB", "Производство", "Почему"])
    for o in stopped:
        base = rows.get(o["task"]) or [o["task"], o["article"], o["track"]][:len(header)]
        base = [int(v) if isinstance(v, str) and v.isdigit() and len(v) < 16 else v for v in base]
        ws.append(base + [f"WB-DBS-{o['task']}" if o["task"] else "", production_of(o["article"]), o["why"]])
    for cell in ws[1]:
        cell.font = openpyxl.styles.Font(bold=True)
    for i in range(1, ws.max_column + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 18
    ws.column_dimensions[openpyxl.utils.get_column_letter(ws.max_column)].width = 60
    for i, name in enumerate(header + ["Номер WB", "Производство", "Почему"], start=1):
        if "Артикул" in name:
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 32
    wb.save(path)


def make_cdek(orders, out, log, today=None):
    """Лист сборки, ШК и акт на каждое производство + партия СДЭК на склад (сверка штрихкода и зеркала)."""
    day = today or datetime.now(MSK).date()
    out = Path(out)
    pdfs, source = [], "СДЭК"
    if cdek_on():
        log(f"… качаю ШК СДЭК: {len(orders)} заказов")
        try:
            pdfs = cdek_barcodes_api([o["track"] for o in orders], log)
        except (RuntimeError, requests.RequestException, ValueError) as e:
            log(f"✖ ШК из СДЭК не скачались — {e}")
    pages, itm = cdek_pages(pdfs) if pdfs else ({}, {})
    by_prod = defaultdict(list)
    for o in orders:
        by_prod[production_of(o["article"])].append(o)
    missing = []
    for prod in (KOMINTERNA, UCHITELSKAYA):
        group = cdek_order(by_prod.get(prod, []))
        if not group:
            continue
        folder = out / prod
        folder.mkdir(parents=True, exist_ok=True)
        write_cdek_sheet(folder / f"СДЭК сборка {day:%d.%m}.xlsx", group)
        write_cdek_act(folder / f"СДЭК акт {day:%d.%m}.docx", group, day)
        line = f"✔ {prod}: {len(group)} заказов — лист сборки и акт"
        if pages:
            writer, got = PdfWriter(), 0
            for o in group:
                found = pages.get(o["track"]) or []
                if not found:
                    missing.append(o)
                for page in found:
                    writer.add_page(page)
                got += bool(found)
            if got:
                with open(folder / f"СДЭК ШК {day:%d.%m}.pdf", "wb") as f:
                    writer.write(f)
                line += f", ШК {got} шт. (в порядке листа)"
        log(line)
    if pdfs:
        for o in missing:
            log(f"⚠ Нет ШК для заказа СДЭК {o['track']} ({o['article']}) — {source} его не отдал")
    elif not cdek_on():
        log("⚠ ШК СДЭК не сделаны: впиши ключи в «Настройки → СДЭК»")
    if sklad_on():
        lines = [{"market": "cdek", "channel": "СДЭК", "posting": o["track"], "article": o["article"], "qty": 1,
                  "production": production_of(o["article"])} for o in orders]
        morning_hold(day, lines, ["cdek"], [("СДЭК", len(lines), "")], log, title="СДЭК")
        # Партия: грузчик пикает штрихкод СДЭК ([CDK]номер) и нашу этикетку — склад сверяет артикул
        sklad_send_batch(day, "cdek", "СДЭК", "", "СДЭК", [
            {"posting": o["track"], "origin": o["task"] or o["track"], "article": o["article"],
             "production": production_of(o["article"]),
             "codes": [f"[CDK]{o['track']}", o["track"]] + ([f"[ITM]{itm[o['track']]}"] if o["track"] in itm else [])}
            for o in orders], log)
    return out


# ================================================================ Склад BLAGO
# Связь со складом: программа отдаёт заказы дня (склад решает по остаткам, что едет)
# и партии скачанных этикеток (по ним грузчики сверяют зеркала при погрузке).
# Если склад не настроен — программа работает как раньше, без него.

SKLAD_CONFIG = BASE_DIR / "sklad.json"
SKLAD_PENDING = BASE_DIR / "sklad_pending.json"   # не ушло на склад — отправится при следующей связи
SKLAD = {"url": "", "key": ""}                     # заполняется из «Настроек»
# Что последним ушло на склад — строка состояния над журналом: (ok|err, текст)
SKLAD_LAST = {"kind": "", "text": ""}


def sklad_note(kind, text):
    SKLAD_LAST["kind"], SKLAD_LAST["text"] = kind, f"{text} · {datetime.now(MSK):%H:%M}"


def sklad_on():
    return bool(SKLAD["url"].strip() and SKLAD["key"].strip())


def sklad_post(path, payload, timeout=30):
    """POST на склад. Возвращает ответ (dict) или бросает RuntimeError с понятным текстом."""
    url = SKLAD["url"].strip().rstrip("/") + path
    try:
        r = requests.post(url, json=payload, timeout=timeout,
                          headers={"X-API-Key": SKLAD["key"].strip(), "Content-Type": "application/json"})
    except requests.RequestException as e:
        raise RuntimeError(f"нет связи со складом ({e.__class__.__name__})") from e
    if r.status_code == 401:
        raise RuntimeError("склад не принял ключ — проверь ключ в «Настройках»")
    try:
        data = r.json()
    except ValueError:
        raise RuntimeError(f"склад ответил непонятно (код {r.status_code})") from None
    if r.status_code != 200 or not data.get("ok", False):
        raise RuntimeError(f"склад: {data.get('error') or data.get('detail') or r.status_code}")
    return data


def sklad_check():
    url = SKLAD["url"].strip().rstrip("/") + "/api/v1/ping"
    r = requests.get(url, headers={"X-API-Key": SKLAD["key"].strip()}, timeout=15)
    if r.status_code == 401:
        raise RuntimeError("ключ не подходит")
    if r.status_code != 200:
        raise RuntimeError(f"ошибка {r.status_code}")


SKLAD_MORNING = BASE_DIR / "sklad_morning.json"   # заказы ждут проверки и кнопки «Отправить на склад»


def morning_load():
    """Сегодняшние заказы, которые ждут отправки: [{title, day, lines, markets, sources}]."""
    try:
        data = json.loads(SKLAD_MORNING.read_text(encoding="utf-8")) if SKLAD_MORNING.exists() else []
    except ValueError:
        data = []
    if isinstance(data, dict):  # файл прошлой версии: одно утро
        data = [{"title": "Утро", **data}]
    return [x for x in data if x.get("day") == today_msk().isoformat()]


def _morning_save(items):
    try:
        if items:
            SKLAD_MORNING.write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")
        elif SKLAD_MORNING.exists():
            SKLAD_MORNING.unlink()
    except OSError:
        pass


def morning_clear():
    _morning_save([])


def morning_text(items):
    """«Утро: Ozon Z 120 · Ozon R ✖ · WB 160 | Озон 2: Ozon Z 15» — что пришло из каждого кабинета."""
    out = []
    for item in items:
        parts = [f"{name} ✖" if count is None else f"{name} {count}" for name, count, _ in item["sources"]]
        out.append(f"{item['title']}: " + (" · ".join(parts) or "заказов нет"))
    return " | ".join(out)


def morning_problems(items):
    """Что подозрительно: кабинет не загрузился или пришло 0 заказов (0 в волне — обычное дело)."""
    bad = []
    for item in items:
        for name, count, err in item["sources"]:
            if count is None:
                bad.append(f"✖ {item['title']} · {name}: не загрузился ({err[:120]})")
            elif count == 0 and item["title"] == "Утро":
                bad.append(f"⚠ {item['title']} · {name}: 0 заказов — точно так?")
    return bad


def morning_hold(day, lines, markets, sources, log, title="Утро"):
    """Отложить заказы до проверки. На склад они уйдут кнопкой «Отправить на склад».

    Утро заменяет прежнее утро, волна — прежнюю такую же волну (скачал отчёт ещё раз)."""
    item = {"title": title, "day": day.isoformat(), "lines": lines, "markets": markets,
            "sources": [list(x) for x in sources]}
    items = [x for x in morning_load() if x["title"] != title] + [item]
    _morning_save(items)
    log("")
    log(f"⏸ На склад ПОКА НЕ ОТПРАВЛЕНО — {morning_text([item])}")
    for line in morning_problems([item]):
        log(f"│ {line}")
    log("│ Проверь отчёт. Если всё верно — нажми вверху «Отправить на склад».")


def morning_send(log):
    """Отправить на склад всё, что ждёт проверки, по порядку. Не ушедшее остаётся ждать."""
    items = morning_load()
    if not items:
        log("Нечего отправлять: сначала скачай отчёты")
        return False
    left = []
    for item in items:
        lost = [name for name, count, _ in item["sources"] if count is None]
        log(f"→ {item['title']}")
        if lost:
            log(f"⚠ Без {', '.join(lost)}: их заказы на складе не трогаю")
        day = datetime.fromisoformat(item["day"]).date()
        answer = sklad_demand(day, item["lines"], log, replace=item["markets"]) if item["lines"] else {}
        if answer is None:
            left.append(item)
            continue
        sklad_summary(answer, log)
    _morning_save(left)
    return not left


def _pending_load():
    try:
        return json.loads(SKLAD_PENDING.read_text(encoding="utf-8")) if SKLAD_PENDING.exists() else []
    except ValueError:
        return []


def _pending_save(items):
    try:
        if items:
            SKLAD_PENDING.write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")
        elif SKLAD_PENDING.exists():
            SKLAD_PENDING.unlink()
    except OSError:
        pass


def sklad_send_batch(day, market, channel, truck, source, items, log):
    """Партия скачанных этикеток → склад. Не ушла — сохраняется и отправится позже."""
    if not sklad_on() or not items:
        return
    payload = {"day": day.isoformat(), "market": market, "channel": channel, "truck": truck,
               "source": source, "items": items}
    title = ("Озон " + truck.removeprefix("Озон")) if truck else channel
    sklad_flush(log)
    try:
        data = sklad_post("/api/v1/batches", payload, timeout=60)
    except RuntimeError as e:
        _pending_save(_pending_load() + [{"path": "/api/v1/batches", "payload": payload}])
        sklad_note("err", f"партия {title} НЕ ушла на склад")
        log(f"⚠ Партия {title} НЕ ушла на склад — {e}")
        log("│ Она сохранена и отправится сама при следующем запуске (или «Настройки → Отправить на склад»).")
        return
    log(f"✔ Партия → склад: {title} — новых {data.get('added', 0)}, обновлено {data.get('updated', 0)}")
    sklad_note("ok", f"партия {title} ушла на склад: {plural(len(items), 'этикетка', 'этикетки', 'этикеток')}")
    for line in (data.get("kept") or [])[:10]:
        log(f"│   уже в другой партии или погружено: {line}")
    unknown = data.get("unknown") or []
    if unknown:
        log(f"⚠ На складе нет артикулов ({len(unknown)}) — их зеркала сверить не получится:")
        for line in unknown[:10]:
            log(f"│   {line}")


def sklad_flush(log):
    """Отправить то, что не ушло раньше."""
    pending = _pending_load()
    if not pending or not sklad_on():
        return
    left = []
    for job in pending:
        try:
            sklad_post(job["path"], job["payload"], timeout=60)
        except (RuntimeError, KeyError):
            left.append(job)
    _pending_save(left)
    sent = len(pending) - len(left)
    if sent:
        log(f"✔ На склад дошло отложенное: {sent}")
        if not left:
            sklad_note("ok", f"отложенное дошло на склад: {sent}")
    if left:
        log(f"⚠ Ещё не ушло на склад: {len(left)}")


def sklad_demand(day, lines, log, replace=None):
    """Заказы → склад. Ответ: {отправление: (решение, почему)} или None, если склада нет/не ответил."""
    if not sklad_on() or not lines:
        return None
    try:
        data = sklad_post("/api/v1/demand", {"day": day.isoformat(), "lines": lines, "replace": replace or []})
    except RuntimeError as e:
        log(f"⚠ Склад не ответил, что едет ({e}) — решаю как раньше, по цветам отчёта")
        sklad_note("err", "заказы НЕ ушли на склад")
        return None
    count = len({ln["posting"] for ln in lines})
    sklad_note("ok", f"заказы ушли на склад: {plural(count, 'отправление', 'отправления', 'отправлений')}")
    for q in data.get("questions") or []:
        if q.get("status") == "open":
            log(f"● Вопрос начальнику {q['site']}: {q['article']} — нужно {q['need']}, на складе {q['stock']}, "
                f"в работе {q['in_work']}, не хватает {q['short'] - q.get('promised', 0)}")
        elif q.get("status") == "yes":
            log(f"● {q['article']}: начальник обещал {q['promised']} шт. к {q.get('due') or '—'}")
    return {num: (v.get("decision"), v.get("reason", "")) for num, v in (data.get("postings") or {}).items()}


def iso(dt):
    return dt.isoformat() if isinstance(dt, datetime) else (str(dt) if dt else None)


def ozon_channel(wave):
    """«Озон утро»/«Озон 2»/«КГТ 3» → канал и машина склада."""
    import re
    w = " ".join(str(wave or "").lower().split())
    if w.startswith("кгт"):
        return "КГТ", ""
    m = re.search(r"озон\s*(\d)", w)
    n = int(m.group(1)) if m else 1
    return "Ozon", f"Озон{n if 1 <= n <= 4 else 1}"


def sklad_decide(market, channel, truck, postings, times, colored, log, confirm, download):
    """Спросить склад, что из отчёта едет. postings — {номер: {артикул: шт}} после покраски.

    Артикулы, которые ты покрасил сам (красный/жёлтый), склад не переспрашивает — твои цвета главнее.
    Возвращает новый словарь отправлений; None — остановиться (ждём ответа начальника).
    """
    if not sklad_on() or not postings:
        return postings
    lines = [{"market": market, "channel": channel, "truck": truck, "posting": num, "article": art, "qty": q,
              "production": production_of(art), "ordered_at": iso(times.get(num))}
             for num, arts in postings.items() for art, q in arts.items()]
    log("Спрашиваю склад, что едет…")
    answer = sklad_demand(today_msk(), lines, log)
    if answer is None:
        return postings
    keep, no, ask, later = {}, defaultdict(list), defaultdict(list), defaultdict(int)
    for num, arts in postings.items():
        if any(a in colored for a in arts):
            keep[num] = arts
            continue
        decision, why = answer.get(num, ("go", ""))
        if decision == "no":
            no[why or "не едет"].append((num, next(iter(arts))))
        elif decision == "ask":
            ask[next(iter(arts))].append(num)
        elif decision == "later":  # обещано к вечеру — поедет в Озон 3, сейчас не качаем
            later[next(iter(arts))] += 1
        else:
            keep[num] = arts
    log(f"✔ Склад: едет {len(keep)}" + (f" · не едет {sum(map(len, no.values()))}" if no else "")
        + (f" · вечером в Озон 3 {sum(later.values())}" if later else "")
        + (f" · ждёт ответа начальника {sum(map(len, ask.values()))}" if ask else ""))
    if later:
        log("● Доделают к вечеру, поедут в Озон 3: "
            + ", ".join(f"{a} × {n}" for a, n in sorted(later.items())))
    for why, items in no.items():
        arts = defaultdict(int)
        for _, a in items:
            arts[a] += 1
        log(f"✖ Не едет — {why}: " + ", ".join(f"{a} × {n}" for a, n in sorted(arts.items())))
    if ask:
        text = "\n".join(f"   {a} — {len(n)} шт." for a, n in sorted(ask.items()))
        log("● Ждут ответа начальника (на складе мало, артикул в работе):")
        for a, n in sorted(ask.items()):
            log(f"│   {a} — {len(n)} шт.")
        if download and not confirm(
                f"Начальник ещё не ответил по этим зеркалам:\n{text}\n\n"
                f"Да — качать без них (они поедут следующей волной)\n"
                f"Нет — остановиться: дождись ответа начальника на складе и нажми ещё раз"):
            log("Остановлено — жду ответа начальника. Ничего не собрано и не скачано.")
            return None
    return keep


# ================================================================ окно

# Палитра — «slate dark» для рабочих панелей (UI/UX Pro Max: Financial/Operations dashboard),
# фирменный золотой акцент BLAGO оставлен. Контраст основного и серого текста — не ниже 4.5:1.
C = {
    "side": "#0B1120", "bg": "#0F172A", "card": "#151F33", "field": "#0B1220", "line": "#24314A",
    "text": "#F1F5F9", "muted": "#94A3B8", "dim": "#64748B", "accent": "#F2B53B", "accent_hi": "#FFC85C",
    "on_accent": "#1A1405", "btn": "#23304A", "btn_hi": "#2E3D5C", "nav_hi": "#141D30",
    "ok": "#22C55E", "err": "#F87171", "warn": "#FB923C",
}
PAGES = ("Сборка", "Ozon", "Wildberries", "СДЭК", "Настройки")
FONT = "Segoe UI"
MONTHS = ("Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
          "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь")
MONTHS_GEN = ("января", "февраля", "марта", "апреля", "мая", "июня",
              "июля", "августа", "сентября", "октября", "ноября", "декабря")
WEEKDAYS = ("понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье")
WAVES = ("Озон 2", "Озон 3", "КГТ 2", "КГТ 3")
FOOTERS = ("Озон утро", "Озон 1", "Озон 2", "Озон 3", "КГТ 1", "КГТ 2", "КГТ 3")


def ozon_report_summary(path, sheet):
    """Короткая строка про отчёт Ozon: сколько едет, по производствам, красные/жёлтые."""
    rpt = read_report(path, sheet)
    prods = defaultdict(int)
    for arts in rpt.postings.values():
        prods[production_of(next(iter(arts)))] += 1
    parts = [f"Едут {len(rpt.postings)}"] + [f"{p} {prods[p]}" for p in PRODUCTIONS]
    if rpt.red:
        parts.append(f"красных {len(rpt.red)}")
    if rpt.limits:
        parts.append(f"жёлтых не едут {len(rpt.yellow)}")
    if rpt.mixed or rpt.notes:
        parts.append("⚠ есть замечания — нажми «Проверить»")
    return "   ·   ".join(parts)


def wb_files_summary(files):
    lines = []
    for p in files:
        try:
            keep, _ = wb_report_filter(p)
            rpt = read_report(p, keep=keep)
            lines.append(f"{Path(p).stem}: едут {len(rpt.postings)}"
                         + (f" · красных {len(rpt.red)}" if rpt.red else "")
                         + (f" · жёлтых не едут {len(rpt.yellow)}" if rpt.limits else ""))
        except Exception as e:  # noqa: BLE001
            lines.append(f"{Path(p).stem}: не получилось прочитать — {e}")
    return "\n".join(lines)


# ================================================================ Пропуска Ozon

PASS_DEFAULT = {"on": True, "driver": "Алексей Солохов", "phone": "",
                "cars": "а139уо69 газель\nв678ха69 газель"}
_LAT2CYR = str.maketrans("ABCEHKMOPTXY", "АВСЕНКМОРТХУ")
CARRIAGE_DONE = ("cancelled", "closed", "sended", "received")


def plate_norm(p):
    """Номер машины: заглавные, без пробелов, латиница-двойник -> кириллица."""
    return "".join(str(p).upper().split()).translate(_LAT2CYR)


def parse_cars(text):
    """«а139уо69 газель; в678ха69 газель» -> [("А139УО69", "Газель"), ...]."""
    cars = []
    for line in str(text).replace(";", "\n").splitlines():
        parts = line.split(None, 1)
        if parts:
            model = (parts[1].strip() if len(parts) > 1 else "") or "Газель"
            cars.append((plate_norm(parts[0]), model[:1].upper() + model[1:]))
    return cars


def driver_full(name):
    """«Алексей Солохов» -> «Солохов Алексей» (Ozon ждёт ФИО с фамилии)."""
    parts = str(name).split()
    return " ".join(parts[1:] + parts[:1]) if len(parts) > 1 else " ".join(parts)


def _carriage_methods(h):
    methods, cursor = [], ""
    while True:
        r = ozon_post(f"{API}/v2/carriage/delivery/list", headers=h, timeout=30,
                      json={"limit": 1000, "cursor": cursor})
        if r.status_code != 200:
            raise RuntimeError(api_error(r))
        data = r.json()
        methods += data.get("methods") or []
        cursor = data.get("cursor") or ""
        if not data.get("has_next") or not cursor:
            return methods


def _make_pass(h, cab, cid, where, cars, driver, phone, log, tries=1):
    passes = [{"driver_name": driver, "driver_phone": phone, "vehicle_license_plate": plate,
               "vehicle_model": model, "with_returns": True} for plate, model in cars]
    r = None
    for i in range(tries):
        if i:
            time.sleep(5)   # только что сформированная отгрузка может появиться не сразу
        try:
            r = ozon_post(f"{API}/v1/carriage/pass/create", headers=h, timeout=30, net_retry=False,
                          json={"carriage_id": cid, "arrival_passes": passes})
        except requests.RequestException as e:
            log(f"✖ Пропуск Ozon {cab} · отгрузка {cid}: {e}")
            return False
        if r.status_code == 200:
            log(f"✔ Пропуск {cab} · {where} · отгрузка {cid} — " + ", ".join(p for p, _ in cars) + f" · {driver}")
            return True
    log(f"✖ Пропуск Ozon {cab} · отгрузка {cid}: {api_error(r)}")
    return False


def ozon_passes(cabinets, cfg, log, form=None, confirm=lambda text: True):
    """Пропуска Ozon (обе машины) на сегодняшние отгрузки без пропуска.
    form = {кабинет: {склады}} — сначала сформировать отгрузку по этим складам (как кнопка
    «Сформировать» в кабинете: в неё попадают все собранные отправления склада), потом пропуск.
    Без form — только пропуска на уже сформированные отгрузки."""
    cars = parse_cars(cfg.get("cars", ""))
    driver = driver_full(cfg.get("driver", ""))
    phone = "".join(ch for ch in str(cfg.get("phone", "")) if ch.isdigit() or ch == "+")
    if not cars or not driver:
        log("⚠ Пропуска Ozon: в настройках не указаны водитель или машины")
        return 0
    if not phone:
        log("⚠ Пропуска Ozon: не указан телефон водителя (Настройки → Пропуска) — Ozon может не принять")
    pass_cab = next(iter(cabinets), None)   # пропуска — только в первом кабинете
    data = {}
    for cab, conf in cabinets.items():
        if cab != pass_cab or (form is not None and cab not in form):   # отгрузка и пропуск — только первый кабинет
            continue
        try:
            data[cab] = _carriage_methods(headers_for(conf))
        except (RuntimeError, requests.RequestException) as e:
            log(f"✖ Отгрузки Ozon {cab}: {e}")

    # что формировать: методы доставки сегодняшних складов из этого запуска
    to_form = []
    if form:
        for cab, methods in data.items():
            whs = {str(w).lower() for w in form.get(cab, ())}
            for m in methods:
                name = m.get("warehouse_name") or m.get("delivery_method_name") or ""
                keys = {str(m.get(k)).lower() for k in ("warehouse_name", "delivery_method_name",
                                                         "warehouse_id", "delivery_method_id") if m.get(k)}
                mid = m.get("warehouse_id") or m.get("delivery_method_id")   # для FBS Ozon просит warehouse_id
                if keys & whs and mid and (cab, mid) not in {(c, x.get("warehouse_id") or x.get("delivery_method_id"))
                                                             for c, x, _ in to_form}:
                    to_form.append((cab, m, name))
        if to_form:
            lines = "\n".join(f"   {cab} · {name}" for cab, _, name in to_form)
            if not confirm(f"Сформировать отгрузку в Ozon и сделать пропуска?\n\n{lines}\n\n"
                           f"В отгрузку попадут все собранные отправления этих складов — "
                           f"как кнопка «Сформировать» в кабинете. На КГТ — только отгрузка, без пропуска."
                           f"\n\nДа — сформировать и сделать пропуска\n"
                           f"Нет — ничего не делать"):
                log("Отгрузки не сформированы — ты выбрал «Нет».")
                return 0

    made, formed = 0, set()
    for cab, m, name in to_form:
        h = headers_for(cabinets[cab])
        mid = m.get("delivery_method_id") or m.get("warehouse_id")
        # /v2/posting/fbs/act/create Ozon отключил: теперь отгрузку создают и подтверждают двумя методами
        try:
            r = ozon_post(f"{API}/v1/carriage/create", headers=h, timeout=60, net_retry=False,
                          json={"delivery_method_id": int(mid)})
        except requests.RequestException as e:
            log(f"✖ Отгрузка {cab} · {name}: {e} — проверь в кабинете, создалась ли")
            continue
        if r.status_code != 200:
            log(f"✖ Отгрузка {cab} · {name} не создана: {api_error(r)}")
            continue
        cid = r.json().get("carriage_id") or (r.json().get("result") or {}).get("carriage_id")
        if not cid:
            log(f"✖ Отгрузка {cab} · {name}: Ozon не вернул номер отгрузки — проверь в кабинете")
            continue
        log(f"… отгрузка {cab} · {name} № {cid} создана, подтверждаю")
        a, err = None, ""
        for i in range(8):   # Ozon собирает отгрузку не сразу — подтверждение может пройти со 2–3 раза
            if i:
                time.sleep(5)
            try:
                a = ozon_post(f"{API}/v1/carriage/approve", headers=h, timeout=60, net_retry=False,
                              json={"carriage_id": int(cid)})
            except requests.RequestException as e:
                err = str(e)
                continue
            if a.status_code == 200:
                break
            err = api_error(a)
        if a is None or a.status_code != 200:
            log(f"✖ Отгрузка {cab} · {name} № {cid} создана, но не подтверждена: {err} "
                f"— подтверди в кабинете и пришли скриншот этой строки")
            continue
        log(f"✔ Отгрузка сформирована {cab} · {name} · № {cid}")
        formed.add(int(cid))
        if is_kgt(name):
            log(f"   {name} — КГТ, пропуск не нужен")
        elif cab == pass_cab:
            made += _make_pass(h, cab, int(cid), name, cars, driver, phone, log, tries=4)

    # уже сформированные отгрузки без пропуска
    tried = 0
    for cab, methods in data.items():
        if cab != pass_cab:
            continue
        h = headers_for(cabinets[cab])
        for m in methods:
            name = m.get("warehouse_name") or m.get("delivery_method_name") or ""
            if is_kgt(name):   # на КГТ пропуск не делаем
                continue
            for c in m.get("carriages") or []:
                cid = c.get("id")
                if not cid or int(cid) in formed or c.get("status") in CARRIAGE_DONE:
                    continue
                if c.get("available_actions") and "set_arrival_passes" not in c["available_actions"]:
                    continue
                try:
                    g = ozon_post(f"{API}/v1/carriage/get", headers=h, timeout=30, json={"carriage_id": cid})
                    if g.status_code == 200 and (g.json().get("arrival_pass_ids") or []):
                        continue
                except requests.RequestException:
                    continue
                tried += 1
                made += _make_pass(h, cab, int(cid), name, cars, driver, phone, log)
    if not tried and not to_form:
        log("… Пропуска Ozon: нет сформированных отгрузок без пропуска")
    return made


def today_msk():
    return datetime.now(MSK).date()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1060x920")
        self.minsize(960, 760)
        self.configure(bg=C["bg"])
        self.q = queue.Queue()
        self.busy = False
        self.buttons = []      # кнопки, которые гасятся, пока идёт работа
        self.cab_rows = []
        self.wb_files = []
        self._save_job = None
        self.setup_style()

        # --- переменные настроек
        self.template = tk.StringVar()
        self.root_dir = tk.StringVar(value=str(self.default_root()))
        self.by_day = tk.BooleanVar(value=False)   # папки по дням больше не делаем: главную папку выбираешь сам
        self.days = tk.IntVar(value=14)
        self.top_margin = tk.DoubleVar(value=DEFAULT_TOP_MARGIN_PT)
        self.font_size = tk.DoubleVar(value=DEFAULT_FONT_SIZE)
        self.wb_token = tk.StringVar()
        self.sklad_url = tk.StringVar(value="https://89-169-45-92.sslip.io")
        self.sklad_key = tk.StringVar()
        self.cdek_account = tk.StringVar()
        self.cdek_secure = tk.StringVar()
        self.excel = tk.StringVar()
        self.sheet = tk.StringVar()
        self.footer = tk.StringVar(value="Озон утро")
        self.autoship = tk.BooleanVar(value=True)
        self._pass_running = {}
        self.last_ozon_out = None   # папка последнего скачивания Ozon — для «Пересобрать ленты»
        self.pass_on = tk.BooleanVar(value=PASS_DEFAULT["on"])
        self.pass_driver = tk.StringVar(value=PASS_DEFAULT["driver"])
        self.pass_phone = tk.StringVar(value=PASS_DEFAULT["phone"])
        self.pass_cars = tk.StringVar(value=PASS_DEFAULT["cars"].replace("\n", "; "))
        self.wave = tk.StringVar(value=WAVES[0])
        self.show_keys = tk.BooleanVar(value=False)

        # --- боковая панель: логотип, разделы, дата
        side = tk.Frame(self, bg=C["side"], width=210)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        logo = tk.Frame(side, bg=C["side"])
        logo.pack(fill="x", padx=24, pady=(26, 30))
        tk.Label(logo, text="BLAGO", bg=C["side"], fg=C["accent"], font=(FONT, 20, "bold")).pack(anchor="w")
        tk.Label(logo, text="OTGRUZKI", bg=C["side"], fg=C["text"], font=(FONT, 13, "bold")).pack(anchor="w")
        self.nav = {}
        for name in PAGES:
            self.nav[name] = self.nav_item(side, name)
        d = today_msk()
        foot = tk.Frame(side, bg=C["side"])
        foot.pack(side="bottom", fill="x", padx=24, pady=22)
        tk.Label(foot, text=f"{d.day} {MONTHS_GEN[d.month - 1]}", bg=C["side"], fg=C["text"],
                 font=(FONT, 13, "bold")).pack(anchor="w")
        tk.Label(foot, text=WEEKDAYS[d.weekday()], bg=C["side"], fg=C["muted"], font=(FONT, 10)).pack(anchor="w")
        tk.Label(foot, text=f"версия {APP_VERSION}", bg=C["side"], fg=C["dim"],
                 font=(FONT, 8)).pack(anchor="w", pady=(10, 0))

        # --- основная часть: заголовок раздела, раздел, журнал
        main = ttk.Frame(self, style="Bg.TFrame")
        main.pack(side="left", fill="both", expand=True, padx=28, pady=(22, 20))
        top = ttk.Frame(main, style="Bg.TFrame")
        top.pack(fill="x", pady=(0, 10))
        self.page_title = ttk.Label(top, text="", style="Page.TLabel")
        self.page_title.pack(side="left")
        # главная папка дня — видна всегда и меняется в один клик
        ttk.Button(top, text="Сменить папку…", command=self.pick_root, cursor="hand2").pack(side="right")
        self.root_label = ttk.Button(top, text="", style="BgPath.TButton", cursor="hand2",
                                     command=lambda: self.open_out(self.root_dir.get() or self.default_root()))
        self.root_label.pack(side="right", padx=(0, 10))
        # Склад: что последним ушло и не застряло ли что-то без интернета
        srow = tk.Frame(main, bg=C["bg"])
        srow.pack(fill="x", pady=(0, 10))
        self.sklad_line = tk.Label(srow, text="", bg=C["bg"], fg=C["muted"], font=(FONT, 10, "bold"), anchor="w")
        self.sklad_line.pack(side="left")
        self.sklad_resend = ttk.Button(srow, text="Отправить неотправленное", command=self.flush_sklad,
                                       cursor="hand2")
        # Заказы дня уходят на склад только после проверки — этой кнопкой
        self.sklad_send = ttk.Button(srow, text="Отправить на склад", style="Accent.TButton",
                                     command=self.send_morning, cursor="hand2")
        self.page_box = ttk.Frame(main, style="Bg.TFrame")
        self.page_box.pack(fill="x")
        self.pages = {name: ttk.Frame(self.page_box, style="Bg.TFrame") for name in PAGES}
        self.build_sborka(self.pages["Сборка"])
        self.build_ozon(self.pages["Ozon"])
        self.build_wb(self.pages["Wildberries"])
        self.build_cdek(self.pages["СДЭК"])
        self.build_settings(self.scroll_page(self.pages["Настройки"]))

        jb = tk.Frame(main, bg=C["line"])
        jb.pack(fill="both", expand=True, pady=(14, 0))
        jcard = tk.Frame(jb, bg=C["card"])
        jcard.pack(fill="both", expand=True, padx=1, pady=1)
        jh = tk.Frame(jcard, bg=C["card"])
        jh.pack(fill="x", padx=18, pady=(12, 0))
        tk.Label(jh, text="Журнал", bg=C["card"], fg=C["text"], font=(FONT, 12, "bold")).pack(side="left")
        self.state_pill = tk.Label(jh, text="●  готова к работе", bg=C["card"], fg=C["ok"], font=(FONT, 9, "bold"))
        self.state_pill.pack(side="left", padx=(14, 0))
        ttk.Button(jh, text="очистить", style="Link.TButton", command=self.clear_log,
                   cursor="hand2").pack(side="right")
        self.bar = ttk.Progressbar(jcard, mode="determinate")   # виден только пока идёт работа
        box = tk.Frame(jcard, bg=C["card"])
        box.pack(fill="both", expand=True, padx=(8, 4), pady=(6, 8))
        self._log_frame = box
        self.log_box = tk.Text(box, height=10, wrap="word", state="disabled", bd=0, relief="flat",
                               bg=C["card"], fg=C["text"], insertbackground=C["text"],
                               selectbackground=C["accent"], selectforeground=C["on_accent"],
                               padx=10, pady=6, font=("Consolas", 10), highlightthickness=0,
                               spacing1=2, spacing3=2)
        sb = ttk.Scrollbar(box, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_box.pack(fill="both", expand=True)
        for tag, color, bold in (("ok", C["ok"], False), ("err", C["err"], False), ("head", C["accent"], True),
                                 ("warn", C["warn"], True), ("warnbody", C["warn"], False),
                                 ("yellow", C["accent"], False), ("big", C["ok"], True),
                                 ("muted", C["dim"], False)):
            self.log_box.tag_configure(tag, foreground=color,
                                       font=("Consolas", 11 if tag == "big" else 10, "bold" if bold else ""))

        self.setup_clipboard()
        self.load_settings()
        self.refresh_wave_chips()
        self.hand_cursor(self)
        for var in (self.template, self.root_dir, self.by_day, self.days, self.top_margin, self.font_size,
                    self.wb_token, self.autoship, self.pass_on, self.pass_driver, self.pass_phone,
                    self.pass_cars, self.sklad_url, self.sklad_key, self.cdek_account, self.cdek_secure):
            var.trace_add("write", lambda *_: self.schedule_save())
        self.footer.trace_add("write", lambda *_: self.refresh_links())
        self.show_page("Сборка")
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(50, self.dark_titlebar)
        self.after(100, self.poll)
        self.after(500, self.refresh_sklad_line)

    def refresh_sklad_line(self):
        """Строка «Склад: …»: подключён ли, что ушло последним, что ждёт отправки."""
        try:
            pending = len(_pending_load())
        except Exception:  # noqa: BLE001
            pending = 0
        morning = morning_load() if sklad_on() else None
        if not sklad_on():
            text, color = "Склад: не подключён — «Настройки → Склад BLAGO»", C["dim"]
        elif morning and not pending:
            warn = morning_problems(morning)
            text = f"Склад: ждёт проверки — {morning_text(morning)}"
            color = C["warn"] if warn else C["accent"]
        elif pending:
            text = f"Склад: ✖ не ушло {plural(pending, 'партия', 'партии', 'партий')} — нет связи"
            color = C["err"]
        elif SKLAD_LAST["text"]:
            ok = SKLAD_LAST["kind"] == "ok"
            text, color = f"Склад: {'✔' if ok else '✖'} {SKLAD_LAST['text']}", C["ok"] if ok else C["err"]
        else:
            text, color = "Склад: подключён, сегодня ещё ничего не отправлялось", C["muted"]
        self.sklad_line.configure(text=text, fg=color)
        if pending and not self.sklad_resend.winfo_ismapped():
            self.sklad_resend.pack(side="left", padx=(12, 0))
        elif not pending and self.sklad_resend.winfo_ismapped():
            self.sklad_resend.pack_forget()
        show_send = bool(morning) and not pending
        if show_send and not self.sklad_send.winfo_ismapped():
            self.sklad_send.pack(side="left", padx=(12, 0))
        elif not show_send and self.sklad_send.winfo_ismapped():
            self.sklad_send.pack_forget()
        self.after(2000, self.refresh_sklad_line)

    # ------------------------------------------------------------ навигация
    def nav_item(self, parent, name):
        row = tk.Frame(parent, bg=C["side"], cursor="hand2")
        row.pack(fill="x")
        bar = tk.Frame(row, bg=C["side"], width=4)
        bar.pack(side="left", fill="y")
        lbl = tk.Label(row, text=name, bg=C["side"], fg=C["muted"], font=(FONT, 11, "bold"),
                       anchor="w", padx=20, pady=11, cursor="hand2")
        lbl.pack(side="left", fill="x", expand=True)
        for w in (row, lbl):
            w.bind("<Button-1>", lambda e, n=name: self.show_page(n))
            w.bind("<Enter>", lambda e, n=name: self.paint_nav(n, hover=True))
            w.bind("<Leave>", lambda e, n=name: self.paint_nav(n))
        return row, bar, lbl

    def paint_nav(self, name, hover=False):
        row, bar, lbl = self.nav[name]
        on = name == getattr(self, "current_page", None)
        bg = C["nav_hi"] if on or hover else C["side"]
        for w in (row, lbl):
            w.configure(bg=bg)
        bar.configure(bg=C["accent"] if on else bg)
        lbl.configure(fg=C["text"] if on or hover else C["muted"])

    def show_page(self, name):
        self.current_page = name
        if name == "СДЭК" and hasattr(self, "cdek_keys"):
            self.refresh_cdek()  # ключи могли появиться в «Настройках»
        for n, frame in self.pages.items():
            if n == name:
                frame.pack(fill="x")
            else:
                frame.pack_forget()
        for n in PAGES:
            self.paint_nav(n)
        self.page_title.configure(text=name)

    def hand_cursor(self, widget):
        """Рука над всем, что нажимается."""
        for w in widget.winfo_children():
            if isinstance(w, (ttk.Button, ttk.Checkbutton, ttk.Radiobutton, ttk.Combobox, tk.Radiobutton)):
                try:
                    w.configure(cursor="hand2")
                except tk.TclError:
                    pass
            self.hand_cursor(w)

    # ------------------------------------------------------------ оформление
    def setup_style(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        base = (FONT, 10)
        st.configure(".", background=C["card"], foreground=C["text"], font=base,
                     bordercolor=C["line"], lightcolor=C["card"], darkcolor=C["card"],
                     troughcolor=C["field"], focuscolor=C["accent"])
        st.configure("Bg.TFrame", background=C["bg"])
        st.configure("Card.TFrame", background=C["card"])
        st.configure("TLabel", background=C["card"], foreground=C["text"])
        st.configure("Muted.TLabel", foreground=C["muted"], font=(FONT, 9))
        st.configure("Info.TLabel", foreground=C["muted"], font=(FONT, 10))
        st.configure("Field.TLabel", foreground=C["muted"], font=(FONT, 10))
        st.configure("Caps.TLabel", background=C["bg"], foreground=C["muted"], font=(FONT, 8, "bold"))
        st.configure("Page.TLabel", background=C["bg"], foreground=C["text"], font=(FONT, 20, "bold"))
        st.configure("Title.TLabel", font=(FONT, 12, "bold"))
        st.configure("File.TLabel", font=(FONT, 11, "bold"))
        st.configure("Status.TLabel", foreground=C["muted"], font=(FONT, 9))

        st.configure("TButton", background=C["btn"], foreground=C["text"], borderwidth=0,
                     padding=(14, 8), focusthickness=0, relief="flat")
        st.map("TButton", background=[("disabled", C["card"]), ("pressed", C["line"]), ("active", C["btn_hi"])],
               foreground=[("disabled", C["dim"])])
        st.configure("Accent.TButton", background=C["accent"], foreground=C["on_accent"],
                     font=(FONT, 11, "bold"), padding=(24, 11))
        st.map("Accent.TButton",
               background=[("disabled", C["btn"]), ("pressed", C["accent"]), ("active", C["accent_hi"])],
               foreground=[("disabled", C["dim"])])
        st.configure("Big.TButton", font=(FONT, 10, "bold"), padding=(18, 11))
        st.configure("Link.TButton", background=C["card"], foreground=C["muted"], padding=(2, 2),
                     font=(FONT, 9, "underline"))
        st.map("Link.TButton", background=[("active", C["card"]), ("pressed", C["card"])],
               foreground=[("active", C["accent"]), ("disabled", C["dim"])])
        st.configure("Path.TButton", background=C["card"], foreground=C["muted"], padding=(2, 2),
                     font=(FONT, 10))
        st.map("Path.TButton", background=[("active", C["card"]), ("pressed", C["card"])],
               foreground=[("active", C["text"])])
        st.configure("BgPath.TButton", background=C["bg"], foreground=C["muted"], padding=(2, 6),
                     font=(FONT, 10))
        st.map("BgPath.TButton", background=[("active", C["bg"]), ("pressed", C["bg"])],
               foreground=[("active", C["text"])])
        st.configure("BgLink.TButton", background=C["bg"], foreground=C["muted"], padding=(2, 2),
                     font=(FONT, 9, "underline"))
        st.map("BgLink.TButton", background=[("active", C["bg"])], foreground=[("active", C["accent"])])

        for w in ("TEntry", "TCombobox", "TSpinbox"):
            st.configure(w, fieldbackground=C["field"], foreground=C["text"], insertcolor=C["text"],
                         bordercolor=C["line"], lightcolor=C["field"], darkcolor=C["field"],
                         arrowcolor=C["muted"], background=C["btn"], padding=6,
                         selectbackground=C["accent"], selectforeground=C["on_accent"])
            st.map(w, fieldbackground=[("readonly", C["field"]), ("disabled", C["card"])],
                   foreground=[("readonly", C["text"])],
                   bordercolor=[("focus", C["accent"])],
                   lightcolor=[("focus", C["accent"])], darkcolor=[("focus", C["accent"])],
                   selectbackground=[("readonly", C["field"])],
                   selectforeground=[("readonly", C["text"])])
        self.option_add("*TCombobox*Listbox.background", C["field"])
        self.option_add("*TCombobox*Listbox.foreground", C["text"])
        self.option_add("*TCombobox*Listbox.selectBackground", C["accent"])
        self.option_add("*TCombobox*Listbox.selectForeground", C["on_accent"])
        self.option_add("*TCombobox*Listbox.font", base)

        st.configure("TCheckbutton", background=C["card"], foreground=C["text"],
                     indicatorbackground=C["field"], indicatorforeground=C["accent"])
        st.map("TCheckbutton", background=[("active", C["card"])], foreground=[("active", C["accent_hi"])],
               indicatorbackground=[("selected", C["field"])])
        st.configure("TNotebook", background=C["bg"], borderwidth=0, tabmargins=(0, 0, 0, 0))
        st.configure("TNotebook.Tab", background=C["btn"], foreground=C["muted"], padding=(18, 8),
                     borderwidth=0, font=(FONT, 10, "bold"))
        st.map("TNotebook.Tab", background=[("selected", C["accent"]), ("active", C["btn_hi"])],
               padding=[("selected", (18, 8))], expand=[("selected", (0, 0, 0, 0))],
               foreground=[("selected", C["on_accent"]), ("active", C["text"])])
        st.configure("Horizontal.TProgressbar", background=C["accent"], troughcolor=C["field"],
                     bordercolor=C["card"], lightcolor=C["accent"], darkcolor=C["accent"], thickness=4)
        # Ползунок — одна гладкая полоса: без стрелок, без рисок и объёмных краёв
        st.layout("Vertical.TScrollbar", [("Vertical.Scrollbar.trough", {"sticky": "ns", "children": [
            ("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
        st.configure("Vertical.TScrollbar", background=C["btn_hi"], troughcolor=C["card"], bordercolor=C["card"],
                     lightcolor=C["btn_hi"], darkcolor=C["btn_hi"], gripcount=0, relief="flat", arrowsize=8,
                     borderwidth=0)
        st.map("Vertical.TScrollbar", background=[("active", C["muted"]), ("pressed", C["muted"])],
               lightcolor=[("active", C["muted"]), ("pressed", C["muted"])],
               darkcolor=[("active", C["muted"]), ("pressed", C["muted"])])

    def card(self, parent, title, side=None, grid=None):
        """Тёмная карточка с заголовком. side — виджет-функция для правой части заголовка."""
        border = tk.Frame(parent, bg=C["line"])
        if grid:
            border.grid(row=grid[0], column=grid[1], sticky="nsew", pady=5,
                        padx=(0, 6) if grid[1] == 0 else (6, 0))
        else:
            border.pack(fill="x", pady=5)
        body = ttk.Frame(border, style="Card.TFrame", padding=(18, 12, 18, 16))
        body.pack(fill="both", expand=True, padx=1, pady=1)
        head = ttk.Frame(body, style="Card.TFrame")
        head.pack(fill="x", pady=(0, 10))
        ttk.Label(head, text=title, style="Title.TLabel").pack(side="left")
        if side:
            side(head)
        return body

    def button(self, parent, text, command, style="TButton", **pack):
        b = ttk.Button(parent, text=text, command=command, style=style, cursor="hand2")
        b.pack(side="left", **pack)
        self.buttons.append(b)
        return b

    def folder_link(self, parent, sub):
        """Строка «Сохраняется в: Рабочий стол / Октябрь 2026 / 29.10 / Озон 2» — нажатие открывает папку."""
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill="x", pady=(12, 0))
        ttk.Label(row, text="Сохраняется в:", style="Field.TLabel").pack(side="left")
        var = tk.StringVar()
        ttk.Button(row, textvariable=var, style="Path.TButton",
                   command=lambda: self.open_out(self.day_dir(sub() if callable(sub) else sub))
                   ).pack(side="left", padx=(6, 0))
        self._links = getattr(self, "_links", [])
        self._links.append((var, sub))
        return var

    def refresh_links(self):
        if hasattr(self, "root_label"):
            r = Path(self.root_dir.get().strip() or self.default_root())
            self.root_label.configure(text="Папка:  " + "  ›  ".join(r.parts[-3:]))
        for var, sub in getattr(self, "_links", []):
            p = self.day_dir(sub() if callable(sub) else sub)
            var.set("  ›  ".join(p.parts[-2:]) + "     открыть ↗")

    def dark_titlebar(self):
        """Тёмный заголовок окна в Windows 10/11."""
        if sys.platform != "win32":
            return
        try:
            import ctypes
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            val = ctypes.c_int(1)
            for attr in (20, 19):
                ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(val), ctypes.sizeof(val))
        except Exception:  # noqa: BLE001 — украшение, не критично
            pass

    # ------------------------------------------------------------ вкладка «Сборка»
    def build_sborka(self, tab):
        c1 = self.card(tab, "Утренняя сборка")
        row = ttk.Frame(c1, style="Card.TFrame")
        row.pack(fill="x")
        self.button(row, "Скачать отчёты и сборку", self.start_morning, "Accent.TButton")
        self.button(row, "Добавить СДЭК…", self.start_cdek, "Big.TButton", padx=(10, 0))
        self.folder_link(c1, "Сборка")
        self.mo_status = ttk.Label(c1, text="", style="Info.TLabel")
        self.mo_status.pack(anchor="w", pady=(10, 0))

        c2 = self.card(tab, "В течение дня")
        chips = ttk.Frame(c2, style="Card.TFrame")
        chips.pack(fill="x")
        self.wave_chips = {}
        for i, name in enumerate(WAVES):
            if i and "кгт" in name.lower() and "кгт" not in WAVES[i - 1].lower():
                tk.Frame(chips, bg=C["line"], width=1, height=26).pack(side="left", padx=10)
            b = tk.Radiobutton(chips, text=name, value=name, variable=self.wave, indicatoron=0,
                               command=self.refresh_wave_chips, bd=0, relief="flat", cursor="hand2",
                               font=(FONT, 10, "bold"), padx=16, pady=7, highlightthickness=0,
                               bg=C["btn"], fg=C["text"], activebackground=C["btn_hi"],
                               activeforeground=C["text"], selectcolor=C["accent"])
            b.pack(side="left", padx=(0, 6))
            self.wave_chips[name] = b
        row2 = ttk.Frame(c2, style="Card.TFrame")
        row2.pack(fill="x", pady=(14, 0))
        self.wave_btn = self.button(row2, "Скачать отчёт и сборку", self.start_wave, "Accent.TButton")
        self.folder_link(c2, self.wave_name)
        self.wave_status = ttk.Label(c2, text="", style="Info.TLabel")
        self.wave_status.pack(anchor="w", pady=(10, 0))

    def wave_name(self):
        return self.wave.get() or WAVES[0]

    def refresh_wave_chips(self):
        for name, b in self.wave_chips.items():
            on = name == self.wave.get()
            b.configure(bg=C["accent"] if on else C["btn"], fg=C["on_accent"] if on else C["text"],
                        activebackground=C["accent_hi"] if on else C["btn_hi"],
                        activeforeground=C["on_accent"] if on else C["text"])
        self.refresh_links()

    # ------------------------------------------------------------ вкладка Ozon
    def build_ozon(self, tab):
        c1 = self.card(tab, "Отчёт")
        row = ttk.Frame(c1, style="Card.TFrame")
        row.pack(fill="x")
        self.excel_name = ttk.Label(row, text="не выбран", style="File.TLabel")
        self.excel_name.pack(side="left")
        ttk.Button(row, text="Выбрать файл…", command=self.pick_excel).pack(side="right")
        self.sheet_box = ttk.Frame(row, style="Card.TFrame")
        ttk.Label(self.sheet_box, text="лист", style="Field.TLabel").pack(side="left", padx=(0, 6))
        self.sheet_cb = ttk.Combobox(self.sheet_box, textvariable=self.sheet, state="readonly", width=14)
        self.sheet_cb.pack(side="left")
        self.sheet_cb.bind("<<ComboboxSelected>>", lambda e: self.count_report())
        self.report_info = ttk.Label(c1, text="Отчёт подставляется сам после «Сборки». "
                                              "Покрась его в Excel, сохрани и закрой.", style="Info.TLabel")
        self.report_info.pack(anchor="w", pady=(8, 0))

        c2 = self.card(tab, "Этикетки")
        g = ttk.Frame(c2, style="Card.TFrame")
        g.pack(fill="x")
        ttk.Label(g, text="Надпись на ШК", style="Field.TLabel").pack(side="left")
        fcb = ttk.Combobox(g, textvariable=self.footer, width=16, values=FOOTERS)
        fcb.pack(side="left", padx=(10, 24))
        fcb.bind("<<ComboboxSelected>>", lambda e: self.refresh_links())
        fcb.bind("<KeyRelease>", lambda e: self.refresh_links())
        ttk.Checkbutton(g, variable=self.autoship, text="Собирать заказы из «Ожидают сборки»"
                        ).pack(side="left")
        row2 = ttk.Frame(c2, style="Card.TFrame")
        row2.pack(fill="x", pady=(16, 0))
        self.button(row2, "Проверить", lambda: self.start("check"), "Big.TButton")
        self.button(row2, "Скачать и собрать", lambda: self.start("run"), "Accent.TButton", padx=(10, 0))
        self.button(row2, "Пересобрать ленты", lambda: self.start("merge"), padx=(10, 0))
        self.button(row2, "Пропуска", self.start_passes, padx=(10, 0))
        self.folder_link(c2, self.ozon_sub)

    def ozon_sub(self):
        """Папка этикеток Ozon: утренние — «Озон», дальше «Озон 2», «Озон 3», «КГТ 2»…"""
        name = " ".join(self.footer.get().split())
        return "Озон" if not name or name.lower() in ("озон утро", "озон 1") else name

    # ------------------------------------------------------------ вкладка Wildberries
    def build_wb(self, tab):
        c1 = self.card(tab, "Отчёты ВБ и СГТ")
        row = ttk.Frame(c1, style="Card.TFrame")
        row.pack(fill="x")
        self.wb_files_label = ttk.Label(row, text="не выбраны", style="File.TLabel")
        self.wb_files_label.pack(side="left")
        ttk.Button(row, text="Очистить", command=self.clear_wb).pack(side="right")
        ttk.Button(row, text="Добавить файл…", command=self.pick_wb).pack(side="right", padx=(0, 8))
        self.wb_info = ttk.Label(c1, text="Отчёты подставляются сами после «Сборки».", style="Info.TLabel",
                                 justify="left")
        self.wb_info.pack(anchor="w", pady=(8, 0))

        c2 = self.card(tab, "Поставки и стикеры")
        row2 = ttk.Frame(c2, style="Card.TFrame")
        row2.pack(fill="x")
        self.button(row2, "Проверить", lambda: self.start_wb("check"), "Big.TButton")
        self.button(row2, "Создать поставки и скачать", lambda: self.start_wb("run"), "Accent.TButton",
                    padx=(10, 0))
        self.button(row2, "Скачать ШК открытых поставок", lambda: self.start_wb("stickers"), padx=(10, 0))
        self.folder_link(c2, "ВБ")

    # ------------------------------------------------------------ вкладка «СДЭК»
    def build_cdek(self, tab):
        self.cdek_file = None
        c1 = self.card(tab, "Заказы СДЭК")
        row = ttk.Frame(c1, style="Card.TFrame")
        row.pack(fill="x")
        self.cdek_file_label = ttk.Label(row, text="не выбран", style="File.TLabel")
        self.cdek_file_label.pack(side="left")
        ttk.Button(row, text="Выбрать файл…", command=self.pick_cdek_file).pack(side="right")
        self.cdek_info = ttk.Label(c1, text="Файл из «макса»: № задания, Артикул продавца, трэк. Покраска — "
                                            "как в отчётах Ozon: красный не едет, жёлтый — сколько указано рядом.",
                                   style="Info.TLabel", justify="left", wraplength=720)
        self.cdek_info.pack(anchor="w", pady=(8, 0))

        c3 = self.card(tab, "Листы сборки, ШК и акты")
        row3 = ttk.Frame(c3, style="Card.TFrame")
        row3.pack(fill="x")
        self.button(row3, "Сделать листы, ШК и акты", self.start_cdek_docs, "Accent.TButton")
        self.cdek_keys = ttk.Label(row3, text="", style="Info.TLabel")
        self.cdek_keys.pack(side="left", padx=(14, 0))
        self.folder_link(c3, "СДЭК")
        self.cdek_done = ttk.Label(c3, text="", style="Info.TLabel")
        self.cdek_done.pack(anchor="w", pady=(10, 0))
        self.refresh_cdek()

    def refresh_cdek(self):
        if cdek_on():
            self.cdek_keys.configure(text="ШК качаются по ключам СДЭК", foreground=C["muted"])
        else:
            self.cdek_keys.configure(text="нет ключей СДЭК — «Настройки → СДЭК»", foreground=C["err"])

    def pick_cdek_file(self):
        path = filedialog.askopenfilename(title="Заказы СДЭК", filetypes=[("Excel", "*.xlsx *.xlsm")])
        if path:
            self.load_cdek_file(path)

    def load_cdek_file(self, path):
        """Прочитать файл заказов: сколько едет на каждое производство с учётом покраски."""
        try:
            orders = read_cdek_orders(path)
            going, stopped = cdek_filter(path, orders) if orders else ([], [])
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"Не удалось прочитать файл: {e}")
            return
        if not orders:
            messagebox.showwarning(APP_NAME, "В файле нет колонки «трэк» или нет заказов с трэком.")
            return
        self.cdek_file = path
        self.cdek_file_label.configure(text=Path(path).name)
        kom = sum(production_of(o["article"]) == KOMINTERNA for o in going)
        text = f"Едут {len(going)}: {KOMINTERNA} {kom}, {UCHITELSKAYA} {len(going) - kom}"
        if stopped:
            text += f" · не едут (покраска): {len(stopped)}"
        self.cdek_info.configure(text=text + ". Покрасил или поправил файл — выбери его ещё раз.")

    def start_cdek_docs(self):
        if self.busy:
            return
        if not self.cdek_file:
            messagebox.showwarning(APP_NAME, "Сначала выбери файл заказов СДЭК.")
            return
        self.refresh_cdek()
        if not cdek_on() and not messagebox.askyesno(
                APP_NAME, "Нет ключей СДЭК — ШК не скачаются.\nСделать только листы сборки и акты?"):
            return
        path, out = self.cdek_file, self.day_dir("СДЭК")
        self.headline("СДЭК: ЛИСТЫ, ШК И АКТЫ")
        self.cdek_done.configure(text="… делаю", foreground=C["muted"])

        def work():
            log = lambda s: self.q.put(("log", s))  # noqa: E731
            try:
                orders, stopped = cdek_filter(path, read_cdek_orders(path), log)  # заново — свежая покраска
                if not orders:
                    log("✖ В файле не осталось заказов, которые едут")
                else:
                    make_cdek(orders, out, log)
                if stopped:
                    Path(out).mkdir(parents=True, exist_ok=True)
                    name = f"СДЭК не уехало {datetime.now(MSK):%d.%m}.xlsx"
                    write_cdek_stopped(Path(out) / name, stopped, path)
                    log(f"⛔ Не уехало {len(stopped)} — список для менеджеров: «СДЭК» / {name}")
                    log("│ " + ", ".join(o["task"] for o in stopped))
                    self.q.put(("status", self.cdek_done,
                                f"✔ Готово в {datetime.now(MSK):%H:%M} — {len(orders)} заказов, папка «СДЭК»",
                                C["ok"]))
            except PermissionError:
                log("✖ Файл в папке «СДЭК» открыт (Excel/Word/PDF) — закрой и нажми ещё раз")
            except Exception as e:  # noqa: BLE001
                log(f"✖ Ошибка: {e}")
            self.q.put(("done",))

        self.run(work)

    # ------------------------------------------------------------ вкладка «Настройки»
    def scroll_page(self, page):
        """Страница с прокруткой: длинные настройки листаются колесом, журнал внизу остаётся виден."""
        canvas = tk.Canvas(page, bg=C["bg"], highlightthickness=0, bd=0, height=300)
        bar = ttk.Scrollbar(page, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas, style="Bg.TFrame")
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=bar.set)
        canvas.pack(side="left", fill="x", expand=True)
        bar.pack(side="right", fill="y", padx=(6, 0))
        journal_min = 190  # столько оставляем журналу внизу окна

        def fit(_=None):
            need = inner.winfo_reqheight()
            top = self.page_box.winfo_rooty() - self.winfo_rooty()
            room = max(220, self.winfo_height() - top - journal_min)
            height = min(need, room)
            if int(canvas.cget("height")) != height:
                canvas.configure(height=height)
            canvas.configure(scrollregion=(0, 0, 0, need))

        def wheel(e):
            if getattr(self, "current_page", "") != "Настройки":
                return
            w = self.winfo_containing(e.x_root, e.y_root)
            if w is not None and str(w).startswith(str(canvas)):
                canvas.yview_scroll(-1 if e.delta > 0 else 1, "units")

        inner.bind("<Configure>", fit)
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))
        self.bind("<Configure>", fit, add="+")
        self.bind_all("<MouseWheel>", wheel, add="+")
        return inner

    def build_settings(self, tab):
        c1 = self.card(tab, "Кабинеты Ozon", side=lambda h: ttk.Checkbutton(
            h, text="показать ключи", variable=self.show_keys, command=self.toggle_keys).pack(side="right"))
        self.cab_box = ttk.Frame(c1, style="Card.TFrame")
        self.cab_box.pack(fill="x")
        for i, t in enumerate(("Название", "Client-Id", "Api-Key", "")):
            ttk.Label(self.cab_box, text=t, style="Muted.TLabel").grid(row=0, column=i, sticky="w", padx=(0, 8))
        self.cab_box.columnconfigure(2, weight=1)
        b1 = ttk.Frame(c1, style="Card.TFrame")
        b1.pack(fill="x", pady=(8, 0))
        self.btn_test = self.button(b1, "Проверить ключи", self.test_keys)
        ttk.Button(b1, text="+ кабинет", command=self.add_cabinet).pack(side="left", padx=(8, 0))

        pair = ttk.Frame(tab, style="Bg.TFrame")
        pair.pack(fill="x")
        pair.columnconfigure(0, weight=1, uniform="p")
        pair.columnconfigure(1, weight=1, uniform="p")

        c2 = self.card(pair, "Wildberries", grid=(0, 0))
        g = ttk.Frame(c2, style="Card.TFrame")
        g.pack(fill="x")
        g.columnconfigure(1, weight=1)
        ttk.Label(g, text="Токен", style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.wb_token_entry = ttk.Entry(g, textvariable=self.wb_token, show="•")
        self.wb_token_entry.grid(row=0, column=1, sticky="ew")
        b2 = ttk.Frame(c2, style="Card.TFrame")
        b2.pack(fill="x", pady=(8, 0))
        self.button(b2, "Проверить токен", self.test_wb)
        self.wb_status = ttk.Label(b2, text="", style="Status.TLabel")
        self.wb_status.pack(side="left", padx=(10, 0))
        self.wb_show = tk.BooleanVar(value=False)
        ttk.Checkbutton(b2, text="показать", variable=self.wb_show,
                        command=lambda: self.wb_token_entry.configure(show="" if self.wb_show.get() else "•")
                        ).pack(side="right")

        c3 = self.card(pair, "Файлы", grid=(0, 1))
        g3 = ttk.Frame(c3, style="Card.TFrame")
        g3.pack(fill="x")
        g3.columnconfigure(1, weight=1)
        for r, (label, var, pick) in enumerate((("Шаблон сборки", self.template, self.pick_template),
                                                ("Куда сохранять", self.root_dir, self.pick_root))):
            ttk.Label(g3, text=label, style="Field.TLabel").grid(row=r, column=0, sticky="w", padx=(0, 10), pady=3)
            ttk.Entry(g3, textvariable=var).grid(row=r, column=1, sticky="ew")
            ttk.Button(g3, text="…", width=3, command=pick).grid(row=r, column=2, padx=(6, 0))

        c5 = self.card(tab, "Склад BLAGO")
        g5 = ttk.Frame(c5, style="Card.TFrame")
        g5.pack(fill="x")
        g5.columnconfigure(1, weight=1)
        ttk.Label(g5, text="Адрес склада", style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=3)
        ttk.Entry(g5, textvariable=self.sklad_url).grid(row=0, column=1, sticky="ew")
        ttk.Label(g5, text="Ключ", style="Field.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=3)
        self.sklad_key_entry = ttk.Entry(g5, textvariable=self.sklad_key, show="•")
        self.sklad_key_entry.grid(row=1, column=1, sticky="ew")
        b5 = ttk.Frame(c5, style="Card.TFrame")
        b5.pack(fill="x", pady=(8, 0))
        self.button(b5, "Проверить связь", self.test_sklad)
        self.button(b5, "Отправить на склад неотправленное", self.flush_sklad, padx=(8, 0))
        self.sklad_status = ttk.Label(b5, text="", style="Status.TLabel")
        self.sklad_status.pack(side="left", padx=(10, 0))
        ttk.Label(c5, text="Ключ создаётся на складе: Администратор → Ключи API. Пусто — программа "
                           "работает без склада, как раньше.", style="Info.TLabel").pack(anchor="w", pady=(8, 0))

        c6 = self.card(tab, "СДЭК")
        g6 = ttk.Frame(c6, style="Card.TFrame")
        g6.pack(fill="x")
        g6.columnconfigure(1, weight=1)
        ttk.Label(g6, text="Account", style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=3)
        ttk.Entry(g6, textvariable=self.cdek_account).grid(row=0, column=1, sticky="ew")
        ttk.Label(g6, text="Secure", style="Field.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=3)
        ttk.Entry(g6, textvariable=self.cdek_secure, show="•").grid(row=1, column=1, sticky="ew")
        b6 = ttk.Frame(c6, style="Card.TFrame")
        b6.pack(fill="x", pady=(8, 0))
        self.button(b6, "Проверить ключи", self.test_cdek)
        self.cdek_status = ttk.Label(b6, text="", style="Status.TLabel")
        self.cdek_status.pack(side="left", padx=(10, 0))
        ttk.Label(c6, text="Ключи — в кабинете СДЭК: Интеграция → Ключи API. По ним программа качает ШК СДЭК.",
                  style="Info.TLabel").pack(anchor="w", pady=(8, 0))

        c4 = self.card(tab, "Этикетки Ozon")
        g4 = ttk.Frame(c4, style="Card.TFrame")
        g4.pack(fill="x")
        for text, var, lo, hi in (("Искать заказы за, дней", self.days, 1, 90),
                                  ("Отступ подписи", self.top_margin, 8, 40),
                                  ("Шрифт подписи", self.font_size, 6, 24)):
            ttk.Label(g4, text=text, style="Field.TLabel").pack(side="left")
            ttk.Spinbox(g4, from_=lo, to=hi, textvariable=var, width=4).pack(side="left", padx=(8, 26))

        c5 = self.card(tab, "Пропуска", side=lambda h: ttk.Checkbutton(
            h, text="делать пропуска", variable=self.pass_on).pack(side="right"))
        g5 = ttk.Frame(c5, style="Card.TFrame")
        g5.pack(fill="x")
        g5.columnconfigure(1, weight=1)
        g5.columnconfigure(3, weight=1)
        for r, c, (label, var) in ((0, 0, ("Водитель", self.pass_driver)), (0, 2, ("Телефон", self.pass_phone)),
                                   (1, 0, ("Машины", self.pass_cars))):
            ttk.Label(g5, text=label, style="Field.TLabel").grid(row=r, column=c, sticky="w",
                                                                 padx=(0 if c == 0 else 18, 10), pady=3)
            ttk.Entry(g5, textvariable=var).grid(row=r, column=c + 1, sticky="ew")

    # ------------------------------------------------------------ папки и настройки
    @staticmethod
    def default_root():
        desk = Path.home() / "Desktop"
        return desk if desk.exists() else BASE_DIR

    DIR_ALIASES = {"озон": ("озон1", "озонутро"), "кгт1": ("кгт",), "вб": ("wb", "вб1")}

    @staticmethod
    def _dir_key(name):
        return "".join(str(name).lower().replace("ё", "е").split())

    def find_dir(self, base, sub):
        """Уже созданная папка с таким названием (без учёта регистра и пробелов: «озон2» = «Озон 2»;
        «Озон 1» = «Озон», «КГТ» = «КГТ 1», «WB» = «ВБ») или None."""
        key = self._dir_key(sub)
        keys = {key, *self.DIR_ALIASES.get(key, ())}
        try:
            for d in sorted(Path(base).iterdir()):
                if d.is_dir() and self._dir_key(d.name) in keys:
                    return d
        except OSError:
            pass
        return None

    def day_dir(self, sub):
        """Главная папка / <sub>. Если такая папка уже создана (хоть «озон 2», хоть «Озон2») — она."""
        base = Path(self.root_dir.get().strip() or self.default_root())
        if self.by_day.get():
            d = today_msk()
            base = base / f"{MONTHS[d.month - 1]} {d.year}" / f"{d:%d.%m}"
        return self.find_dir(base, sub) or base / sub

    def place_reports(self, base, files, log):
        """Утренние отчёты Озон/КГТ/ВБ/СГТ — в папки «Озон», «КГТ 1», «ВБ», если такие папки уже созданы."""
        for key, sub in (("ozon", "Озон"), ("kgt", "КГТ 1"), ("wb", "ВБ"), ("sgt", "ВБ")):
            src = files.get(key)
            dst_dir = self.find_dir(base, sub)
            if not src or not dst_dir or Path(src).parent == dst_dir:
                continue
            dst = dst_dir / Path(src).name
            try:
                os.replace(src, dst)
                files[key] = dst
                log(f"   {Path(src).name} → папка «{dst_dir.name}»")
            except OSError as e:
                log(f"⚠ {Path(src).name} не перенёс в «{dst_dir.name}»: {e}")
        return files

    def load_settings(self):
        if CONFIG.exists():
            try:
                for name, cfg in json.loads(CONFIG.read_text(encoding="utf-8")).items():
                    self.add_cabinet(name, str(cfg.get("client_id", "")), cfg.get("api_key", ""))
            except ValueError:
                self.log(f"✖ Файл {CONFIG.name} повреждён, ключи не загружены")
        while len(self.cab_rows) < 3:
            self.add_cabinet()
        if WB_CONFIG.exists():
            try:
                self.wb_token.set(json.loads(WB_CONFIG.read_text(encoding="utf-8")).get("token", ""))
            except ValueError:
                pass
        if SKLAD_CONFIG.exists():
            try:
                cfg = json.loads(SKLAD_CONFIG.read_text(encoding="utf-8"))
                self.sklad_url.set(cfg.get("url") or self.sklad_url.get())
                self.sklad_key.set(cfg.get("key", ""))
            except ValueError:
                pass
        SKLAD.update(url=self.sklad_url.get(), key=self.sklad_key.get())
        if CDEK_CONFIG.exists():
            try:
                cfg = json.loads(CDEK_CONFIG.read_text(encoding="utf-8"))
                self.cdek_account.set(cfg.get("account", ""))
                self.cdek_secure.set(cfg.get("secure", ""))
            except ValueError:
                pass
        CDEK.update(account=self.cdek_account.get().strip(), secure=self.cdek_secure.get().strip())
        data = {}
        if MORNING_CONFIG.exists():
            try:
                data = json.loads(MORNING_CONFIG.read_text(encoding="utf-8"))
            except ValueError:
                pass
        self.template.set(data.get("template") or (str(BASE_DIR / "BLAGO.xlsx")
                                                   if (BASE_DIR / "BLAGO.xlsx").exists() else ""))
        root = data.get("root")
        if not root and data.get("out"):
            # прежняя версия: «…/Октябрь 2026/28.10/Сборка» -> рабочая папка «…»
            import re
            p = Path(data["out"])
            if len(p.parts) > 3 and re.fullmatch(r"\d\d\.\d\d", p.parts[-2]) and re.search(r"\d{4}$", p.parts[-3]):
                root = str(p.parents[2])
        if root:
            self.root_dir.set(root)
        self.by_day.set(False)
        for key, var in (("days", self.days), ("margin", self.top_margin), ("font", self.font_size)):
            if key in data:
                var.set(data[key])
        self.autoship.set(data.get("autoship", True))
        pc = {**PASS_DEFAULT, **(data.get("passes") or {})}
        self.pass_on.set(bool(pc["on"]))
        self.pass_driver.set(pc["driver"])
        self.pass_phone.set(pc["phone"])
        self.pass_cars.set(str(pc["cars"]).replace("\n", "; "))

    def schedule_save(self):
        if self._save_job:
            self.after_cancel(self._save_job)
        self._save_job = self.after(700, self.save_settings)

    def save_settings(self):
        self._save_job = None
        try:
            CONFIG.write_text(json.dumps(self.cabinets(), ensure_ascii=False, indent=2), encoding="utf-8")
            WB_CONFIG.write_text(json.dumps({"token": self.wb_token.get().strip()}), encoding="utf-8")
            SKLAD.update(url=self.sklad_url.get().strip(), key=self.sklad_key.get().strip())
            SKLAD_CONFIG.write_text(json.dumps(SKLAD), encoding="utf-8")
            CDEK.update(account=self.cdek_account.get().strip(), secure=self.cdek_secure.get().strip())
            CDEK_CONFIG.write_text(json.dumps(CDEK), encoding="utf-8")
            MORNING_CONFIG.write_text(json.dumps({
                "template": self.template.get().strip(), "root": self.root_dir.get().strip(),
                "by_day": bool(self.by_day.get()), "days": int(self.num(self.days, 14)),
                "margin": self.num(self.top_margin, DEFAULT_TOP_MARGIN_PT),
                "font": self.num(self.font_size, DEFAULT_FONT_SIZE), "autoship": bool(self.autoship.get()),
                "passes": self.pass_cfg(),
            }, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError as e:
            self.log(f"✖ Не удалось сохранить настройки: {e}")
        self.refresh_links()

    def pass_cfg(self):
        return {"on": bool(self.pass_on.get()), "driver": " ".join(self.pass_driver.get().split()),
                "phone": self.pass_phone.get().strip(), "cars": self.pass_cars.get().strip()}

    def passes_bg(self, func, *args):
        """Пропуска идут отдельно от кнопок: WB может ждать между пропусками по 10 минут."""
        if self._pass_running.get(func.__name__):
            self.log("… пропуска уже делаются — смотри журнал")
            return
        self._pass_running[func.__name__] = True
        cfg = self.pass_cfg()

        def work():
            try:
                func(*args, cfg, lambda s: self.q.put(("log", s)))
            except Exception as e:  # noqa: BLE001
                self.q.put(("log", f"✖ Пропуска: {e}"))
            self._pass_running[func.__name__] = False
        threading.Thread(target=work, daemon=True).start()

    def start_passes(self):
        cabinets = self.cabinets()
        if not cabinets:
            messagebox.showwarning(APP_NAME, "Заполни ключи Ozon во вкладке «Настройки».")
            return
        self.passes_bg(ozon_passes, cabinets)

    def on_close(self):
        self.save_settings()
        self.destroy()

    def pick_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.template.set(path)

    def pick_root(self):
        """Рабочая папка — одна на всё; внутри программа сама делает «Октябрь 2026 / 29.10 / …»."""
        path = filedialog.askdirectory(title="Где хранить отчёты и этикетки",
                                       initialdir=self.root_dir.get() or str(self.default_root()))
        if path:
            self.root_dir.set(path)
            self.save_settings()

    def open_out(self, path):
        p = Path(path)
        if not p.exists():
            p.mkdir(parents=True, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(p)  # noqa: S606
        else:
            subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(p)])

    def check_template(self):
        t = self.template.get().strip()
        if t and not Path(t).exists():
            messagebox.showwarning(APP_NAME, f"Не найден шаблон сборки:\n{t}\n\nВыбери его во вкладке «Настройки».")
            return None
        if not t:
            messagebox.showwarning(APP_NAME, "Выбери шаблон сборки во вкладке «Настройки».")
            return None
        return t

    # ------------------------------------------------------------ копирование / вставка
    def setup_clipboard(self):
        """Ctrl+C/V/X/A при русской раскладке и меню по правой кнопке мыши."""
        codes = {86: "<<Paste>>", 67: "<<Copy>>", 88: "<<Cut>>", 65: "<<SelectAll>>"}
        latin = {"v", "c", "x", "a"}

        def select_all(w):
            if isinstance(w, tk.Text):
                w.tag_add("sel", "1.0", "end-1c")
            else:
                w.select_range(0, "end")
                w.icursor("end")

        def on_ctrl_key(event):
            if sys.platform != "win32" or event.keysym.lower() in latin:
                return None
            action = codes.get(event.keycode)
            if not action:
                return None
            if action == "<<SelectAll>>":
                select_all(event.widget)
            else:
                event.widget.event_generate(action)
            return "break"

        def on_select_all(event):
            select_all(event.widget)
            return "break"

        menu = tk.Menu(self, tearoff=0, bg=C["card"], fg=C["text"], bd=0,
                       activebackground=C["accent"], activeforeground=C["on_accent"])

        def popup(event):
            w = event.widget
            w.focus_set()
            menu.delete(0, "end")
            readonly = isinstance(w, tk.Text) and str(w.cget("state")) == "disabled"
            if not readonly:
                menu.add_command(label="Вставить", command=lambda: w.event_generate("<<Paste>>"))
                menu.add_command(label="Вырезать", command=lambda: w.event_generate("<<Cut>>"))
            menu.add_command(label="Копировать", command=lambda: w.event_generate("<<Copy>>"))
            menu.add_separator()
            menu.add_command(label="Выделить всё", command=lambda: select_all(w))
            menu.tk_popup(event.x_root, event.y_root)

        for cls in ("TEntry", "Entry", "Text", "TCombobox", "TSpinbox"):
            self.bind_class(cls, "<Control-KeyPress>", on_ctrl_key, add="+")
            self.bind_class(cls, "<Control-a>", on_select_all)
            self.bind_class(cls, "<Control-A>", on_select_all)
            self.bind_class(cls, "<Button-3>", popup)

    # ------------------------------------------------------------ кабинеты Ozon
    def add_cabinet(self, name="", cid="", key=""):
        if not hasattr(self, "show_keys"):
            self.show_keys = tk.BooleanVar(value=False)
        r = len(self.cab_rows) + 1
        v = (tk.StringVar(value=name or f"Озон {r}"), tk.StringVar(value=cid), tk.StringVar(value=key))
        e_name = ttk.Entry(self.cab_box, textvariable=v[0], width=14)
        e_cid = ttk.Entry(self.cab_box, textvariable=v[1], width=12)
        e_key = ttk.Entry(self.cab_box, textvariable=v[2], show="" if self.show_keys.get() else "•")
        status = ttk.Label(self.cab_box, text="", width=22, style="Status.TLabel")
        e_name.grid(row=r, column=0, padx=(0, 8), pady=3)
        e_cid.grid(row=r, column=1, padx=(0, 8))
        e_key.grid(row=r, column=2, padx=(0, 10), sticky="ew")
        status.grid(row=r, column=3, sticky="w")
        for var in v:
            var.trace_add("write", lambda *_: self.schedule_save())
        self.cab_rows.append((v, e_key, status))

    def toggle_keys(self):
        for _, e_key, _ in self.cab_rows:
            e_key.configure(show="" if self.show_keys.get() else "•")

    def cabinets(self):
        """Заполненные кабинеты в порядке строк: {название: {client_id, api_key}}."""
        result = {}
        for (name, cid, key), _, _ in self.cab_rows:
            if cid.get().strip() and key.get().strip():
                result[name.get().strip() or cid.get().strip()] = {
                    "client_id": cid.get().strip(), "api_key": key.get().strip()}
        return result

    def test_keys(self):
        if self.busy:
            return
        rows = [(v[0].get(), v[1].get().strip(), v[2].get().strip(), st) for v, _, st in self.cab_rows]
        self.save_settings()

        def one(name, cid, key, st):
            if not cid or not key:
                self.q.put(("status", st, "не заполнен", C["muted"]))
                return
            self.q.put(("status", st, "проверяю…", C["muted"]))
            try:
                check_key(headers_for({"client_id": cid, "api_key": key}))
                self.q.put(("status", st, "✔ работает", C["ok"]))
            except Exception as e:  # noqa: BLE001
                self.q.put(("status", st, "✖ ошибка", C["err"]))
                self.q.put(("log", f"✖ {name}: {e}"))

        def work():
            with ThreadPoolExecutor(max_workers=max(1, len(rows))) as pool:
                list(pool.map(lambda r: one(*r), rows))
            self.q.put(("done",))

        self.run(work)

    def test_sklad(self):
        self.save_settings()
        if not sklad_on():
            self.sklad_status.configure(text="впиши адрес и ключ", foreground=C["muted"])
            return
        self.sklad_status.configure(text="проверяю…", foreground=C["muted"])

        def work():
            try:
                sklad_check()
                self.q.put(("status", self.sklad_status, "✔ склад на связи", C["ok"]))
            except (RuntimeError, requests.RequestException) as e:
                self.q.put(("status", self.sklad_status, f"✖ {e}", C["err"]))

        threading.Thread(target=work, daemon=True).start()

    def test_cdek(self):
        self.save_settings()
        if not cdek_on():
            self.cdek_status.configure(text="впиши Account и Secure", foreground=C["muted"])
            return
        self.cdek_status.configure(text="проверяю…", foreground=C["muted"])

        def work():
            try:
                cdek_token()
                self.q.put(("status", self.cdek_status, "✔ СДЭК пустил по ключам", C["ok"]))
            except (RuntimeError, requests.RequestException) as e:
                self.q.put(("status", self.cdek_status, f"✖ {e}", C["err"]))

        threading.Thread(target=work, daemon=True).start()

    def send_morning(self):
        """«Отправить на склад»: заказы дня — только после того, как ты проверил отчёты."""
        if self.busy:
            return
        data = morning_load()
        if not data:
            messagebox.showinfo(APP_NAME, "Нечего отправлять: сначала скачай отчёты.")
            return
        problems = morning_problems(data)
        text = "Отправить на склад?\n\n" + morning_text(data).replace(" | ", "\n")
        if problems:
            text += ("\n\nВНИМАНИЕ:\n" + "\n".join(problems)
                     + "\n\nНе загрузившиеся кабинеты на складе не трогаются — там останутся их заказы со вчера."
                     + "\nЛучше нажать «Утро» ещё раз. Всё равно отправить?")
        if not messagebox.askyesno(APP_NAME, text, icon="warning" if problems else "question"):
            return
        self.headline("ОТПРАВКА ЗАКАЗОВ ДНЯ НА СКЛАД")

        def work():
            morning_send(lambda t: self.q.put(("log", t)))
            self.q.put(("done",))

        self.run(work)

    def flush_sklad(self):
        if self.busy:
            return
        self.save_settings()
        self.headline("ОТПРАВКА НА СКЛАД")

        def work():
            left = len(_pending_load())
            if not left:
                self.q.put(("log", "✔ Всё уже на складе, отправлять нечего"))
            sklad_flush(lambda t: self.q.put(("log", t)))
            self.q.put(("done",))

        self.run(work)

    def test_wb(self):
        if self.busy:
            return
        token = self.wb_token.get().strip()
        if not token:
            self.wb_status.configure(text="не заполнен", foreground=C["muted"])
            return
        self.save_settings()
        self.wb_status.configure(text="проверяю…", foreground=C["muted"])

        def work():
            try:
                wb_check(token)
                self.q.put(("status", self.wb_status, "✔ работает", C["ok"]))
            except Exception as e:  # noqa: BLE001
                self.q.put(("status", self.wb_status, "✖ ошибка", C["err"]))
                self.q.put(("log", f"✖ WB: {e}"))
            self.q.put(("done",))

        self.run(work)

    # ------------------------------------------------------------ отчёт Ozon
    def set_excel(self, path):
        try:
            names = sheet_names(path)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Не удалось открыть файл", str(e))
            return
        self.excel.set(path)
        self.excel_name.configure(text=Path(path).name)
        self.sheet_cb["values"] = names
        self.sheet.set(names[0])
        if len(names) > 1:
            self.sheet_box.pack(side="right", padx=(0, 14))
        else:
            self.sheet_box.pack_forget()
        self.count_report()

    def pick_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if path:
            self.set_excel(path)

    def background(self, func, done):
        """Чтение файлов — в фоне, чтобы окно не подвисало; результат — в done(...) в окне."""
        def work():
            try:
                res, err = func(), None
            except Exception as e:  # noqa: BLE001
                res, err = None, e
            self.q.put(("call", lambda: done(res, err)))
        threading.Thread(target=work, daemon=True).start()

    def count_report(self):
        path, sheet = self.excel.get(), self.sheet.get()
        self.report_info.configure(text="читаю отчёт…", foreground=C["muted"])

        def done(text, err):
            if path != self.excel.get():
                return  # пока читали, выбрали другой файл
            if err:
                self.report_info.configure(text=f"Не получилось прочитать: {err}", foreground=C["err"])
            else:
                self.report_info.configure(text=text, foreground=C["text"])

        self.background(lambda: ozon_report_summary(path, sheet), done)

    # ------------------------------------------------------------ отчёты WB
    def clear_wb(self):
        self.set_wb_files([])

    def set_wb_files(self, files):
        self.wb_files = list(files)
        self.wb_files_label.configure(text="   ·   ".join(Path(p).stem for p in files) or "не выбраны")
        if not files:
            self.wb_info.configure(text="Отчёты подставляются сами после «Сборки».")
            return
        self.wb_info.configure(text="читаю отчёты…")
        snapshot = list(files)

        def done(text, err):
            if snapshot == self.wb_files:
                self.wb_info.configure(text=text if not err else f"Не получилось прочитать: {err}")

        self.background(lambda: wb_files_summary(snapshot), done)

    def pick_wb(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if paths:
            self.set_wb_files(self.wb_files + [p for p in paths if p not in self.wb_files])

    # ------------------------------------------------------------ запуск работ
    def run(self, work):
        self.set_busy(True)
        threading.Thread(target=work, daemon=True).start()

    def headline(self, text):
        self.bar["value"] = 0
        self.log("")
        self.log(f"=== {text} ===")

    def start_morning(self):
        if self.busy:
            return
        template = self.check_template()
        if not template:
            return
        cabinets, token = self.cabinets(), self.wb_token.get().strip()
        if not cabinets and not token:
            messagebox.showwarning(APP_NAME, "Заполни ключи Ozon и токен WB во вкладке «Настройки».")
            return
        self.save_settings()
        out, days = self.day_dir("Сборка"), int(self.num(self.days, 14))
        base = Path(self.root_dir.get().strip() or self.default_root())
        d = today_msk()
        self.headline(f"УТРЕННЯЯ СБОРКА · {d:%d.%m}")
        self.mo_status.configure(text="скачиваю…", foreground=C["muted"])

        def work():
            try:
                files = morning(cabinets, token, template, "", out, days,
                                log=lambda s: self.q.put(("log", s)),
                                progress=lambda a, b: self.q.put(("progress", a, b)))
                files = self.place_reports(base, files, lambda s: self.q.put(("log", s)))
                self.q.put(("call", lambda: self.after_morning(files)))
            except Exception as e:  # noqa: BLE001
                self.q.put(("log", f"✖ Ошибка: {e}"))
                self.q.put(("status", self.mo_status, "✖ ошибка — см. журнал", C["err"]))
            self.q.put(("done",))

        self.run(work)

    def after_morning(self, files):
        if not files.get("ozon") and self.cabinets():
            self.mo_status.configure(text="✖ Озон не скачался — нажми «Скачать отчёты и сборку» ещё раз.",
                                     foreground=C["err"])
            wb = [str(files[k]) for k in ("wb", "sgt") if files.get(k)]
            if wb:
                self.set_wb_files(wb)
            return
        if files.get("ozon"):
            self.set_excel(str(files["ozon"]))
            self.footer.set("Озон утро")
        wb = [str(files[k]) for k in ("wb", "sgt") if files.get(k)]
        if wb:
            self.set_wb_files(wb)
        self.refresh_links()
        t = datetime.now(MSK)
        self.mo_status.configure(text=f"✔ Готово в {t:%H:%M}. Отчёты уже стоят на вкладках Ozon и Wildberries.",
                                 foreground=C["ok"])

    def start_cdek(self):
        if self.busy:
            return
        template = self.check_template()
        if not template:
            return
        cdek = filedialog.askopenfilename(title="Файл СДЭК", filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not cdek:
            return
        try:
            orders = read_cdek_orders(cdek)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror(APP_NAME, f"Не удалось прочитать файл СДЭК: {e}")
            return
        out = self.day_dir("Сборка")
        self.headline("СДЭК В СБОРКУ")

        def work():
            log = lambda s: self.q.put(("log", s))  # noqa: E731
            try:
                # файл с трэками: склад, листы, ШК и акты — на странице «СДЭК»
                add_cdek(template, cdek, out, log=log, send_sklad=not orders)
                if orders:
                    log("│ Листы сборки, ШК и акты СДЭК — на странице «СДЭК» слева.")
            except PermissionError:
                self.q.put(("log", "✖ Файл сборки открыт в Excel — закрой его и нажми ещё раз"))
            except Exception as e:  # noqa: BLE001
                self.q.put(("log", f"✖ Ошибка: {e}"))
            self.q.put(("done",))

        self.run(work)

    def start_wave(self):
        if self.busy:
            return
        cabinets = self.cabinets()
        if not cabinets:
            messagebox.showwarning(APP_NAME, "Заполни ключи Ozon во вкладке «Настройки».")
            return
        template = self.check_template()
        if not template:
            return
        wave = self.wave_name()
        kgt = "кгт" in wave.lower()
        out, days = self.day_dir(wave), int(self.num(self.days, 14))
        self.save_settings()
        self.headline(f"{wave.upper()} · {'КГТ' if kgt else 'МГТ'} · только «ожидают сборки»")
        self.wave_status.configure(text="скачиваю…", foreground=C["muted"])

        def work():
            try:
                files = ozon_wave(cabinets, template, out, days, wave, kgt,
                                  log=lambda s: self.q.put(("log", s)),
                                  progress=lambda a, b: self.q.put(("progress", a, b)))
                self.q.put(("call", lambda: self.after_wave(wave, files)))
            except Exception as e:  # noqa: BLE001
                self.q.put(("log", f"✖ Ошибка: {e}"))
                self.q.put(("status", self.wave_status, "✖ ошибка — см. журнал", C["err"]))
            self.q.put(("done",))

        self.run(work)

    def after_wave(self, wave, files):
        report = files.get("kgt") or files.get("ozon")
        if report:
            self.set_excel(str(report))
        self.footer.set(wave)
        self.refresh_links()
        t = datetime.now(MSK)
        self.wave_status.configure(text=f"✔ {wave} готов в {t:%H:%M}. Отчёт уже стоит на вкладке Ozon.",
                                   foreground=C["ok"])

    def start_wb(self, mode):
        if self.busy:
            return
        token = self.wb_token.get().strip()
        if not token:
            messagebox.showwarning(APP_NAME, "Вставь токен WB во вкладке «Настройки».")
            return
        if mode != "stickers" and not self.wb_files:
            messagebox.showwarning(APP_NAME, "Добавь отчёты ВБ и СГТ.")
            return
        self.save_settings()
        out = self.day_dir("ВБ")
        self.headline({"run": "WB · ПОСТАВКИ И СТИКЕРЫ", "check": "WB · ПРОВЕРКА",
                       "stickers": "WB · СТИКЕРЫ ОТКРЫТЫХ ПОСТАВОК"}[mode])
        files = list(self.wb_files)

        def work():
            try:
                process_wb(files, token, out, log=lambda s: self.q.put(("log", s)),
                           progress=lambda a, b: self.q.put(("progress", a, b)),
                           mode=mode, confirm=self.ask)
            except Exception as e:  # noqa: BLE001
                self.q.put(("log", f"✖ Ошибка: {e}"))
            self.q.put(("done",))

        self.run(work)

    def num(self, var, default):
        try:
            return float(var.get())
        except (tk.TclError, ValueError):
            return default

    def start(self, mode):
        """Вкладка Ozon. mode: run — скачать и собрать, check — проверка, merge — пересобрать ленты."""
        if self.busy:
            return
        font, margin = self.num(self.font_size, DEFAULT_FONT_SIZE), self.num(self.top_margin, DEFAULT_TOP_MARGIN_PT)
        out = self.day_dir(self.ozon_sub())
        if mode == "merge":
            footer = self.footer.get()
            src = self.last_ozon_out
            note = ""
            if src and Path(src) != out and Path(src).exists():
                # надпись поменяли после скачивания — переименуем папку под новую надпись
                if not out.exists():
                    try:
                        Path(src).rename(out)
                        note = f"Папка «{Path(src).name}» переименована в «{out.name}»"
                    except OSError as e:
                        out, note = Path(src), f"⚠ Не удалось переименовать папку «{Path(src).name}»: {e}"
                else:
                    out = Path(src)
                    note = f"⚠ Папка «{self.ozon_sub()}» уже есть — пересобираю «{out.name}» с новой надписью"
            self.last_ozon_out = out
            self.refresh_links()
            self.headline(f"ПЕРЕСБОРКА ЛЕНТ · {' '.join(footer.split()) or 'без надписи'}")

            def work():
                try:
                    if note:
                        self.q.put(("log", note))
                    merge_all(out, font, margin, lambda s: self.q.put(("log", s)), footer)
                except Exception as e:  # noqa: BLE001
                    self.q.put(("log", f"✖ Ошибка: {e}"))
                self.q.put(("done",))
        else:
            cabinets = self.cabinets()
            if not cabinets:
                messagebox.showwarning(APP_NAME, "Заполни ключи Ozon во вкладке «Настройки».")
                return
            if not self.excel.get():
                messagebox.showwarning(APP_NAME, "Выбери отчёт.")
                return
            self.save_settings()
            self.headline(f"OZON · {'СКАЧИВАНИЕ И СБОРКА' if mode == 'run' else 'ПРОВЕРКА'} · {self.ozon_sub()}")
            args = (self.excel.get(), self.sheet.get(), cabinets, int(self.num(self.days, 14)), out)
            autoship, footer = bool(self.autoship.get()), self.footer.get()
            pass_cfg = self.pass_cfg() if self.pass_on.get() else None
            if mode == "run":
                self.last_ozon_out = out
    
            def work():
                try:
                    form = process(*args, log=lambda s: self.q.put(("log", s)),
                            progress=lambda a, b: self.q.put(("progress", a, b)),
                            download=mode == "run", font_size=font, top_margin=margin,
                            autoship=autoship, confirm=self.ask, footer=footer)
                    if mode == "run" and pass_cfg and form:
                        self.q.put(("log", "\n=== ОТГРУЗКА И ПРОПУСКА ==="))
                        ozon_passes(cabinets, pass_cfg, lambda s: self.q.put(("log", s)),
                                    form=form, confirm=self.ask)
                except Exception as e:  # noqa: BLE001
                    self.q.put(("log", f"✖ Ошибка: {e}"))
                self.q.put(("done",))

        self.run(work)

    def ask(self, text):
        """Вопрос «да/нет» из рабочего потока (сам диалог показывается в окне)."""
        done, answer = threading.Event(), {}
        self.q.put(("ask", text, done, answer))
        done.wait()
        return answer.get("ok", False)

    def set_busy(self, busy):
        self.busy = busy
        self.state_pill.configure(text="●  работаю…" if busy else "●  готова к работе",
                                  fg=C["accent"] if busy else C["ok"])
        if busy:
            # пока не пришёл первый шаг — бегущая полоса, чтобы было видно, что программа живая
            self.bar.configure(mode="indeterminate", value=0)
            self.bar.start(15)
            self.bar.pack(fill="x", padx=18, pady=(10, 0), before=self._log_frame)
        else:
            self.bar.stop()
            self.bar.pack_forget()
        for b in self.buttons:
            b.configure(state="disabled" if busy else "normal")
        for b in getattr(self, "wave_chips", {}).values():
            b.configure(state="disabled" if busy else "normal")

    # ------------------------------------------------------------ журнал
    def clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def log(self, text):
        s = text.lstrip()
        tag = ("big" if s.startswith(("✔ ВСЕГО ШК", "✔ СКАЧАНО ШК", "✔ ГОТОВО")) else
               "ok" if s.startswith("✔") else "err" if s.startswith("✖") else
               "head" if s.startswith("===") else "warn" if s.startswith("⚠") else
               "warnbody" if s.startswith("│") else "yellow" if s.startswith("●") else
               "muted" if s.startswith("…") else "")
        if tag == "head":
            text = s.strip("= ").strip()
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n", tag)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                if m[0] == "log":
                    self.log(m[1])
                elif m[0] == "progress":
                    if str(self.bar.cget("mode")) == "indeterminate":
                        self.bar.stop()
                        self.bar.configure(mode="determinate")
                    self.bar["maximum"], self.bar["value"] = max(m[2], 1), m[1]
                elif m[0] == "status":
                    m[1].configure(text=m[2], foreground=m[3])
                elif m[0] == "ask":
                    m[3]["ok"] = messagebox.askyesno(APP_NAME, m[1])
                    m[2].set()
                elif m[0] == "call":
                    m[1]()
                elif m[0] == "done":
                    self.set_busy(False)
        except queue.Empty:
            pass
        self.after(100, self.poll)


if __name__ == "__main__":
    App().mainloop()
