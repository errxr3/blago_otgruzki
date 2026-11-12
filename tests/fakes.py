"""Подделки Ozon Seller API и WB API для тестов и построители отчётов.

Настоящая сеть не используется: ответы отдаёт `responses`. Каждая подделка записывает
все пришедшие запросы, чтобы тест мог проверить, что программа ничего лишнего не вызвала.
"""

import base64
import io
import json
import re
from datetime import datetime

import openpyxl
import responses
from openpyxl.styles import PatternFill
from reportlab.pdfgen import canvas

import blago_otgruzki as b

RED = PatternFill("solid", fgColor="FFFF0000")
YELLOW = PatternFill("solid", fgColor="FFFFFF00")
OZON_HEADER = ["Номер отправления", "Принят в обработку", "Артикул", "Количество"]
WB_HEADER = ["№ задания", "Артикул продавца", "Склад продавца", "Доставить на склад", "к"]

# выдуманные ключи трёх кабинетов: первый — Z, второй — R, третий — IZ
CABINETS = {
    "Z": {"client_id": "1001", "api_key": "test-key-z"},
    "R": {"client_id": "1002", "api_key": "test-key-r"},
    "IZ": {"client_id": "1003", "api_key": "test-key-iz"},
}
CAB_BY_ID = {cfg["client_id"]: name for name, cfg in CABINETS.items()}


def label_pdf(numbers, size=(164, 113)):
    """Фейковые этикетки Ozon: по странице на отправление, на странице — его номер."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=size)
    for num in numbers:
        c.drawString(10, 60, str(num))
        c.showPage()
    c.save()
    return buf.getvalue()


def png_sticker():
    """Маленькая картинка-стикер WB в base64."""
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (58, 40), "white").save(buf, "PNG")
    return base64.b64encode(buf.getvalue()).decode()


def ozon_report(path, rows, red=(), yellow=None, red_rows=(), sheet="Озон"):
    """Отчёт Ozon как у владельца: список слева, сводная в G (артикул), H (сумма), I («едет»).
    rows — [(номер, время, артикул, шт)]; red — красные артикулы сводной;
    yellow — {артикул: число в I | None}; red_rows — номера, чья строка списка красная."""
    b.write_report(path, sheet, OZON_HEADER, [list(r) for r in rows], 3, "Сумма по полю Количество")
    wbk = openpyxl.load_workbook(path)
    ws = wbk.worksheets[0]
    for row in ws.iter_rows(min_row=2):
        if row[0].value in red_rows:
            row[0].fill = RED
        cell = row[6]  # столбец G
        if cell.value in red:
            cell.fill = RED
        if yellow and cell.value in yellow:
            cell.fill = YELLOW
            if yellow[cell.value] is not None:
                ws.cell(cell.row, 9, yellow[cell.value])  # столбец I
    wbk.save(path)
    return path


def wb_report(path, rows, warehouse=None, deliver=None, red=(), yellow=None):
    """Отчёт WB: rows — [(id, артикул, склад, «доставить на склад», 1)].
    warehouse / deliver — значения фильтров сводной («Склад продавца», «Доставить на склад»)."""
    b.write_report(path, "Сборочные задания", WB_HEADER, [list(r) for r in rows], 4, "Сумма по полю к")
    wbk = openpyxl.load_workbook(path)
    ws = wbk.worksheets[0]
    col = len(WB_HEADER) + 6  # правее сводной
    line = 1
    for label, value in (("Склад продавца", warehouse), ("Доставить на склад", deliver)):
        if value is not None:
            ws.cell(line, col, label)
            ws.cell(line, col + 1, value)
            line += 1
    for row in ws.iter_rows(min_row=2):
        cell = row[len(WB_HEADER) + 2]
        if cell.value in red:
            cell.fill = RED
        if yellow and cell.value in yellow:
            cell.fill = YELLOW
            ws.cell(cell.row, cell.column + 2, yellow[cell.value])
    wbk.save(path)
    return path


def posting(status, *products, warehouse="FBS Тверь", time="2026-11-09T07:00:00Z"):
    """Отправление в кабинете: products — (артикул, шт)."""
    return {"status": status, "warehouse": warehouse, "time": time,
            "products": [{"offer_id": art, "sku": abs(hash(art)) % 10**9, "quantity": qty}
                         for art, qty in products]}


class FakeOzon:
    """Кабинеты Ozon в памяти. cabinets[имя] = {номер: posting(...)}."""

    def __init__(self, mock):
        self.mock = mock
        self.cabinets = {name: {} for name in CABINETS}
        self.calls = []          # (кабинет, путь, тело)
        self.fail = {}           # путь -> (код, тело) | исключение | список таких ответов по очереди
        self.tasks = {}          # task_id -> номера
        self.label_bytes = None  # подмена файла этикеток (битый PDF)
        self.carriages = {}      # кабинет -> методы доставки для /v2/carriage/delivery/list
        self.created = []        # созданные отгрузки
        self.unprinted = set()   # номера, на которые Ozon не печатает этикетку
        mock.add_callback(responses.POST, re.compile(r"https://api-seller\.ozon\.ru/.*"), callback=self._post)
        mock.add_callback(responses.GET, re.compile(r"https://files\.test/.*"), callback=self._file)

    # ------------------------------------------------------------ что вызывали
    def paths(self, cab=None):
        return [p for c, p, _ in self.calls if cab is None or c == cab]

    def count(self, path, cab=None):
        return self.paths(cab).count(path)

    # ------------------------------------------------------------ ответы
    def _post(self, request):
        cab = CAB_BY_ID.get(request.headers.get("Client-Id"), "?")
        path = request.url.split("ozon.ru", 1)[1]
        body = json.loads(request.body or b"{}")
        self.calls.append((cab, path, body))
        forced = self.fail.get((cab, path), self.fail.get(path))
        if isinstance(forced, list):
            forced = forced.pop(0) if forced else None
        if isinstance(forced, Exception):
            raise forced
        if forced:
            return forced[0], {}, json.dumps(forced[1])
        handler = {
            "/v4/posting/fbs/list": self._list,
            "/v4/posting/fbs/ship": self._ship,
            "/v3/posting/fbs/package-label/create": self._label_create,
            "/v2/posting/fbs/package-label/get": self._label_get,
            "/v2/carriage/delivery/list": self._carriage_list,
            "/v1/carriage/create": self._carriage_create,
            "/v1/carriage/approve": lambda cab, body: (200, {}),
            "/v1/carriage/pass/create": lambda cab, body: (200, {"arrival_pass_ids": [1, 2]}),
            "/v1/carriage/get": lambda cab, body: (200, {"arrival_pass_ids": []}),
        }.get(path)
        if handler is None:
            return 404, {}, json.dumps({"message": f"нет такого метода в подделке: {path}"})
        code, data = handler(cab, body)
        return code, {}, json.dumps(data)

    def _list(self, cab, body):
        items = [{"posting_number": num, "status": p["status"], "products": p["products"],
                  "in_process_at": p["time"], "barcodes": {"upper_barcode": f"%{num}%"},
                  "delivery_method": {"warehouse": p["warehouse"], "warehouse_id": 500 + len(p["warehouse"]),
                                      "id": 900 + len(p["warehouse"])}}
                 for num, p in self.cabinets.get(cab, {}).items()]
        return 200, {"postings": items, "has_next": False, "cursor": ""}

    def _ship(self, cab, body):
        """Сборка: каждая упаковка — отдельное отправление с новым номером «заказ-1», «-3», «-5»…"""
        num = body["posting_number"]
        old = self.cabinets[cab].pop(num)
        by_sku = {p["sku"]: p["offer_id"] for p in old["products"]}
        order = num.rsplit("-", 1)[0]
        numbers, extra = [], []
        for i, package in enumerate(body["packages"]):
            new = f"{order}-{1 + 2 * i}"
            art = by_sku[package["products"][0]["product_id"]]
            self.cabinets[cab][new] = {**old, "status": b.READY,
                                       "products": [{"offer_id": art, "sku": package["products"][0]["product_id"],
                                                     "quantity": 1}]}
            numbers.append(new)
            extra.append({"posting_number": new, "products": [{"offer_id": art}]})
        return 200, {"result": numbers, "additional_data": extra}

    def _label_create(self, cab, body):
        task = len(self.tasks) + 1
        self.tasks[task] = body["posting_numbers"]
        return 200, {"result": {"tasks": [{"task_id": task, "task_type": "small_label"}]}}

    def _label_get(self, cab, body):
        numbers = self.tasks[body["task_id"]]
        skipped = [{"posting_number": n, "message": "Ozon не напечатал"} for n in numbers if n in self.unprinted]
        return 200, {"result": {"file_url": f"https://files.test/{body['task_id']}.pdf",
                                "status": {"code": "completed", "postings_count": len(numbers),
                                           "unprinted_postings": skipped}}}

    def _file(self, request):
        task = int(re.search(r"/(\d+)\.pdf", request.url).group(1))
        if self.label_bytes is not None:
            return 200, {}, self.label_bytes
        return 200, {}, label_pdf([n for n in self.tasks[task] if n not in self.unprinted])

    def _carriage_list(self, cab, body):
        return 200, {"methods": self.carriages.get(cab, []), "has_next": False}

    def _carriage_create(self, cab, body):
        cid = 7000 + len(self.created)
        self.created.append((cab, body["delivery_method_id"], cid))
        return 200, {"carriage_id": cid}


class FakeWB:
    """Кабинет WB в памяти."""

    def __init__(self, mock):
        self.mock = mock
        self.new = {}        # id -> задание (новые)
        self.status = {}     # id -> (supplierStatus, wbStatus) для не новых
        self.supplies = []   # {"id", "name", "done", "orders": [id]}
        self.calls = []      # (метод, путь, тело)
        self.fail = {}       # (метод, путь) -> (код, тело) | исключение | список ответов по очереди
        self.bad_sticker = set()   # id с битой картинкой
        self.history = {}    # id -> задание (для поиска артикула старых заданий)
        self.no_move = set()  # id, которые WB не даёт перенести в другую поставку
        base = r"https://marketplace-api\.wildberries\.ru"
        for method in (responses.GET, responses.POST, responses.PATCH, responses.PUT, responses.DELETE):
            mock.add_callback(method, re.compile(base + r"/.*"), callback=self._handle)

    def order(self, oid, article, cargo=1, warehouse_id=1, created="2026-11-10T05:00:00Z", **extra):
        self.new[str(oid)] = {"id": int(oid), "article": article, "cargoType": cargo,
                              "warehouseId": warehouse_id, "createdAt": created, **extra}

    def supply(self, sid, name, orders=(), done=False):
        self.supplies.append({"id": sid, "name": name, "done": done, "orders": [str(i) for i in orders]})

    def methods(self):
        return [(m, p) for m, p, _ in self.calls]

    def _handle(self, request):
        path = request.url.split("wildberries.ru", 1)[1].split("?")[0]
        body = json.loads(request.body) if request.body else {}
        self.calls.append((request.method, path, body))
        forced = self.fail.get((request.method, path))
        if isinstance(forced, list):
            forced = forced.pop(0) if forced else None
        if isinstance(forced, Exception):
            raise forced
        if forced:
            return forced[0], {}, json.dumps(forced[1])
        if request.method == "GET" and path == "/api/v3/orders/new":
            return 200, {}, json.dumps({"orders": list(self.new.values())})
        if request.method == "GET" and path == "/api/v3/orders":
            return 200, {}, json.dumps({"orders": list(self.history.values()) + list(self.new.values()),
                                        "next": 0})
        if request.method == "POST" and path == "/api/v3/orders/status":
            return 200, {}, json.dumps({"orders": [
                {"id": i, "supplierStatus": self.status[str(i)][0], "wbStatus": self.status[str(i)][1]}
                for i in body["orders"] if str(i) in self.status]})
        if request.method == "GET" and path == "/api/v3/supplies":
            return 200, {}, json.dumps({"supplies": [{k: s[k] for k in ("id", "name", "done")}
                                                     for s in self.supplies], "next": 0})
        if request.method == "POST" and path == "/api/v3/supplies":
            sid = f"WB-GI-{100 + len(self.supplies)}"
            self.supply(sid, body["name"])
            return 201, {}, json.dumps({"id": sid})
        m = re.fullmatch(r"/api/marketplace/v3/supplies/([^/]+)/orders", path)
        if request.method == "PATCH" and m:
            target = next(s for s in self.supplies if s["id"] == m.group(1))
            for i in body["orders"]:
                order = self.new.pop(str(i), None)
                if order is None:
                    # перенос из другой незакрытой поставки
                    source = next((s for s in self.supplies if str(i) in s["orders"] and not s["done"]), None)
                    if source is None or str(i) in self.no_move:
                        return 409, {}, json.dumps({"message": "задание нельзя перенести"})
                    source["orders"].remove(str(i))
                    target["orders"].append(str(i))
                    continue
                self.history[str(i)] = order
                self.status[str(i)] = ("confirm", "waiting")
                target["orders"].append(str(i))
            return 204, {}, ""
        m = re.fullmatch(r"/api/v3/supplies/([^/]+)", path)
        if request.method == "DELETE" and m:
            target = next(s for s in self.supplies if s["id"] == m.group(1))
            if target["orders"]:
                return 409, {}, json.dumps({"message": "поставка не пустая"})
            self.supplies.remove(target)
            return 204, {}, ""
        m = re.fullmatch(r"/api/marketplace/v3/supplies/([^/]+)/order-ids", path)
        if request.method == "GET" and m:
            target = next(s for s in self.supplies if s["id"] == m.group(1))
            return 200, {}, json.dumps({"orderIds": [int(i) for i in target["orders"]]})
        if request.method == "POST" and path == "/api/v3/orders/stickers":
            return 200, {}, json.dumps({"stickers": [
                {"orderId": i, "partA": str(i)[:3], "partB": str(i)[3:], "barcode": f"*WB{i}*",
                 "file": "AAAA" if str(i) in self.bad_sticker else png_sticker()} for i in body["orders"]]})
        return 404, {}, json.dumps({"message": f"нет такого метода в подделке: {request.method} {path}"})


def t(hour, minute=0, day=9):
    return datetime(2026, 11, day, hour, minute)
