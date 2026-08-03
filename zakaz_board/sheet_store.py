"""
sheet_store.py
==============
Google Таблица как база данных, теперь из трёх листов вместо одного:

  parts_pool     — все детали из загруженных экселей-спецификаций.
                    id | order | product | uzel | grade | thickness | custom_sheet |
                    code | name | size | qty_total | area
                    "uzel" — имя xlsm-файла (узла), из которого взята деталь.
                    Без статуса — это просто «что вообще существует и сколько нужно».

  raskroi        — сущность раскроя (одна физическая раскладка на лист):
                    id | name | grade | thickness | custom_sheet | status |
                    pdf_name | created
                    status: "Создан" -> "Выдан" -> "Вырезан"
                    (технолог создаёт -> учётчик выдаёт лист -> плазменщик режет)

  raskroi_items  — связка: какой раскрой каких деталей и сколько включает.
                    raskroy_id | part_id | qty
                    Одна деталь может быть раскидана по нескольким раскроям
                    (когда партия режется не за один лист).

Настройка (см. SETUP.md):
  1. Создать сервисный аккаунт в Google Cloud, включить Google Sheets API.
  2. Скачать json-ключ, положить его СОДЕРЖИМОЕ (весь json) в переменную
     окружения GOOGLE_CREDENTIALS_JSON (одной строкой).
  3. Расшарить саму Google Таблицу на email сервисного аккаунта
     (он выглядит как ...@...iam.gserviceaccount.com), с правами Редактор.
  4. Задать переменную окружения GOOGLE_SHEET_NAME = точное имя таблицы.
  5. В таблице должно быть три листа с именами и заголовками, как описано выше —
     см. SETUP.md, там есть точный список колонок для каждого.
"""

import os
import json
import uuid

import gspread
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

PARTS_SHEET = "parts_pool"
RASKROI_SHEET = "raskroi"
ITEMS_SHEET = "raskroi_items"

PARTS_COLUMNS = ["id", "order", "product", "uzel", "grade", "thickness", "custom_sheet",
                 "code", "name", "size", "qty_total", "area", "material_raw"]
RASKROI_COLUMNS = ["id", "name", "grade", "thickness", "custom_sheet", "status",
                    "pdf_name", "sheet_count", "created"]
ITEMS_COLUMNS = ["raskroy_id", "part_id", "qty"]

STATUS_FLOW = ["Создан", "Выдан", "Вырезан"]

_client_cache = None
_sheet_cache = None


def _client():
    global _client_cache
    if _client_cache is not None:
        return _client_cache
    creds_json = os.environ["GOOGLE_CREDENTIALS_JSON"]
    creds_dict = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    _client_cache = gspread.authorize(creds)
    return _client_cache


def _spreadsheet():
    global _sheet_cache
    if _sheet_cache is not None:
        return _sheet_cache
    sheet_name = os.environ["GOOGLE_SHEET_NAME"]
    _sheet_cache = _client().open(sheet_name)
    return _sheet_cache


def _ws(name):
    return _spreadsheet().worksheet(name)


# ------------------------------------------------------------- parts_pool ---

def get_parts_pool():
    return _ws(PARTS_SHEET).get_all_records()


def add_parts(rows):
    """
    rows: список словарей с ключами из PARTS_COLUMNS (без id — генерируется сам).
    Возвращает список добавленных id.
    """
    if not rows:
        return []
    ws = _ws(PARTS_SHEET)
    ids = []
    values = []
    for r in rows:
        new_id = uuid.uuid4().hex[:8]
        ids.append(new_id)
        values.append([r.get(c, "") if c != "id" else new_id for c in PARTS_COLUMNS])
    ws.append_rows(values)
    return ids


def find_part_by_natural_key(existing_rows, order, product, uzel, code):
    """Ищет уже существующую деталь по (заказ, изделие, узел, код) — для повторной заливки."""
    for r in existing_rows:
        if (str(r.get("order")) == str(order) and str(r.get("product")) == str(product)
                and str(r.get("uzel")) == str(uzel) and str(r.get("code")) == str(code)):
            return r
    return None


# --------------------------------------------------------------- raskroi ---

def get_raskroi():
    return _ws(RASKROI_SHEET).get_all_records()


def create_raskroy(name, grade, thickness, custom_sheet=""):
    ws = _ws(RASKROI_SHEET)
    new_id = uuid.uuid4().hex[:8]
    import datetime
    row = [new_id, name, grade, thickness, custom_sheet, "Создан", "", "1",
           datetime.date.today().isoformat()]
    ws.append_row(row)
    return new_id


def update_raskroy_status(raskroy_id, status):
    ws = _ws(RASKROI_SHEET)
    ids = ws.col_values(1)
    if raskroy_id not in ids:
        return False
    row = ids.index(raskroy_id) + 1
    col = RASKROI_COLUMNS.index("status") + 1
    ws.update_cell(row, col, status)
    return True


def attach_pdf(raskroy_id, pdf_name, sheet_count=None):
    ws = _ws(RASKROI_SHEET)
    ids = ws.col_values(1)
    if raskroy_id not in ids:
        return False
    row = ids.index(raskroy_id) + 1
    col = RASKROI_COLUMNS.index("pdf_name") + 1
    ws.update_cell(row, col, pdf_name)
    if sheet_count is not None:
        sc_col = RASKROI_COLUMNS.index("sheet_count") + 1
        ws.update_cell(row, sc_col, sheet_count)
    return True


# --------------------------------------------------------- raskroi_items ---

def get_raskroi_items():
    return _ws(ITEMS_SHEET).get_all_records()


def add_items(raskroy_id, items):
    """items: список (part_id, qty)."""
    if not items:
        return
    ws = _ws(ITEMS_SHEET)
    rows = [[raskroy_id, part_id, qty] for part_id, qty in items]
    ws.append_rows(rows)
