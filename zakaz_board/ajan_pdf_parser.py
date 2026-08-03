"""
ajan_pdf_parser.py
===================
Разбор PDF-отчётов, которые генерирует AJANCAM после раскладки (04S_SL01.pdf и т.п.).

Каждый такой PDF может содержать НЕСКОЛЬКО листов-УП (одна страница = одна УП) плюс
итоговую страницу "Общий список деталей" в конце — она пропускается автоматически
(на ней нет поля "Имя файла").

Важный нюанс реальных данных: в названии детали код и наименование иногда
транслитерируются AJAN по-разному — например, "БКМ" превращается в "BKM",
а "ФР" остаётся кириллицей. Чтобы сопоставление с кодами из xlsm-спецификаций
(которые все набраны кириллицей) работало, используется normalize_code() —
приводит и то, и другое к одному регистру латиницы перед сравнением.

Использование как модуля:
    from ajan_pdf_parser import parse_ajan_pdf, normalize_code
    ups = parse_ajan_pdf("04S_SL01.pdf")
    # ups -> [{"up_name": "04S_SL01", "material": "09Г2С", "thickness": "4",
    #          "sheet_size": "6000X1500mm",
    #          "parts": [{"code":..., "name":..., "qty_in_name":..., "size":...,
    #                     "weight":..., "time":..., "qty":..., "perimeter":...}, ...]}, ...]

Зависимости: pdfplumber (pip install pdfplumber)
"""

import re

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

TRANSLIT = {
    'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'G', 'Д': 'D', 'Е': 'E', 'Ё': 'E', 'Ж': 'ZH',
    'З': 'Z', 'И': 'I', 'Й': 'I', 'К': 'K', 'Л': 'L', 'М': 'M', 'Н': 'N', 'О': 'O',
    'П': 'P', 'Р': 'R', 'С': 'S', 'Т': 'T', 'У': 'U', 'Ф': 'F', 'Х': 'H', 'Ц': 'TS',
    'Ч': 'CH', 'Ш': 'SH', 'Щ': 'SCH', 'Ъ': '', 'Ы': 'Y', 'Ь': '', 'Э': 'E', 'Ю': 'YU', 'Я': 'YA',
}


def normalize_code(code):
    """Приводит код детали к сравнимому виду независимо от того, набран он
    кириллицей, латиницей или их смесью (как бывает в PDF от AJAN)."""
    if not code:
        return ""
    s = str(code).upper()
    s = "".join(TRANSLIT.get(ch, ch) for ch in s)
    s = re.sub(r"[\s\-]+", "", s)
    return s


def parse_ajan_pdf(path):
    """Возвращает список УП (программ раскроя) с деталями, найденными в PDF."""
    if pdfplumber is None:
        raise ImportError("Нужен pdfplumber: pip install pdfplumber")

    results = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            meta = {}
            parts_table = None

            for t in tables:
                header_zone = " ".join(str(c) for row in t for c in row if c)[:400]
                if "Имяфайла" in header_zone:
                    for row in t:
                        for i, cell in enumerate(row):
                            if cell == "Имяфайла" and i + 1 < len(row) and row[i + 1]:
                                meta["up_name"] = row[i + 1]
                            elif cell == "Толщина(mm)":
                                for j in range(i + 1, len(row)):
                                    if row[j] not in (None, ""):
                                        meta["thickness"] = row[j]
                                        break
                            elif cell == "Материал" and i + 1 < len(row) and row[i + 1]:
                                meta["material"] = row[i + 1]
                            elif cell == "Размерлиста" and i + 1 < len(row) and row[i + 1]:
                                meta["sheet_size"] = row[i + 1]
                            elif cell == "Кол-воповторов" and i + 1 < len(row) and row[i + 1]:
                                meta["repeat_count"] = row[i + 1]
                if t and t[0] and t[0][0] and "Номер" in str(t[0][0]):
                    parts_table = t

            if not meta.get("up_name") or not parts_table:
                continue  # сводная страница или служебный лист — пропускаем

            parts = []
            for row in parts_table[1:]:
                name_cell = next(
                    (c for c in row if c and re.match(r"^\d*_", c.replace("\n", ""))), None
                )
                if not name_cell:
                    continue
                lines = name_cell.split("\n")
                m0 = re.match(r"^\d+(?:[.,]\d+)?_(.+)$", lines[0].strip())
                code = (m0.group(1) if m0 else lines[0].strip().lstrip("_")).rstrip("-").strip()
                rest = " ".join(lines[1:]).strip().lstrip("-").strip()
                m2 = re.match(r"^(.+?)\((\d+)шт\)$", rest)
                name, qty_in_name = (m2.group(1), m2.group(2)) if m2 else (rest, None)

                parts.append({
                    "code": code,
                    "name": name,
                    "qty_in_name": qty_in_name,
                    "size": (row[3] or "").replace("\n", " ") if len(row) > 3 else None,
                    "weight": row[4] if len(row) > 4 else None,
                    "time": row[5] if len(row) > 5 else None,
                    "qty": row[6] if len(row) > 6 else None,
                    "perimeter": row[7] if len(row) > 7 else None,
                })

            results.append({
                "up_name": meta["up_name"].rsplit(".", 1)[0],
                "material": meta.get("material"),
                "thickness": meta.get("thickness"),
                "sheet_size": meta.get("sheet_size"),
                "repeat_count": meta.get("repeat_count", "1"),
                "parts": parts,
            })
    return results
