"""
sync_to_sheet.py
=================
То же самое, что zakaz_report.py (план ПДО -> папки заказов -> детали на плазму),
только результат не сохраняется в локальный xlsx, а заливается в Google Таблицу,
которую читает веб-доска (app.py + sheet_store.py).

Логика сохранения уже введённых Статус/УП работает так же, как в zakaz_report.py:
если деталь (по коду + изделию) уже есть в таблице — обновляются все поля,
КРОМЕ status и up (их меняют через сайт). Если деталь новая — добавляется
новая строка с пустыми status/up.

Запускать с того же ПК, где обычно запускали zakaz_report.py — на нём же
должен быть настроен GOOGLE_CREDENTIALS_JSON и GOOGLE_SHEET_NAME (см. SETUP.md).

Использование:
    python sync_to_sheet.py "Заготовительный_участок.xlsx"
"""

import sys
import os
import uuid

import openpyxl
import gspread

import sheet_store
# переиспользуем весь парсинг из zakaz_report.py — он должен лежать рядом
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import zakaz_report as zr


def sync():
    if len(sys.argv) < 2:
        print('Использование: python sync_to_sheet.py "Заготовительный_участок.xlsx"')
        sys.exit(1)

    plan_path = sys.argv[1]
    wb = openpyxl.load_workbook(plan_path, data_only=True)
    entries = zr.parse_plan(wb[zr.PLAN_SHEET_NAME])
    groups = zr.group_plan_entries(entries)

    labels = sorted(groups.keys())
    print(f"В плане {len(labels)} заказов/изделий.\n")

    new_rows = []  # накопим все детали для заливки
    for label in labels:
        path = input(f"[{label}]\nпапка (Enter — пропустить): ").strip().strip('"')
        if not path or not os.path.isdir(path):
            continue
        order_name = os.path.basename(os.path.normpath(path))
        target_produkciya = groups[label]["entries"][0]["produkciya"] if groups[label]["entries"] else None
        structure = zr.build_product_groups(path, target_produkciya=target_produkciya)
        dates = ", ".join(sorted(groups[label]["dates"]))
        for product_name, product_groups in structure.items():
            for (grade, thickness, custom_sheet), g in product_groups.items():
                custom_str = f"{custom_sheet[0]:g}x{custom_sheet[1]:g}" if custom_sheet else ""
                for p in g["parts"]:
                    new_rows.append({
                        "order": order_name, "product": product_name, "date": dates,
                        "grade": grade, "thickness": thickness, "custom_sheet": custom_str,
                        "code": str(p["code"]), "name": p["name"],
                        "size": zr.size_to_str(p["size_info"]), "qty": p["qty"],
                        "area": round(p["area_m2"], 3),
                    })
        print(f"  ok: собрано строк — {len(new_rows)} (нарастающим итогом)")

    if not new_rows:
        print("Нечего заливать, выхожу.")
        return

    push_to_sheet(new_rows)


def push_to_sheet(new_rows):
    ws = sheet_store._worksheet()
    existing = ws.get_all_records()
    existing_index = {(str(r["product"]), str(r["code"])): r for r in existing}

    updates = []  # (row_number, values) для уже существующих
    appends = []  # новые строки целиком

    for i, r in enumerate(existing, start=2):  # 2, т.к. строка 1 — заголовок
        key = (str(r["product"]), str(r["code"]))
        match = next((nr for nr in new_rows if (nr["product"], nr["code"]) == key), None)
        if match:
            values = [r["id"], match["order"], match["product"], match["date"],
                      match["grade"], match["thickness"], match["custom_sheet"],
                      match["code"], match["name"], match["size"], match["qty"],
                      match["area"], r["status"], r["up"]]  # status/up сохраняем как были
            updates.append((i, values))

    matched_keys = {(str(r["product"]), str(r["code"])) for r in existing}
    for nr in new_rows:
        key = (nr["product"], nr["code"])
        if key not in matched_keys:
            new_id = str(uuid.uuid4())[:8]
            appends.append([new_id, nr["order"], nr["product"], nr["date"], nr["grade"],
                             nr["thickness"], nr["custom_sheet"], nr["code"], nr["name"],
                             nr["size"], nr["qty"], nr["area"], "", ""])

    if updates:
        batch = [{"range": f"A{row}:N{row}", "values": [values]} for row, values in updates]
        ws.batch_update(batch)
    if appends:
        ws.append_rows(appends)

    print(f"\nОбновлено существующих строк: {len(updates)}")
    print(f"Добавлено новых строк: {len(appends)}")


if __name__ == "__main__":
    sync()
