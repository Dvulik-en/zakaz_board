"""
sync_ajan_pdfs.py
==================
Берёт папку с PDF-отчётами AJANCAM (04S_SL01.pdf, 06S_SL02.pdf и т.д.),
разбирает каждый и проставляет в Google Таблице (той же, что читает сайт):
  - колонку "up" — по имени файла PDF (например, "04S_SL01")
  - колонку "cut_time" — суммарное время резки этой детали на этой УП
    (на будущее, для планирования загрузки плазмы; сайт её пока не показывает)

Сопоставление детали из PDF со строкой в таблице идёт по коду детали,
приведённому к общему виду (см. ajan_pdf_parser.normalize_code) — это нужно,
потому что AJAN иногда транслитерирует код ("БКМ" -> "BKM"), а в таблице
код остаётся кириллицей, как в исходных xlsm.

Если код детали из PDF не находится ни в одной строке таблицы — деталь
выводится отдельным списком в конце, ничего не падает и не портится.

Использование:
    python sync_ajan_pdfs.py "C:\\путь\\к\\папке\\с\\pdf"

Переменные окружения нужны те же, что для sync_to_sheet.py:
    GOOGLE_CREDENTIALS_JSON, GOOGLE_SHEET_NAME
"""

import sys
import os
import glob
from collections import defaultdict

import gspread

import sheet_store
from ajan_pdf_parser import parse_ajan_pdf, normalize_code


def build_lookup(existing_rows):
    """normalize_code(код) -> список номеров строк (2-based, с учётом заголовка)"""
    lookup = defaultdict(list)
    for i, r in enumerate(existing_rows, start=2):
        key = normalize_code(r.get("code", ""))
        if key:
            lookup[key].append(i)
    return lookup


def main():
    if len(sys.argv) < 2:
        print('Использование: python sync_ajan_pdfs.py "папка с PDF"')
        sys.exit(1)

    folder = sys.argv[1]
    pdf_files = sorted(glob.glob(os.path.join(folder, "**", "*.pdf"), recursive=True))
    if not pdf_files:
        print("В указанной папке PDF не найдено (ищу и во вложенных подпапках).")
        sys.exit(1)
    print(f"Найдено PDF: {len(pdf_files)}")

    ws = sheet_store._worksheet()
    existing = ws.get_all_records()
    lookup = build_lookup(existing)

    # добавим колонку cut_time в конец, если её ещё нет
    header = ws.row_values(1)
    if "cut_time" not in header:
        ws.update_cell(1, len(header) + 1, "cut_time")
        cut_time_col = len(header) + 1
    else:
        cut_time_col = header.index("cut_time") + 1
    up_col = header.index("up") + 1 if "up" in header else None
    if up_col is None:
        print("В таблице нет колонки 'up' — проверьте структуру листа 'parts'.")
        sys.exit(1)

    updates = []
    not_found = []
    matched_count = 0

    for pdf_path in pdf_files:
        try:
            ups = parse_ajan_pdf(pdf_path)
        except Exception as e:
            print(f"  ! Не удалось разобрать {os.path.basename(pdf_path)}: {e}")
            continue
        for up in ups:
            for p in up["parts"]:
                key = normalize_code(p["code"])
                rows = lookup.get(key)
                if not rows:
                    not_found.append((up["up_name"], p["code"], p["name"]))
                    continue
                for row_num in rows:
                    updates.append({"range": gspread.utils.rowcol_to_a1(row_num, up_col),
                                     "values": [[up["up_name"]]]})
                    updates.append({"range": gspread.utils.rowcol_to_a1(row_num, cut_time_col),
                                     "values": [[p["time"] or ""]]})
                    matched_count += 1

    if updates:
        # batch_update принимает не более ~几千 диапазонов за раз без проблем для наших объёмов
        ws.batch_update(updates)

    print(f"\nСопоставлено и обновлено строк: {matched_count}")
    if not_found:
        print(f"Не найдено соответствия в таблице: {len(not_found)} деталей")
        print("(это нормально, если PDF относится к изделию, которое ещё не заливали через sync_to_sheet.py)")
        for up_name, code, name in not_found[:20]:
            print(f"   УП {up_name}: {code} — {name}")
        if len(not_found) > 20:
            print(f"   ...и ещё {len(not_found) - 20}")


if __name__ == "__main__":
    main()
