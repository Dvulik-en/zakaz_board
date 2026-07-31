"""
upload_parts.py
================
Заливает детали в parts_pool БЕЗ привязки к плану ПДО — просто из папки заказа.
Дробит на уровне отдельного xlsm-файла (узла), не только на уровне папки-изделия,
так что дальше в "Технологе" можно фильтровать и по конкретному узлу тоже.

Структура ожидается та же, что раньше: Заказ / ПРОДУКЦИЯ / <изделие> / *.xlsm
(на любом уровне вложенности внутри папки изделия).

При повторном запуске на той же папке уже существующие детали (по совпадению
заказ+изделие+узел+код) не дублируются — только новые.

Использование:
    python upload_parts.py "путь к папке заказа"
"""

import sys
import os
import glob

import sheet_store
import zakaz_report as zr  # переиспользуем extract_plasma_parts и поиск папки ПРОДУКЦИЯ


def collect_rows(folder):
    order_name = os.path.basename(os.path.normpath(folder))
    product_root = zr.find_produkciya_folder(folder) or folder
    entries = sorted(os.listdir(product_root))
    product_dirs = [e for e in entries if os.path.isdir(os.path.join(product_root, e))]
    loose_files = [e for e in entries if e.lower().endswith(".xlsm") and not e.startswith("~$")]

    rows = []

    def process_file(filepath, product_name):
        uzel_name = os.path.splitext(os.path.basename(filepath))[0]
        for p in zr.extract_plasma_parts(filepath):
            custom_str = (f"{p['custom_sheet'][0]:g}x{p['custom_sheet'][1]:g}"
                          if p["custom_sheet"] else "")
            rows.append({
                "order": order_name, "product": product_name, "uzel": uzel_name,
                "grade": p["grade"], "thickness": p["thickness"], "custom_sheet": custom_str,
                "code": str(p["code"]), "name": p["name"],
                "size": zr.size_to_str(p["size_info"]), "qty_total": p["qty"],
                "area": round(p["area_m2"], 3),
            })

    for pdir in product_dirs:
        full_path = os.path.join(product_root, pdir)
        files = [f for f in glob.glob(os.path.join(full_path, "**", "*.xlsm"), recursive=True)
                 if not os.path.basename(f).startswith("~$")]
        for f in sorted(files):
            process_file(f, pdir)

    for lf in loose_files:
        process_file(os.path.join(product_root, lf), os.path.splitext(lf)[0])

    return rows


def main():
    if len(sys.argv) < 2:
        print('Использование: python upload_parts.py "папка заказа"')
        sys.exit(1)

    folder = sys.argv[1]
    if not os.path.isdir(folder):
        print("Такой папки нет.")
        sys.exit(1)

    print("Читаю эксельки...")
    rows = collect_rows(folder)
    if not rows:
        print("Деталей на плазму не найдено — проверьте путь и структуру папок.")
        return
    print(f"Найдено деталей: {len(rows)}")

    existing = sheet_store.get_parts_pool()

    to_add = []
    skipped = 0
    for r in rows:
        if sheet_store.find_part_by_natural_key(existing, r["order"], r["product"], r["uzel"], r["code"]):
            skipped += 1
            continue
        to_add.append(r)

    if to_add:
        sheet_store.add_parts(to_add)
    print(f"Добавлено новых: {len(to_add)}")
    print(f"Уже были в таблице (пропущено): {skipped}")


if __name__ == "__main__":
    main()
