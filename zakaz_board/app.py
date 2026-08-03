import os
import time
from collections import defaultdict, OrderedDict
from datetime import datetime

from flask import Flask, render_template, request, redirect, url_for
import requests

import sheet_store

app = Flask(__name__)

CACHE_TTL = 15
_cache = {"data": None, "ts": 0}


def load_all(force=False):
    now = time.time()
    if force or _cache["data"] is None or now - _cache["ts"] > CACHE_TTL:
        parts = sheet_store.get_parts_pool()
        raskroi = sheet_store.get_raskroi()
        items = sheet_store.get_raskroi_items()
        _cache["data"] = (parts, raskroi, items)
        _cache["ts"] = now
    return _cache["data"]


def invalidate_cache():
    _cache["data"] = None


def compute_part_stats(parts, raskroi, items):
    """part_id -> (назначено_в_раскрои, вырезано)"""
    raskroi_by_id = {r["id"]: r for r in raskroi}
    assigned = defaultdict(int)
    cut = defaultdict(int)
    for it in items:
        pid = it.get("part_id")
        try:
            qty = int(it.get("qty") or 0)
        except (TypeError, ValueError):
            qty = 0
        assigned[pid] += qty
        rk = raskroi_by_id.get(it.get("raskroy_id"))
        if rk and rk.get("status") == "Вырезан":
            cut[pid] += qty
    return assigned, cut


def notify_bitrix24(text):
    webhook = os.environ.get("BITRIX_WEBHOOK_URL")
    user_ids = os.environ.get("BITRIX_USER_IDS", "")
    if not webhook or not user_ids:
        return
    webhook = webhook.rstrip("/") + "/"
    for uid in [u.strip() for u in user_ids.split(",") if u.strip()]:
        try:
            requests.post(f"{webhook}im.notify.system.add.json",
                          json={"USER_ID": uid, "MESSAGE": text}, timeout=5)
        except Exception as e:
            print("Битрикс24 уведомление не отправлено:", e)


def check_products_complete(affected_part_ids):
    """После того как раскрой вырезан — проверяет, не закрылось ли этим целиком изделие."""
    parts, raskroi, items = load_all(force=True)
    assigned, cut = compute_part_stats(parts, raskroi, items)
    parts_by_id = {p["id"]: p for p in parts}
    touched_products = set()
    for pid in affected_part_ids:
        p = parts_by_id.get(pid)
        if p:
            touched_products.add((p.get("order"), p.get("product")))

    for order, product in touched_products:
        product_parts = [p for p in parts if p.get("order") == order and p.get("product") == product]
        if not product_parts:
            continue
        total = sum(int(p.get("qty_total") or 0) for p in product_parts)
        done = sum(cut.get(p["id"], 0) for p in product_parts)
        if total and done >= total:
            notify_bitrix24(f"Изделие «{product}» (заказ {order}) полностью вырезано.")


def _safe_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def build_full_tree(parts):
    """Заказ -> Изделие -> Узел -> [детали], детали отсортированы стабильно."""
    tree = OrderedDict()
    for p in sorted(parts, key=lambda x: (str(x.get("order") or ""), str(x.get("product") or ""),
                                           str(x.get("uzel") or ""), str(x.get("grade") or ""),
                                           _safe_float(x.get("thickness")))):
        o = tree.setdefault(p.get("order"), OrderedDict())
        pr = o.setdefault(p.get("product"), OrderedDict())
        pr.setdefault(p.get("uzel"), []).append(p)
    return tree


# ------------------------------------------------------------------ ПДО ---

@app.route("/")
def pdo_view():
    parts, raskroi, items = load_all()
    assigned, cut = compute_part_stats(parts, raskroi, items)
    
    # 1. Загружаем приоритеты узлов из БД/Google Sheets
    # Ожидаемый формат priorities: {(order, product, uzel): {'target_date': '2026-08-10', 'comment': 'срочно'}}
    priorities = sheet_store.load_uzel_priorities()

    for p in parts:
        p["_assigned"] = assigned.get(p["id"], 0)
        p["_cut"] = cut.get(p["id"], 0)

    # 2. Строим базовое дерево
    tree = build_full_tree(parts)

    # 3. Обогащаем дерево приоритетами и сортируем узлы
    today_str = date.today().isoformat()
    
    for order_name, products in tree.items():
        for product_name, uzly in products.items():
            
            # Собираем новую структурированную информацию по каждому узлу
            uzly_with_meta = {}
            for uzel_name, part_list in uzly.items():
                prio_key = (order_name, product_name, uzel_name)
                prio_data = priorities.get(prio_key, {})
                
                target_date = prio_data.get("target_date", "")
                comment = prio_data.get("comment", "")
                
                # Проверяем, просрочена ли дата резки
                is_expired = False
                if target_date and target_date < today_str:
                    is_expired = True

                uzly_with_meta[uzel_name] = {
                    "parts": part_list,
                    "target_date": target_date,
                    "comment": comment,
                    "is_expired": is_expired
                }

            # 4. Функция-ключ для сортировки узлов
            def sort_uzel_key(item):
                u_name, u_data = item
                t_date = u_data["target_date"]
                if t_date:
                    try:
                        # (0 = Блок с датами, дата для сортировки, имя узла)
                        return (0, datetime.strptime(t_date, "%Y-%m-%d"), u_name)
                    except ValueError:
                        pass
                # (1 = Блок "Без даты", максимальная дата, имя узла)
                return (1, datetime.max, u_name)

            # Перезаписываем узлы отсортированным словарём
            products[product_name] = dict(sorted(uzly_with_meta.items(), key=sort_uzel_key))

    return render_template("pdo.html", tree=tree, active="pdo")

@app.route("/pdo/update_priority", methods=["POST"])
def update_priority():
    order = request.form.get("order")
    product = request.form.get("product")
    uzel = request.form.get("uzel")
    target_date = request.form.get("target_date") # строка YYYY-MM-DD
    comment = request.form.get("comment")

    # Функция в sheet_store, которая записывает/обновляет строчку в Google Таблице
    sheet_store.save_uzel_priority(order, product, uzel, target_date, comment)

    return redirect(url_for("pdo_view"))

# -------------------------------------------------------------- Технолог ---

def group_uzel_by_thickness(tree):
    """Заказ->Изделие->Узел->[детали] превращает в ...->Узел->{исходная_строка_материала:[детали]}"""
    out = OrderedDict()
    for order, products in tree.items():
        out[order] = OrderedDict()
        for product, uzly in products.items():
            out[order][product] = OrderedDict()
            for uzel, parts_list in uzly.items():
                groups = OrderedDict()
                for p in parts_list:
                    label = p.get("material_raw") or f'{p.get("grade") or "?"} {p.get("thickness") or "?"}мм'
                    groups.setdefault(label, []).append(p)
                out[order][product][uzel] = groups
    return out


@app.route("/tehnolog")
def tehnolog_view():
    parts, raskroi, items = load_all()
    assigned, _ = compute_part_stats(parts, raskroi, items)

    f_order = request.args.get("order", "")
    f_product = request.args.get("product", "")
    f_q = request.args.get("q", "").strip().lower()

    orders = sorted({p["order"] for p in parts if p.get("order")})
    products = sorted({p["product"] for p in parts if p.get("product") and (not f_order or p["order"] == f_order)})

    def matches(p):
        if f_order and p.get("order") != f_order:
            return False
        if f_product and p.get("product") != f_product:
            return False
        if f_q and f_q not in f"{p.get('code','')} {p.get('name','')}".lower():
            return False
        return True

    filtered = []
    for p in parts:
        if not matches(p):
            continue
        remaining = int(p.get("qty_total") or 0) - assigned.get(p["id"], 0)
        if remaining <= 0:
            continue  # уже полностью разложено по раскроям — нечего больше экспортировать
        p = dict(p)
        p["_remaining"] = remaining
        filtered.append(p)

    tree = group_uzel_by_thickness(build_full_tree(filtered))

    editable_raskroi = [r for r in raskroi if r.get("status") == "Создан"]
    active_raskroy_id = request.args.get("raskroy", "")
    active_raskroy = next((r for r in editable_raskroi if r["id"] == active_raskroy_id), None)

    active_items = []
    if active_raskroy:
        parts_by_id = {p["id"]: p for p in parts}
        for it in items:
            if it.get("raskroy_id") == active_raskroy_id:
                p = dict(parts_by_id.get(it["part_id"], {}))
                p["qty_in_raskroy"] = it.get("qty")
                active_items.append(p)

    return render_template("tehnolog.html", tree=tree, orders=orders, products=products,
                            f_order=f_order, f_product=f_product, f_q=f_q,
                            active_raskroy=active_raskroy, active_items=active_items,
                            msg=request.args.get("msg", ""), active="tehnolog")


@app.route("/tehnolog/export", methods=["POST"])
def export_view():
    name = request.form.get("name", "").strip()[:8]
    group_values = request.form.getlist("group_ids")  # каждое значение: "id1,id2,id3"
    if not name:
        return redirect(url_for("tehnolog_view", msg="Укажите название раскроя (УП)"))
    if not group_values:
        return redirect(url_for("tehnolog_view", msg="Отметьте хотя бы один лист"))

    part_ids = set()
    for s in group_values:
        part_ids.update(x for x in s.split(",") if x)

    parts, raskroi, items = load_all()
    assigned, _ = compute_part_stats(parts, raskroi, items)
    parts_by_id = {p["id"]: p for p in parts}
    selected = [parts_by_id[pid] for pid in part_ids if pid in parts_by_id]
    if not selected:
        return redirect(url_for("tehnolog_view", msg="Детали не найдены"))

    keys = {(p.get("grade"), p.get("thickness"), p.get("custom_sheet") or "") for p in selected}
    if len(keys) > 1:
        return redirect(url_for("tehnolog_view",
                                 msg="Выбраны листы с разной толщиной/маркой — за один экспорт можно только одну толщину"))

    grade, thickness, custom_sheet = next(iter(keys))
    raskroy_id = sheet_store.create_raskroy(name, grade, thickness, custom_sheet)

    to_add = []
    for p in selected:
        remaining = int(p.get("qty_total") or 0) - assigned.get(p["id"], 0)
        if remaining > 0:
            to_add.append((p["id"], remaining))
    sheet_store.add_items(raskroy_id, to_add)
    invalidate_cache()
    return redirect(url_for("tehnolog_view", raskroy=raskroy_id,
                             msg=f"Раскрой «{name}» создан, деталей: {len(to_add)}. "
                                 f"Список для поиска DXF можно скачать в карточке ниже."))


@app.route("/tehnolog/download/<raskroy_id>")
def download_raskroy(raskroy_id):
    parts, raskroi, items = load_all()
    parts_by_id = {p["id"]: p for p in parts}
    lines = ["Код;Наименование;Толщина;Размер;Количество"]
    for it in items:
        if it.get("raskroy_id") == raskroy_id:
            p = parts_by_id.get(it["part_id"], {})
            lines.append(f"{p.get('code','')};{p.get('name','')};{p.get('thickness','')};"
                         f"{p.get('size','')};{it.get('qty','')}")
    content = "\ufeff" + "\n".join(lines)
    from flask import Response
    return Response(content, mimetype="text/csv; charset=utf-8",
                     headers={"Content-Disposition": f"attachment; filename=raskroy_{raskroy_id}.csv"})


@app.route("/tehnolog/upload_pdf", methods=["POST"])
def upload_pdf():
    raskroy_id = request.form.get("raskroy_id")
    file = request.files.get("pdf")
    if file and raskroy_id:
        import tempfile
        from ajan_pdf_parser import parse_ajan_pdf
        sheet_count = "1"
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                file.save(tmp.name)
                tmp_path = tmp.name
            ups = parse_ajan_pdf(tmp_path)
            if ups:
                sheet_count = str(ups[0].get("repeat_count") or "1")
        except Exception as e:
            print("Не удалось разобрать PDF (сохранили только имя файла):", e)
        finally:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
        sheet_store.attach_pdf(raskroy_id, file.filename, sheet_count=sheet_count)
        invalidate_cache()
    return redirect(url_for("tehnolog_view", raskroy=raskroy_id, msg="PDF прикреплён"))


# -------------------------------------------------------------- Учётчик ---

@app.route("/uchetnik")
def uchetnik_view():
    parts, raskroi, items = load_all()
    pending = [r for r in raskroi if r.get("status") == "Создан"]

    groups = OrderedDict()
    for r in sorted(pending, key=lambda x: (str(x.get("grade") or ""), str(x.get("thickness") or ""))):
        key = (r.get("grade"), r.get("thickness"), r.get("custom_sheet") or "")
        groups.setdefault(key, []).append(r)

    item_count = defaultdict(int)
    for it in items:
        item_count[it.get("raskroy_id")] += 1

    def sheet_count_of(r):
        try:
            return int(r.get("sheet_count") or 1)
        except (TypeError, ValueError):
            return 1

    group_totals = {key: sum(sheet_count_of(r) for r in group) for key, group in groups.items()}

    return render_template("uchetnik.html", groups=groups, item_count=item_count,
                            group_totals=group_totals, sheet_count_of=sheet_count_of, active="uchetnik")


@app.route("/uchetnik/issue/<raskroy_id>", methods=["POST"])
def issue_raskroy(raskroy_id):
    sheet_store.update_raskroy_status(raskroy_id, "Выдан")
    invalidate_cache()
    return redirect(url_for("uchetnik_view"))


# ----------------------------------------------------------- Плазменщик ---

@app.route("/plazmenshik")
def plazmenshik_view():
    parts, raskroi, items = load_all()
    queue = [r for r in raskroi if r.get("status") == "Выдан"]
    parts_by_id = {p["id"]: p for p in parts}

    queue_with_items = []
    for r in queue:
        r_items = []
        uzly = set()
        for it in items:
            if it.get("raskroy_id") == r["id"]:
                p = dict(parts_by_id.get(it["part_id"], {}))
                p["qty_in_raskroy"] = it.get("qty")
                r_items.append(p)
                if p.get("uzel"):
                    uzly.add(f'{p.get("product")} / {p.get("uzel")}')
        queue_with_items.append((r, r_items, sorted(uzly)))

    return render_template("plazmenshik.html", queue=queue_with_items, active="plazmenshik")


@app.route("/plazmenshik/cut/<raskroy_id>", methods=["POST"])
def cut_raskroy(raskroy_id):
    parts, raskroi, items = load_all()
    affected_part_ids = [it["part_id"] for it in items if it.get("raskroy_id") == raskroy_id]

    sheet_store.update_raskroy_status(raskroy_id, "Вырезан")
    invalidate_cache()
    check_products_complete(affected_part_ids)
    return redirect(url_for("plazmenshik_view"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
