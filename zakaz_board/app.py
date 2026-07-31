import os
import time
from collections import defaultdict, OrderedDict

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


# ------------------------------------------------------------------ ПДО ---

@app.route("/")
def pdo_view():
    parts, raskroi, items = load_all()
    assigned, cut = compute_part_stats(parts, raskroi, items)

    tree = OrderedDict()
    for p in sorted(parts, key=lambda x: (str(x.get("order") or ""), str(x.get("product") or ""))):
        o = tree.setdefault(p["order"], OrderedDict())
        pr = o.setdefault(p["product"], {"total": 0, "assigned": 0, "cut": 0})
        qty = int(p.get("qty_total") or 0)
        pr["total"] += qty
        pr["assigned"] += min(assigned.get(p["id"], 0), qty)
        pr["cut"] += cut.get(p["id"], 0)

    return render_template("pdo.html", tree=tree, active="pdo")


# -------------------------------------------------------------- Технолог ---

@app.route("/tehnolog")
def tehnolog_view():
    parts, raskroi, items = load_all()
    assigned, _ = compute_part_stats(parts, raskroi, items)

    f_order = request.args.get("order", "")
    f_product = request.args.get("product", "")
    f_uzel = request.args.get("uzel", "")
    f_q = request.args.get("q", "").strip().lower()

    orders = sorted({p["order"] for p in parts if p.get("order")})
    products = sorted({p["product"] for p in parts if p.get("product") and (not f_order or p["order"] == f_order)})
    uzly = sorted({p["uzel"] for p in parts if p.get("uzel")
                   and (not f_order or p["order"] == f_order)
                   and (not f_product or p["product"] == f_product)})

    def matches(p):
        if f_order and p.get("order") != f_order:
            return False
        if f_product and p.get("product") != f_product:
            return False
        if f_uzel and p.get("uzel") != f_uzel:
            return False
        if f_q and f_q not in f"{p.get('code','')} {p.get('name','')}".lower():
            return False
        return True

    filtered = [p for p in parts if matches(p)]
    for p in filtered:
        p["_remaining"] = int(p.get("qty_total") or 0) - assigned.get(p["id"], 0)

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

    return render_template("tehnolog.html", parts=filtered, orders=orders, products=products,
                            uzly=uzly, f_order=f_order, f_product=f_product, f_uzel=f_uzel, f_q=f_q,
                            editable_raskroi=editable_raskroi, active_raskroy=active_raskroy,
                            active_items=active_items, msg=request.args.get("msg", ""), active="tehnolog")


@app.route("/tehnolog/create_raskroy", methods=["POST"])
def create_raskroy():
    name = request.form.get("name", "").strip()[:8]
    grade = request.form.get("grade", "").strip()
    thickness = request.form.get("thickness", "").strip()
    custom_sheet = request.form.get("custom_sheet", "").strip()
    if not name or not grade or not thickness:
        return redirect(url_for("tehnolog_view", msg="Заполните название, марку и толщину"))
    rid = sheet_store.create_raskroy(name, grade, thickness, custom_sheet)
    invalidate_cache()
    return redirect(url_for("tehnolog_view", raskroy=rid, msg=f"Раскрой «{name}» создан"))


@app.route("/tehnolog/add_items", methods=["POST"])
def add_items():
    raskroy_id = request.form.get("raskroy_id")
    part_ids = request.form.getlist("part_id")
    parts, raskroi, items = load_all()
    raskroy = next((r for r in raskroi if r["id"] == raskroy_id), None)
    if not raskroy:
        return redirect(url_for("tehnolog_view", msg="Раскрой не найден"))

    parts_by_id = {p["id"]: p for p in parts}
    to_add = []
    skipped = 0
    for pid in part_ids:
        p = parts_by_id.get(pid)
        try:
            qty = int(request.form.get(f"qty_{pid}") or 0)
        except ValueError:
            qty = 0
        if not p or qty <= 0:
            continue
        if str(p.get("grade")) != str(raskroy.get("grade")) or str(p.get("thickness")) != str(raskroy.get("thickness")):
            skipped += 1
            continue
        to_add.append((pid, qty))

    sheet_store.add_items(raskroy_id, to_add)
    invalidate_cache()
    msg = f"Добавлено деталей: {len(to_add)}"
    if skipped:
        msg += f", пропущено (не совпала марка/толщина): {skipped}"
    return redirect(url_for("tehnolog_view", raskroy=raskroy_id,
                             order=request.form.get("f_order", ""),
                             product=request.form.get("f_product", ""),
                             uzel=request.form.get("f_uzel", ""),
                             q=request.form.get("f_q", ""), msg=msg))


@app.route("/tehnolog/upload_pdf", methods=["POST"])
def upload_pdf():
    raskroy_id = request.form.get("raskroy_id")
    file = request.files.get("pdf")
    if file and raskroy_id:
        sheet_store.attach_pdf(raskroy_id, file.filename)
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

    return render_template("uchetnik.html", groups=groups, item_count=item_count, active="uchetnik")


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
        for it in items:
            if it.get("raskroy_id") == r["id"]:
                p = dict(parts_by_id.get(it["part_id"], {}))
                p["qty_in_raskroy"] = it.get("qty")
                r_items.append(p)
        queue_with_items.append((r, r_items))

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
