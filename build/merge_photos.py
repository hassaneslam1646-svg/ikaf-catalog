"""
دمج صور الأصناف (غير الـPDF) وكميات تقرير المخزون مع الكتالوج.

بيتشغّل بعد build.py (أو لوحده على index.html الموجود):

    python3 build/merge_photos.py <فولدر الصور> "<تقرير المخزون>.xlsx" "<شجرة الأكواد>.xlsx"

فولدر الصور: صور jpg باسم الكود 14 رقم، والصور الإضافية لنفس الكود: <الكود>_2.jpg ...
(ده ناتج prepare_site_images.ps1 على جهاز المستخدم).

اللي بيعمله:
  1) كمية كل صنف = عمود "المستودعات الرئيسية" من شيت "المخزون" في تقرير المخزون
     (أصناف الـPDF كمان — تقرير المخزون أحدث من الرقم المطبوع).
  2) أصناف الـPDF اللي ليها صور في الفولدر: الصور بتتضاف كصور إضافية للصنف.
  3) أكواد الفولدر اللي مش في الـPDF: بتتضاف كأصناف جديدة لو رصيدها أكبر من 20،
     والريشيو والموديل بيتحسبوا من الكود بشجرة الأكواد.
  4) صور الموقع بتتكتب في photos/full و photos/thumb (webp).

التشغيل أكتر من مرة آمن: الأصناف والصور المضافة قبل كده بتتشال وتتبني من جديد.
"""
import json
import math
import os
import re
import shutil
import sys

import numpy as np
import openpyxl
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build import ROOT, load_tree, ratio_label, width_label, log  # noqa: E402

MIN_QTY = 20                 # الصنف الجديد بيظهر لو رصيده 20 أو أكتر
FULL_W, FULL_Q = 700, 72
THUMB_W, THUMB_Q = 330, 72
STOCK_SHEET = "المخزون"
STOCK_COL = "المستودعات الرئيسية"
# المخازن اللي مجموعها = المستودعات الرئيسية (عمود في التقرير -> الاسم اللي بيظهر للمندوب)
WAREHOUSES = [("الانتاج التام(الرياض", "الرياض"), ("م جدة الرئيسى", "جدة"),
              ("مستودع مكة المكرمة", "مكة")]
WH = {}   # الكود -> {المخزن: الكمية}

NAME_RE = re.compile(r"^(1\d{13})(?:_(\d+))?\.(jpe?g|png|webp)$", re.I)


def load_materials(xlsx):
    """الخامات من شيت ARABIC في الشجرة: العمود C اسم الخامة، D كودها (خانة 2-3 من الكود)."""
    ws = openpyxl.load_workbook(xlsx, data_only=True)["ARABIC"]
    out = {}
    for r in range(4, ws.max_row + 1):
        name, code = ws[f"C{r}"].value, ws[f"D{r}"].value
        if name is None or code is None or not str(name).strip():
            continue
        code = str(code).strip()
        if re.fullmatch(r"\d+", code):
            out[code.zfill(2)] = " ".join(str(name).split())
    return out


def load_stock(xlsx):
    wb = openpyxl.load_workbook(xlsx, data_only=True, read_only=True)
    if STOCK_SHEET not in wb.sheetnames:
        sys.exit(f"شيت '{STOCK_SHEET}' مش موجود في تقرير المخزون")
    rows = wb[STOCK_SHEET].iter_rows(values_only=True)
    head = [str(h).strip() if h is not None else "" for h in next(rows)]
    if STOCK_COL not in head:
        sys.exit(f"عمود '{STOCK_COL}' مش موجود في شيت المخزون")
    qi = head.index(STOCK_COL)
    wcols = [(head.index(c), lbl) for c, lbl in WAREHOUSES if c in head]
    out = {}
    for r in rows:
        if r[0] is None:
            continue
        code = str(r[0]).strip().split(".")[0].lstrip("0")
        if not code.isdigit():
            continue                                   # Grand Total وغيره
        q = r[qi] if isinstance(r[qi], (int, float)) else 0
        out[code] = out.get(code, 0) + q
        w = WH.setdefault(code, {})
        for i, lbl in wcols:
            if isinstance(r[i], (int, float)):
                w[lbl] = w.get(lbl, 0) + r[i]
    return out


def wh_split(code, total):
    """تفصيلة المخازن للصنف (من غير كسور). لو الكمية متعدّلة يدويًا التفصيلة بتتظبط بنفس النسبة."""
    w = {k: v for k, v in WH.get(code, {}).items() if v}
    raw = sum(w.values())
    if not w or raw <= 0:
        return None
    f = total / raw if abs(raw - total) > 0.01 else 1.0
    exact = {k: v * f for k, v in w.items()}
    out = {k: fmt_qty(v) for k, v in exact.items()}
    # التفصيلة لازم مجموعها = الكمية اللي ظاهرة بره؛ الباقي بيروح للمخزن اللي كسره أكبر
    for k in sorted(exact, key=lambda k: exact[k] - out[k], reverse=True)[:fmt_qty(total) - sum(out.values())]:
        out[k] += 1
    return {k: v for k, v in out.items() if v > 0} or None


def today():
    """تاريخ التحديث بتوقيت الرياض، بنفس شكل المستخدم: 1-10-2026."""
    from datetime import datetime
    from zoneinfo import ZoneInfo
    d = datetime.now(ZoneInfo("Asia/Riyadh"))
    return f"{d.day}-{d.month}-{d.year}"


def fmt_qty(q):
    # من غير كسور: الكمية بتتقرب للرقم الأقل (42.9 -> 42)
    return int(math.floor(float(q) + 1e-9))


def load_excluded():
    """أكواد المستخدم طلب تتشال من الكتالوج: build/excluded.json  {"الكود": "السبب"}."""
    p = os.path.join(ROOT, "build", "excluded.json")
    if not os.path.exists(p):
        return {}
    data = json.load(open(p, encoding="utf-8"))
    return {str(k).strip().lstrip("0"): v for k, v in data.items() if not str(k).startswith("_")}


def load_overrides():
    """كميات متعدّلة يدويًا من المستخدم: build/overrides.json  {"الكود": الكمية}.
    بتغلب على تقرير المخزون في كل تحديث لحد ما تتشال من الملف."""
    p = os.path.join(ROOT, "build", "overrides.json")
    if not os.path.exists(p):
        return {}
    data = json.load(open(p, encoding="utf-8"))
    return {str(k).strip().lstrip("0"): v for k, v in data.items() if not str(k).startswith("_")}


def read_items():
    html = open(os.path.join(ROOT, "index.html"), encoding="utf-8").read()
    m = re.search(r'<script id="data" type="application/json">(.*?)</script>', html, re.S)
    if not m:
        sys.exit("index.html مفيهوش بيانات الأصناف")
    src = re.search(r'var SRC\s*=\s*"(.*?)";', html)
    return json.loads(m.group(1)), (src.group(1) if src else "")


# ---------------------------------------------------------------------------
# تكرار الصور: الصورة الزيادة بتتشال لو هي نفس الصورة، أو نفس المجموعة اللونية
# (نفس مربعات COLORS تحت الكود). المجموعة اللونية المختلفة بتفضل.
# ---------------------------------------------------------------------------
CARD = (0.12267, 0.11500, 0.90044, 0.77900)   # صورة المنتج في قالب الكتالوج
SWATCH = (0.40, 0.862, 0.80, 0.912)           # مربعات COLORS
SAME_IMAGE = 3.0      # فرق متوسط البكسل في صورة المنتج
SAME_SWATCH = 7.5     # فرق متوسط البكسل في مربعات الألوان


def _region(im, box, size):
    w, h = im.size
    crop = im.crop((int(box[0] * w), int(box[1] * h), int(box[2] * w), int(box[3] * h)))
    return np.asarray(crop.resize(size, Image.BILINEAR), dtype=float)


def _sig(path):
    im = Image.open(path).convert("RGB")
    px = np.asarray(im.resize((50, 90)), dtype=float)[45, 1]
    poster = px[2] > px[0] + 40 and px[2] > 100        # خلفية الكتالوج الكحلي
    return {"card": _region(im, CARD, (32, 48)), "sw": _region(im, SWATCH, (40, 4)),
            "poster": poster}


def _same_group(a, b):
    if np.abs(a["card"] - b["card"]).mean() < SAME_IMAGE:
        return True
    return a["poster"] and b["poster"] and np.abs(a["sw"] - b["sw"]).mean() < SAME_SWATCH


def save_image(src, name):
    im = Image.open(src).convert("RGB")
    full = im.resize((FULL_W, round(im.height * FULL_W / im.width)), Image.LANCZOS) \
        if im.width > FULL_W else im
    full.save(os.path.join(ROOT, "photos", "full", name + ".webp"), "WEBP",
              quality=FULL_Q, method=6)
    thumb = im.resize((THUMB_W, round(im.height * THUMB_W / im.width)), Image.LANCZOS)
    thumb.save(os.path.join(ROOT, "photos", "thumb", name + ".webp"), "WEBP",
               quality=THUMB_Q, method=6)


def main():
    if len(sys.argv) != 4:
        sys.exit(__doc__)
    photos_dir, stock_xlsx, tree_xlsx = sys.argv[1:]
    # لو المصدر هو photos/ نفسه (مفيش فولدر الصور الأصلي)، ننسخه الأول لأن photos/ بيتمسح ويتبني من جديد
    if os.path.abspath(photos_dir).startswith(os.path.join(ROOT, "photos")):
        import tempfile
        tmp = tempfile.mkdtemp(prefix="ikaf-photos-")
        for fn in os.listdir(photos_dir):
            shutil.copy2(os.path.join(photos_dir, fn), tmp)
        photos_dir = tmp

    models, ratios = load_tree(tree_xlsx)
    materials = load_materials(tree_xlsx)
    stock = load_stock(stock_xlsx)
    overrides = load_overrides()
    for c, q in overrides.items():
        log(f"كمية متعدّلة يدويًا: {c} = {q} (التقرير: {stock.get(c)})")
        stock[c] = q
    log(f"تقرير المخزون: {len(stock)} كود")

    # الصور مجمّعة بالكود، بالترتيب: الكود.jpg ثم _2 ثم _3 ...
    photos = {}
    skipped = []
    for fn in sorted(os.listdir(photos_dir)):
        m = NAME_RE.match(fn)
        if not m:
            if not fn.lower().endswith(".csv"):
                skipped.append(fn)
            continue
        photos.setdefault(m.group(1), []).append((int(m.group(2) or 1), fn))
    for c in photos:
        photos[c] = [fn for _, fn in sorted(photos[c])]
    log(f"الصور: {sum(len(v) for v in photos.values())} صورة لـ {len(photos)} كود")

    items, source = read_items()
    excluded = load_excluded()
    if excluded:
        log(f"أكواد متشالة بطلب المستخدم: {sorted(excluded)}")
    items = [it for it in items if it["c"] not in excluded]
    items = [it for it in items if it.get("s") != "photo"]   # من تشغيل سابق
    for it in items:
        it.pop("g", None)

    shutil.rmtree(os.path.join(ROOT, "photos"), ignore_errors=True)
    os.makedirs(os.path.join(ROOT, "photos", "full"))
    os.makedirs(os.path.join(ROOT, "photos", "thumb"))

    dropped = []
    # صور المستخدم أكد إنها مكررة (اسم الملف من غير امتداد): build/drop_photos.json
    dp = os.path.join(ROOT, "build", "drop_photos.json")
    drop_manual = set(json.load(open(dp, encoding="utf-8")).get("photos", [])) if os.path.exists(dp) else set()

    pages_of = {}
    for it in items:
        if it.get("p"):
            pages_of.setdefault(it["c"], []).append(it["p"])
    done = {}

    def gallery(code, page=None):
        """صور الصنف من غير تكرار: صورة واحدة لكل مجموعة لونية.
        صفحات الـPDF للكود (لو موجودة) بتتحسب الأول وبتتفضّل."""
        if code in done:
            return list(done[code])
        kept = [_sig(os.path.join(ROOT, "full", f"{pg:03d}.webp")) for pg in pages_of.get(code, [])]
        names = []
        for fn in photos.get(code, []):
            src = os.path.join(photos_dir, fn)
            if os.path.splitext(fn)[0] in drop_manual:
                dropped.append(fn)
                continue
            sig = _sig(src)
            if any(_same_group(sig, k) for k in kept):
                dropped.append(fn)
                continue
            kept.append(sig)
            name = os.path.splitext(fn)[0]
            save_image(src, name)
            names.append(name)
        done[code] = names
        return list(names)

    pdf_codes = {it["c"] for it in items}
    qty_changed, pdf_no_stock = [], []
    for it in items:
        old = it.get("q")
        if it["c"] in stock:
            it["q"] = fmt_qty(stock[it["c"]])
            w = wh_split(it["c"], stock[it["c"]])
            if w:
                it["w"] = w
            else:
                it.pop("w", None)
        else:
            it["q"] = 0
            pdf_no_stock.append(it["c"])
        if old != it["q"]:
            qty_changed.append((it["c"], old, it["q"]))
        g = gallery(it["c"], it.get("p"))
        if g:
            it["g"] = g

    added, low, missing, unknown = [], [], [], []
    for code in sorted(set(photos) - pdf_codes - set(excluded)):
        q = stock.get(code)
        if q is None:
            missing.append(code)
            continue
        if q < MIN_QTY:
            low.append((code, q))
            continue
        rl, rn = ratio_label(code[8:10], ratios)
        wl, mn = width_label(code[4:6], models)
        if rl is None or mn is None:
            unknown.append((code, code[8:10], code[4:6]))
            continue
        items.append({"p": None, "c": code, "r": rl + (wl or ""), "rn": rn,
                      "q": fmt_qty(q), "w": wh_split(code, q), "m": mn, "mc": code[4:6], "rc": code[8:10],
                      "u": code[:11], "s": "photo", "g": gallery(code)})
        added.append(code)

    # الخامة لكل صنف (خانة 2-3) — للفلتر في الصفحة
    nomat = sorted({it["c"][1:3] for it in items if it["c"][1:3] not in materials})
    for it in items:
        it["k"] = materials.get(it["c"][1:3], "")
    if nomat:
        log(f"!! أكواد خامة مش في الشجرة: {nomat}")

    # الصفحة بتتبني من القالب عشان أي تعديل فيه يوصل
    tpl = open(os.path.join(ROOT, "build", "template.html"), encoding="utf-8").read()
    tpl = tpl.replace("__DATA__", json.dumps(items, ensure_ascii=False)).replace("__SOURCE__", source)
    tpl = tpl.replace("__UPDATED__", today())
    open(os.path.join(ROOT, "index.html"), "w", encoding="utf-8").write(tpl)

    pdf_items = [it for it in items if it.get("s") != "photo"]
    log(f"أصناف الـPDF: {len(pdf_items)} ({len(pdf_codes)} كود) — {sum(1 for it in pdf_items if it.get('g'))} منهم اتضافلهم صور")
    log(f"كميات اتحدّثت من تقرير المخزون: {len(qty_changed)}")
    if pdf_no_stock:
        log(f"!! أصناف PDF مش في تقرير المخزون (الكمية بقت 0): {pdf_no_stock}")
    log(f"أصناف جديدة من الصور: {len(added)}")
    if low:
        log(f"اتسابت (رصيد أقل من {MIN_QTY}): {low}")
    if missing:
        log(f"اتسابت (مش في تقرير المخزون): {missing}")
    if unknown:
        log(f"!! مش في الشجرة: {unknown}")
    if skipped:
        log(f"ملفات اسمها مش كود: {skipped}")
    if dropped:
        log(f"صور مكررة اتشالت (نفس الصورة أو نفس المجموعة اللونية): {len(dropped)}")
    log(f"الإجمالي: {len(items)} صنف")


if __name__ == "__main__":
    main()
