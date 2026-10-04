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

NEW_STOCK = False            # --new-stock: تقرير مخزون جديد -> تاريخ التحديث = النهارده
MIN_QTY = 20                 # الصنف الجديد بيظهر لو رصيده 20 أو أكتر
FULL_W, FULL_Q = 700, 72
THUMB_W, THUMB_Q = 330, 72
STOCK_SHEET = "المخزون"
STOCK_COL = "المستودعات الرئيسية"  # بعد توحيد المسافات
STOCK_COLS = [STOCK_COL, "كميات المخازن الرئيسية"]   # أسماء عمود الإجمالي في التقارير المختلفة
# المخازن اللي مجموعها = المستودعات الرئيسية (عمود في التقرير -> الاسم اللي بيظهر للمندوب)
WAREHOUSES = [("الانتاج التام(الرياض", "الرياض"), ("م جدة الرئيسى", "جدة"),
              ("مستودع مكة المكرمة", "مكة")]
WH = {}   # الكود -> {المخزن: الكمية}
CAT = {}  # الكود -> التصنيف (من عمود "التصنيف" في التقرير لو موجود)
PENDING_CATS = {"جلابيه", "جلابية"}   # الأصناف المتعرضة "جاري التصوير" لو مالهاش صورة

# اسم الصورة = الكود 14 رقم، ومسموح بأصفار قبله وأي زيادة بعده: 0104...jpg / 1042... (2).jpg / 1042..._2.jpg
NAME_RE = re.compile(r"^\s*0*(1\d{13})(.*?)\.(jpe?g|png|webp)$", re.I)


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


def load_colors(xlsx):
    """الألوان من شيت ARABIC: العمود O اسم اللون، P كوده (خانة 11 من الكود)."""
    ws = openpyxl.load_workbook(xlsx, data_only=True)["ARABIC"]
    out = {}
    for r in range(4, ws.max_row + 1):
        name, code = ws[f"O{r}"].value, ws[f"P{r}"].value
        if name is None or code is None or not str(name).strip():
            continue
        code = str(code).strip()
        if re.fullmatch(r"\d", code):
            out[code] = " ".join(str(name).split())
    return out


def _norm(h):
    return " ".join(str(h).split()) if h is not None else ""


def load_stock(xlsx):
    """تقرير المخزون: شيت "المخزون" (أو أول شيت فيه عمود كود الصنف).
    الكمية = عمود "المستودعات الرئيسية"، ولو مش موجود = مجموع المخازن الرئيسية التلاتة.
    صف العناوين ممكن مايبقاش أول صف (تصدير النظام أحيانًا بيبدأ بعنوان التقرير)."""
    wb = openpyxl.load_workbook(xlsx, data_only=True, read_only=True)
    names = ([STOCK_SHEET] if STOCK_SHEET in wb.sheetnames else []) + \
            [n for n in wb.sheetnames if n != STOCK_SHEET]
    wnames = [_norm(c) for c, _ in WAREHOUSES]
    for name in names:
        rows = list(wb[name].iter_rows(values_only=True))
        for hi, row in enumerate(rows[:15]):
            head = [_norm(h) for h in row]
            if "كود الصنف" not in head:
                continue
            ci = head.index("كود الصنف")
            qcol = next((c for c in STOCK_COLS if c in head), None)
            qi = head.index(qcol) if qcol else None
            wcols = [(head.index(n), lbl) for n, (_, lbl) in zip(wnames, WAREHOUSES) if n in head]
            if qi is None and not wcols:
                continue
            log(f"تقرير المخزون: شيت '{name}'، الكمية من "
                + (f"'{qcol}'" if qi is not None else "مجموع المخازن الرئيسية"))
            ti = head.index("التصنيف") if "التصنيف" in head else None
            ni = head.index("اسم الصنف") if "اسم الصنف" in head else None
            if ti is None and ni is not None:
                log("مفيش عمود 'التصنيف' - الجلابية بتتعرف من اسم الصنف (أي اسم فيه 'ثوب' = ثوب)")
            return _read_rows(rows[hi + 1:], ci, qi, wcols, ti, ni)
    sys.exit("مش لاقي في تقرير المخزون عمود 'كود الصنف' ومعاه 'المستودعات الرئيسية' أو أعمدة المخازن")


def _read_rows(rows, ci, qi, wcols, ti=None, ni=None):
    out = {}
    for r in rows:
        if r is None or ci >= len(r) or r[ci] is None:
            continue
        code = str(r[ci]).strip().split(".")[0].lstrip("0")
        if not code.isdigit():
            continue                                   # Grand Total وغيره
        ws = {lbl: r[i] for i, lbl in wcols if i < len(r) and isinstance(r[i], (int, float))}
        if qi is not None:
            q = r[qi] if qi < len(r) and isinstance(r[qi], (int, float)) else 0
        else:
            q = sum(ws.values())
        out[code] = out.get(code, 0) + q
        if ti is not None and ti < len(r) and r[ti]:
            CAT[code] = _norm(r[ti])
        elif ti is None and ni is not None and ni < len(r) and r[ni]:
            CAT[code] = "ثوب" if "ثوب" in str(r[ni]) else "جلابيه"
        w = WH.setdefault(code, {})
        for lbl, v in ws.items():
            w[lbl] = w.get(lbl, 0) + v
    return out


SNAPSHOT = os.path.join(ROOT, "build", "stock_latest.xlsx")


def save_snapshot(stock):
    """نسخة مختصرة من آخر تقرير مخزون (الكود + التصنيف + المخازن الرئيسية بس، من غير أسعار)
    عشان رفع الصور لوحده يقدر يعيد بناء الموقع من غير تقرير جديد."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = STOCK_SHEET
    labels = [lbl for _, lbl in WAREHOUSES]
    ws.append(["كود الصنف", "التصنيف"] + [c for c, _ in WAREHOUSES] + [STOCK_COL])
    for code in sorted(stock):
        w = WH.get(code, {})
        ws.append([code, CAT.get(code, "")] + [w.get(l, 0) for l in labels] + [stock[code]])
        ws.cell(ws.max_row, 1).number_format = "@"
    wb.save(SNAPSHOT)


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


DATE_FILE = os.path.join(ROOT, "build", "stock_date.txt")


def stock_date(new_stock):
    """تاريخ آخر تحديث للمخزون. بيتغير لتاريخ النهارده بس لما يكون فيه تقرير مخزون جديد (--new-stock)،
    وغير كده بيفضل زي ما هو (تعديل الشكل أو الصور مش بيغيّر التاريخ)."""
    if new_stock or not os.path.exists(DATE_FILE):
        open(DATE_FILE, "w", encoding="utf-8").write(today() + "\n")
    return open(DATE_FILE, encoding="utf-8").read().strip()


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
    global NEW_STOCK
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    NEW_STOCK = "--new-stock" in sys.argv[1:]
    if len(args) != 3:
        sys.exit(__doc__)
    photos_dir, stock_xlsx, tree_xlsx = args
    # لو المصدر هو photos/ نفسه (مفيش فولدر الصور الأصلي)، ننسخه الأول لأن photos/ بيتمسح ويتبني من جديد
    if os.path.abspath(photos_dir).startswith(os.path.join(ROOT, "photos")):
        import tempfile
        tmp = tempfile.mkdtemp(prefix="ikaf-photos-")
        for fn in os.listdir(photos_dir):
            shutil.copy2(os.path.join(photos_dir, fn), tmp)
        photos_dir = tmp

    models, ratios = load_tree(tree_xlsx)
    materials = load_materials(tree_xlsx)
    colors = load_colors(tree_xlsx)
    stock = load_stock(stock_xlsx)
    if NEW_STOCK:
        save_snapshot(stock)
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
        rest = m.group(2).strip()
        photos.setdefault(m.group(1), []).append(((0 if not rest else 1), rest, fn))
    for c in photos:
        photos[c] = [x[-1] for x in sorted(photos[c])]
    log(f"الصور: {sum(len(v) for v in photos.values())} صورة لـ {len(photos)} كود")

    items, source = read_items()
    # أصناف الـPDF الأساسية محفوظة في build/pdf_items.json (build.py بيكتبه)، عشان الصنف اللي
    # اتشال عشان رصيده قل يرجع لوحده لما الرصيد يزيد
    base = os.path.join(ROOT, "build", "pdf_items.json")
    if os.path.exists(base):
        items = json.load(open(base, encoding="utf-8"))
    else:
        items = [it for it in items if it.get("s") not in ("photo", "pending")]
        for it in items:
            for k in ("g", "w", "n", "k", "cl"):
                it.pop(k, None)
        json.dump(items, open(base, "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    excluded = load_excluded()
    if excluded:
        log(f"أكواد متشالة بطلب المستخدم: {sorted(excluded)}")
    items = [it for it in items if it["c"] not in excluded]
    for it in items:
        it.pop("g", None)

    shutil.rmtree(os.path.join(ROOT, "photos"), ignore_errors=True)
    os.makedirs(os.path.join(ROOT, "photos", "full"))
    os.makedirs(os.path.join(ROOT, "photos", "thumb"))

    dropped = []
    # صور المستخدم أكد إنها مكررة (اسم الملف من غير امتداد): build/drop_photos.json
    dp = os.path.join(ROOT, "build", "drop_photos.json")
    drop_manual = set(json.load(open(dp, encoding="utf-8")).get("photos", [])) if os.path.exists(dp) else set()

    # الكود اللي متكرر على أكتر من صفحة في الـPDF بيظهر كارت واحد بس:
    # الصفحة التانية بتتشال لو نفس المجموعة اللونية، أو بتبقى صورة إضافية لو مجموعة مختلفة
    first, merged = {}, []
    for it in items:
        if it.get("p") is None:
            continue
        if it["c"] not in first:
            first[it["c"]] = it
            continue
        keep = first[it["c"]]
        mine = [keep["p"]] + keep.get("xp", [])
        sigs = [_sig(os.path.join(ROOT, "full", f"{pg:03d}.webp")) for pg in mine]
        if not any(_same_group(_sig(os.path.join(ROOT, "full", f"{it['p']:03d}.webp")), k) for k in sigs):
            keep.setdefault("xp", []).append(it["p"])
        merged.append((it["c"], it["p"]))
        it["_drop"] = True
    items = [it for it in items if not it.get("_drop")]
    if merged:
        log(f"أكواد متكررة في الـPDF اتجمعت في كارت واحد: {merged}")

    pages_of = {}
    for it in items:
        if it.get("p"):
            pages_of.setdefault(it["c"], []).extend([it["p"]] + it.get("xp", []))
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
            name = code if not names else f"{code}_{len(names) + 1}"   # اسم نضيف للموقع
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
        if it["q"] < MIN_QTY:
            it["_drop"] = True
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

    # جلابية رصيدها 20 أو أكتر ومالهاش صورة لسه: بتظهر "جاري التصوير" في آخر الصفحة
    pending, pend_skip = [], []
    if not CAT:
        log("!! تقرير المخزون مفيهوش عمود 'التصنيف' - أصناف 'جاري التصوير' مش هتتضاف")
    shown_codes = {it["c"] for it in items}
    for code in sorted(c for c, k in CAT.items() if k in PENDING_CATS):
        if code in shown_codes or code in excluded or stock.get(code, 0) < MIN_QTY:
            continue
        if not re.fullmatch(r"1\d{13}", code):
            pend_skip.append(code)                      # مش 14 رقم - مايتفكش من الشجرة
            continue
        rl, rn = ratio_label(code[8:10], ratios)
        wl, mn = width_label(code[4:6], models)
        if rl is None or mn is None:
            unknown.append((code, code[8:10], code[4:6]))
            continue
        q = stock[code]
        items.append({"p": None, "c": code, "r": rl + (wl or ""), "rn": rn,
                      "q": fmt_qty(q), "w": wh_split(code, q), "m": mn, "mc": code[4:6], "rc": code[8:10],
                      "u": code[:11], "s": "pending"})
        pending.append(code)

    # الخامة لكل صنف (خانة 2-3) — للفلتر في الصفحة
    nomat = sorted({it["c"][1:3] for it in items if it["c"][1:3] not in materials})
    for it in items:
        it["k"] = materials.get(it["c"][1:3], "")
        it["cl"] = colors.get(it["c"][10], "") if len(it["c"]) == 14 else ""
    if nomat:
        log(f"!! أكواد خامة مش في الشجرة: {nomat}")

    pdf_low = [(it["c"], it["q"]) for it in items if it.get("_drop")]
    items = [it for it in items if not it.get("_drop")]
    if pdf_low:
        log(f"أصناف PDF اتشالت عشان رصيدها أقل من {MIN_QTY}: {pdf_low}")

    # علامة "منتج جديد": build/new_products.json
    npf = os.path.join(ROOT, "build", "new_products.json")
    newp = {str(c).strip().lstrip("0") for c in json.load(open(npf, encoding="utf-8")).get("codes", [])} \
        if os.path.exists(npf) else set()
    for it in items:
        if it["c"] in newp:
            it["n"] = 1
        else:
            it.pop("n", None)
    miss_new = sorted(newp - {it["c"] for it in items})
    log(f"منتج جديد: {sum(1 for it in items if it.get('n'))} صنف"
        + (f" — مش ظاهرين في الكتالوج: {miss_new}" if miss_new else ""))

    # الصفحة بتتبني من القالب عشان أي تعديل فيه يوصل
    tpl = open(os.path.join(ROOT, "build", "template.html"), encoding="utf-8").read()
    tpl = tpl.replace("__DATA__", json.dumps(items, ensure_ascii=False)).replace("__SOURCE__", source)
    tpl = tpl.replace("__UPDATED__", stock_date(NEW_STOCK))
    open(os.path.join(ROOT, "index.html"), "w", encoding="utf-8").write(tpl)

    pdf_items = [it for it in items if it.get("s") not in ("photo", "pending")]
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
    log(f"جاري التصوير (جلابية من غير صورة، رصيد {MIN_QTY}+): {len(pending)}")
    if pend_skip:
        log(f"جلابية من غير صورة اتسابت عشان الكود مش 14 رقم: {pend_skip}")
    if dropped:
        log(f"صور مكررة اتشالت (نفس الصورة أو نفس المجموعة اللونية): {len(dropped)}")
    log(f"الإجمالي: {len(items)} صنف")


if __name__ == "__main__":
    main()
