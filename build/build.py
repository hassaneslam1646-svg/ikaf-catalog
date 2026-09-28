#!/usr/bin/env python3
"""
بناء كتالوج إكاف القابل للبحث بالكود.

المدخلات:
  1) ملف الكتالوج PDF  (صفحة غلاف + صفحة لكل صنف، فيها شريط برتقالي "Code : ...")
  2) ملف شجرة أكواد الجلابية .xlsx  (شيت ARABIC)

المخرجات (في جذر المستودع، جاهزة للرفع على GitHub Pages):
  index.html          صفحة البحث
  thumb/NNN.webp      صور مصغّرة (مقصوصة على صورة المنتج)
  full/NNN.webp       الصفحة كاملة
  out/labeled.pdf     نسخة الـPDF بعد كتابة الريشيو جنب شريط الكود
  out/review.xlsx     جدول مراجعة: الصفحة | الكود | الريشيو | اسم الموديل

التشغيل:
  python3 build/build.py "كتالوج.pdf" "شجرة الاكواد.xlsx"

المتطلبات:
  apt: poppler-utils  tesseract-ocr
  pip: pillow numpy openpyxl pypdf reportlab
"""

import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

import numpy as np
import openpyxl
from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

# ---------------------------------------------------------------------------
# ثوابت قالب الكتالوج — نِسب من أبعاد الصفحة، فمش بتتأثر بدقة التحويل (dpi)
# ---------------------------------------------------------------------------
BAND = (0.21926, 0.81583, 0.80889, 0.85417)   # الشريط البرتقالي بتاع الكود
CARD = (0.12267, 0.11500, 0.90044, 0.77900)   # الكارت الأبيض اللي فيه صورة المنتج

RENDER_DPI = 75          # دقة تحويل الصفحات لصور
FULL_W, FULL_Q = 600, 70 # الصفحة الكاملة (webp)
THUMB_W, THUMB_Q = 330, 72
LABEL_MAX_PT = 52.0      # أكبر حجم خط للريشيو المكتوب على الـPDF
LABEL_GAP_PT = 22.0      # المسافة بين الشريط والريشيو
LABEL_MARGIN_PT = 42.0   # هامش من حرف الصفحة

FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/google-fonts/Poppins-Bold.ttf",
    "/usr/share/fonts/truetype/Poppins-Bold.ttf",
    "build/Poppins-Bold.ttf",
]

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def log(msg):
    print(msg, flush=True)


def need(cmd):
    if shutil.which(cmd) is None:
        sys.exit(f"ناقص أداة مطلوبة: {cmd}")


# ---------------------------------------------------------------------------
# 1) شجرة الأكواد
# ---------------------------------------------------------------------------
def load_tree(xlsx):
    """يرجّع (الموديلات, الريشيوهات) — كل واحد dict من الكود للاسم."""
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    if "ARABIC" not in wb.sheetnames:
        sys.exit("شيت ARABIC مش موجود في ملف شجرة الأكواد")
    ws = wb["ARABIC"]

    def column(name_col, code_col, width):
        out = {}
        for r in range(4, ws.max_row + 1):
            name = ws[f"{name_col}{r}"].value
            code = ws[f"{code_col}{r}"].value
            if code is None or name is None or not str(name).strip():
                continue
            code = str(code).strip()
            if not re.fullmatch(r"\d+", code):
                continue
            out[code.zfill(width)] = str(name).strip()
        return out

    # G/H = الموديل، M/N = الريشيو
    return column("G", "H", 2), column("M", "N", 2)


def ratio_label(code, ratios):
    """كود الريشيو (خانة 9-10) -> الحرف المعروض، واسمه في الشجرة."""
    name = ratios.get(code)
    if name is None:
        return None, None
    name = name.strip()
    m = re.match(r"^مقاس\s+(.+)$", name)
    if m:
        return m.group(1).strip().upper(), name
    if name.lower().startswith("speichal"):
        return "SPECIAL", name
    if name == "Free Size":
        return "FREE", name
    return name.upper(), name


def width_label(code, models):
    """كود الموديل (خانة 5-6) -> العرض المستخرج من اسم الموديل، واسم الموديل."""
    name = models.get(code)
    if name is None:
        return None, None
    n = name.strip()
    if re.search(r"عرض\s*عاد", n):       # "عرض عادي" = من غير رقم عرض
        return "", n
    m = re.search(r"(\d{2})\s*[/:]\s*(\d{2})", n)   # مدى: 22:24 أو 32/34
    if m:
        return f"{m.group(1)}/{m.group(2)}", n
    m = re.search(r"عرض\s*\(?\s*(\d{2})", n)
    if m:
        return m.group(1), n
    return "", n


# ---------------------------------------------------------------------------
# 2) قراءة الأكواد من صور الصفحات
# ---------------------------------------------------------------------------
def rasterize(pdf, outdir):
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(
        ["pdftoppm", "-jpeg", "-jpegopt", "quality=88", "-r", str(RENDER_DPI),
         pdf, os.path.join(outdir, "p")],
        check=True,
    )
    pages = {}
    for f in sorted(os.listdir(outdir)):
        m = re.match(r"p-(\d+)\.jpg$", f)
        if m:
            pages[int(m.group(1))] = os.path.join(outdir, f)
    return pages


def has_band(im):
    """هل الصفحة فيها الشريط البرتقالي؟ (صفحة الغلاف مفيهاش)"""
    a = np.asarray(im.convert("RGB"))
    h, w, _ = a.shape
    sub = a[int(h * 0.70):int(h * 0.92)]
    r, g, b = sub[:, :, 0].astype(int), sub[:, :, 1].astype(int), sub[:, :, 2].astype(int)
    return int(((r > 200) & (g > 100) & (g < 200) & (b < 120)).sum()) > 500


def read_code(im, tmp):
    """OCR لشريط الكود — يرجّع 14 رقم أو None."""
    w, h = im.size
    box = (int(BAND[0] * w), int(BAND[1] * h), int(BAND[2] * w), int(BAND[3] * h))
    crop = im.crop(box).convert("L")
    crop = crop.resize((crop.width * 2, crop.height * 2), Image.LANCZOS)
    for threshold in (190, 160, 215):
        crop.point(lambda v, t=threshold: 0 if v > t else 255).save(tmp)
        res = subprocess.run(
            ["tesseract", tmp, "stdout", "--psm", "7",
             "-c", "tessedit_char_whitelist=Code:0123456789oO"],
            capture_output=True, text=True,
        )
        text = res.stdout.strip().replace("\n", "")
        m = re.search(r":\s*([0-9oO ]+)", text)
        digits = re.sub(r"\D", "", (m.group(1) if m else "").replace("o", "0").replace("O", "0"))
        if len(digits) == 14:
            return digits
    return None


# ---------------------------------------------------------------------------
# 3) كتابة الريشيو على الـPDF
# ---------------------------------------------------------------------------
def pick_font():
    for path in FONT_CANDIDATES:
        p = path if os.path.isabs(path) else os.path.join(ROOT, path)
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont("CatalogBold", p))
            return "CatalogBold"
    sys.exit("خط Poppins-Bold مش موجود — حطّه في build/Poppins-Bold.ttf")


def write_labels(src_pdf, labels, dest):
    font = pick_font()
    reader = PdfReader(src_pdf)
    writer = PdfWriter()
    page_w = float(reader.pages[0].mediabox.width)
    page_h = float(reader.pages[0].mediabox.height)

    x_start = BAND[2] * page_w + LABEL_GAP_PT
    avail = page_w - x_start - LABEL_MARGIN_PT
    y_mid = page_h - ((BAND[1] + BAND[3]) / 2) * page_h

    for i, page in enumerate(reader.pages, start=1):
        text = labels.get(i)
        if text:
            size = LABEL_MAX_PT
            while pdfmetrics.stringWidth(text, font, size) > avail and size > 18:
                size -= 0.5
            buf = io.BytesIO()
            c = canvas.Canvas(buf, pagesize=(page_w, page_h))
            c.setFont(font, size)
            c.setFillColorRGB(1, 1, 1)            # أبيض، زي باقي كتابة الشريط
            c.drawString(x_start, y_mid - size * 0.35, text)
            c.save()
            buf.seek(0)
            page.merge_page(PdfReader(buf).pages[0])
        writer.add_page(page)

    os.makedirs(os.path.dirname(dest), exist_ok=True)
    with open(dest, "wb") as f:
        writer.write(f)


# ---------------------------------------------------------------------------
# 4) صور الموقع
# ---------------------------------------------------------------------------
def build_images(labeled_pdf, pages_wanted, workdir):
    raster = rasterize(labeled_pdf, os.path.join(workdir, "labeled"))
    for folder in ("thumb", "full"):
        d = os.path.join(ROOT, folder)
        shutil.rmtree(d, ignore_errors=True)
        os.makedirs(d)
    for page in pages_wanted:
        im = Image.open(raster[page]).convert("RGB")
        w, h = im.size
        full = im.resize((FULL_W, round(im.height * FULL_W / im.width)), Image.LANCZOS)
        full.save(os.path.join(ROOT, "full", f"{page:03d}.webp"), "WEBP",
                  quality=FULL_Q, method=6)
        crop = im.crop((int(CARD[0] * w), int(CARD[1] * h),
                        int(CARD[2] * w), int(CARD[3] * h)))
        thumb = crop.resize((THUMB_W, round(crop.height * THUMB_W / crop.width)), Image.LANCZOS)
        thumb.save(os.path.join(ROOT, "thumb", f"{page:03d}.webp"), "WEBP",
                   quality=THUMB_Q, method=6)


# ---------------------------------------------------------------------------
# 5) الصفحة وجدول المراجعة
# ---------------------------------------------------------------------------
def write_page(items):
    tpl = os.path.join(ROOT, "build", "template.html")
    html = open(tpl, encoding="utf-8").read()
    if "__DATA__" not in html:
        sys.exit("قالب الصفحة build/template.html مفيهوش __DATA__")
    html = html.replace("__DATA__", json.dumps(items, ensure_ascii=False))
    html = re.sub(r"(<span class=\"sub\">)\s*\d+\s*(صنف</span>)",
                  rf"\g<1>{len(items)} \g<2>", html)
    open(os.path.join(ROOT, "index.html"), "w", encoding="utf-8").write(html)


def write_review(items, dest):
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "مراجعة الريشيو"
    ws.sheet_view.rightToLeft = True
    headers = ["رقم الصفحة", "الكود", "الريشيو المكتوب في الـPDF",
               "كود الريشيو (خانة 9-10)", "اسم الريشيو في الشجرة",
               "كود الموديل (خانة 5-6)", "اسم الموديل في الشجرة", "العرض المستخرج"]
    ws.append(headers)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3C88")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border

    for it in items:
        m = re.match(r"^(.*?)(\d{2}(?:/\d{2})?)$", it["r"])
        width = m.group(2) if m and m.group(1) else ""
        ws.append([it["p"], it["c"], it["r"], it["rc"], it["rn"], it["mc"], it["m"], width])

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for c in row:
            c.font = Font(name="Arial", size=11)
            c.border = border
            c.alignment = Alignment(horizontal="center", vertical="center")
        row[1].number_format = "@"
        row[2].font = Font(name="Arial", size=11, bold=True)
        row[4].alignment = Alignment(horizontal="right")
        row[6].alignment = Alignment(horizontal="right")

    for i, w in enumerate([12, 20, 22, 20, 22, 20, 42, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    wb.save(dest)


# ---------------------------------------------------------------------------
def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    pdf, xlsx = sys.argv[1], sys.argv[2]
    for p in (pdf, xlsx):
        if not os.path.exists(p):
            sys.exit(f"الملف مش موجود: {p}")
    need("pdftoppm")
    need("tesseract")

    models, ratios = load_tree(xlsx)
    log(f"شجرة الأكواد: {len(models)} موديل، {len(ratios)} ريشيو")

    workdir = tempfile.mkdtemp(prefix="ikaf-")
    try:
        pages = rasterize(pdf, os.path.join(workdir, "raw"))
        log(f"عدد الصفحات: {len(pages)}")

        tmp_png = os.path.join(workdir, "ocr.png")
        items, labels, unread, unknown = [], {}, [], []

        for page in sorted(pages):
            im = Image.open(pages[page])
            if not has_band(im):
                continue                                  # صفحة غلاف
            code = read_code(im, tmp_png)
            if code is None:
                unread.append(page)
                continue
            rl, rn = ratio_label(code[8:10], ratios)
            wl, mn = width_label(code[4:6], models)
            if rl is None or mn is None:
                unknown.append((page, code, code[8:10], code[4:6]))
                continue
            label = rl + (wl or "")
            labels[page] = label
            items.append({"p": page, "c": code, "r": label, "rn": rn,
                          "m": mn, "mc": code[4:6], "rc": code[8:10], "u": code[:11]})

        log(f"اتقرا بنجاح: {len(items)} صنف")
        if unread:
            log(f"!! صفحات الكود مش مقروء فيها: {unread}")
        if unknown:
            log("!! أكواد مش موجودة في الشجرة (اتسابت من غير ريشيو):")
            for page, code, rc, mc in unknown:
                log(f"   صفحة {page} — {code} — ريشيو {rc} / موديل {mc}")
        if unread or unknown:
            log("راجع الحالات دي قبل النشر.")

        labeled = os.path.join(ROOT, "out", "labeled.pdf")
        write_labels(pdf, labels, labeled)
        log(f"اتكتب: {labeled}")

        build_images(labeled, [it["p"] for it in items], workdir)
        write_page(items)
        write_review(items, os.path.join(ROOT, "out", "review.xlsx"))
        log("اتكتب: index.html + thumb/ + full/ + out/review.xlsx")
        log("\nخلص. للنشر:  git add -A && git commit -m 'تحديث الكتالوج' && git push")
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


if __name__ == "__main__":
    main()
