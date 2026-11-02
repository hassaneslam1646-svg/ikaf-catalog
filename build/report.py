# تقرير بعد كل تحديث: بيقارن اللينك قبل وبعد ويكتب الفرق
# الاستخدام: python3 build/report.py old_index.html new_index.html [build_log.txt]
# التقرير بيظهر في صفحة Actions (Summary) ومش بيتنشر على اللينك
import json, re, sys


def load(path):
    try:
        h = open(path, encoding="utf-8").read()
    except OSError:
        return [], [], ""
    def blk(i):
        m = re.search(r'<script id="%s" type="application/json">(.*?)</script>' % i, h, re.S)
        try:
            return json.loads(m.group(1)) if m else []
        except ValueError:
            return []
    dt = " ".join(re.findall(r'class="dt">([^<]*)', h))
    return blk("data"), blk("lowdata"), dt


def key(it):
    return f"ثوب {it.get('mn')} {it.get('b')} {it.get('cl')}" if it.get("t") else it["c"]


def name(it):
    if it.get("t"):
        return f"ثوب موديل {it.get('mn')} - {it.get('b')} - {it.get('cl')}"
    return f"{it['c']} - {it.get('m','')} - {it.get('cl','')} - {it.get('r','')}"


def unit(it):
    return "قطعة" if it.get("t") else "درزن"


def main():
    old_d, old_l, old_dt = load(sys.argv[1])
    new_d, new_l, new_dt = load(sys.argv[2])
    blog = ""
    if len(sys.argv) > 3:
        try:
            blog = open(sys.argv[3], encoding="utf-8").read().strip()
        except OSError:
            pass

    def cards(d):
        return {key(x): x for x in d if x.get("s") != "pending"}

    def pend(d):
        return {x["c"]: x for x in d if x.get("s") == "pending"}

    oc, nc = cards(old_d), cards(new_d)
    op, np_ = pend(old_d), pend(new_d)
    ol = {x["c"]: x for x in old_l}
    nl = {x["c"]: x for x in new_l}

    jal = [x for x in nc.values() if not x.get("t")]
    th = [x for x in nc.values() if x.get("t")]
    out = []
    w = out.append
    w("# 📊 تقرير تحديث لينك المخزون")
    w(f"**التاريخ على اللينك:** {new_dt or '-'}  (اللي قبله: {old_dt or '-'})\n")

    w("## الأرقام")
    w("| البند | دلوقتي | قبل | الفرق |")
    w("|---|---|---|---|")
    def row(lbl, a, b):
        df = a - b
        w(f"| {lbl} | {a} | {b} | {'+' if df > 0 else ''}{df if df else '—'} |")
    ojal = [x for x in oc.values() if not x.get("t")]
    oth = [x for x in oc.values() if x.get("t")]
    row("كروت الجلابيات", len(jal), len(ojal))
    row("جاري التصوير", len(np_), len(op))
    row("جدول أقل من 20 درزن", len(nl), len(ol))
    row("الثوب الرسمي (موديل)", sum(1 for x in th if x.get("tk") == "r"), sum(1 for x in oth if x.get("tk") == "r"))
    row("الثوب المودرن (موديل)", sum(1 for x in th if x.get("tk") == "m"), sum(1 for x in oth if x.get("tk") == "m"))
    row("منتج جديد", sum(1 for x in new_d if x.get("n")), sum(1 for x in old_d if x.get("n")))
    row("مخفض", sum(1 for x in new_d + new_l if x.get("d")), sum(1 for x in old_d + old_l if x.get("d")))
    row("إجمالي الجلابيات (درزن)", sum(x["q"] for x in jal) + sum(x["q"] for x in np_.values()) + sum(x["q"] for x in nl.values()),
        sum(x["q"] for x in ojal) + sum(x["q"] for x in op.values()) + sum(x["q"] for x in ol.values()))
    row("إجمالي الثوب (قطعة)", sum(x["q"] for x in th), sum(x["q"] for x in oth))
    w("")

    def where(c):
        if c in nc: return "كارت"
        if c in np_: return "جاري التصوير"
        if c in nl: return "جدول أقل من 20"
        return "مش ظاهر (خلص أو أقل من درزن)"

    added = [nc[k] for k in nc if k not in oc]
    removed = [oc[k] for k in oc if k not in nc]
    w(f"## ✅ كروت اتضافت ({len(added)})")
    if added:
        w("| الصنف | الكمية | كان فين |")
        w("|---|---|---|")
        for x in added:
            was = "جاري التصوير" if x["c"] in op else ("جدول أقل من 20" if x["c"] in ol else "جديد")
            w(f"| {name(x)} | {x['q']} {unit(x)} | {was} |")
    else:
        w("مفيش")
    w(f"\n## ❌ كروت اتشالت ({len(removed)})")
    if removed:
        w("| الصنف | كانت | راحت فين |")
        w("|---|---|---|")
        for x in removed:
            w(f"| {name(x)} | {x['q']} {unit(x)} | {where(x['c']) if not x.get('t') else 'مش ظاهر'} |")
    else:
        w("مفيش")

    ch = [(nc[k], oc[k]["q"], nc[k]["q"]) for k in nc if k in oc and nc[k]["q"] != oc[k]["q"]]
    ch.sort(key=lambda t: -abs(t[2] - t[1]))
    up = sum(1 for _, a, b in ch if b > a)
    w(f"\n## 🔄 كميات اتغيرت في الكروت ({len(ch)}: زادت {up} / قلت {len(ch) - up})")
    if ch:
        w("| الصنف | قبل | دلوقتي | الفرق |")
        w("|---|---|---|---|")
        for x, a, b in ch[:40]:
            w(f"| {name(x)} | {a} | {b} | {'+' if b > a else ''}{b - a} {unit(x)} |")
        if len(ch) > 40:
            w(f"\n…و{len(ch) - 40} صنف تاني بفرق أصغر")
    else:
        w("مفيش")

    la = [nl[c] for c in nl if c not in ol]
    lr = [ol[c] for c in ol if c not in nl]
    w(f"\n## 📋 جدول أقل من 20 درزن: دخل {len(la)} / خرج {len(lr)}")
    if la or lr:
        w("<details><summary>التفاصيل</summary>\n")
        for x in la:
            w(f"- ➕ {name(x)} ({x['q']} درزن)")
        for x in lr:
            w(f"- ➖ {name(x)} ← {where(x['c'])}")
        w("\n</details>")

    w(f"\n## 📷 جاري التصوير ({len(np_)}) — أكواد محتاجة صورة")
    if np_:
        w("<details><summary>اللستة</summary>\n")
        for x in sorted(np_.values(), key=lambda x: -x["q"]):
            w(f"- {name(x)} ({x['q']} درزن){' 🆕' if x['c'] not in op else ''}")
        w("\n</details>")

    if blog:
        w("\n## 🛠️ سجل البناء (أسباب الصور اللي مظهرتش وغيره)")
        w("<details><summary>افتح</summary>\n\n```\n" + blog[-60000:] + "\n```\n</details>")
    print("\n".join(out))


if __name__ == "__main__":
    main()
