# -*- coding: utf-8 -*-
"""Analyze scraped Sanaa ads and build Excel workbook."""
import json, re, statistics
from collections import Counter, defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

data = json.load(open("sanaa_ads.json", encoding="utf-8"))
ads = data["ads"]

BRANDS = [
    ("آبل / iPhone", [r"iphone", r"آيفون", r"ايفون", r"macbook", r"ماك بوك", r"ماكبوك", r"ipad", r"آيباد", r"ايباد", r"airpods", r"ايربودز", r"apple"]),
    ("سامسونج", [r"samsung", r"سامسونج", r"galaxy", r"جالكسي", r"جلاكسي"]),
    ("شاومي / ريدمي / بوكو", [r"xiaomi", r"شاومي", r"redmi", r"ريدمي", r"poco", r"بوكو", r"POCO"]),
    ("هواوي", [r"huawei", r"هواوي", r"honor", r"هونر"]),
    ("إنفينكس", [r"infinix", r"انفينكس", r"إنفينكس", r"إنفنيكس"]),
    ("تكنو", [r"tecno", r"تكنو"]),
    ("ريلمي", [r"realme", r"ريلمي"]),
    ("أوبو / فيفو", [r"oppo", r"أوبو", r"اوبو", r"vivo", r"فيفو"]),
    ("نوكيا", [r"nokia", r"نوكيا"]),
    ("HP", [r"\bhp\b", r"إتش بي", "اتش بي"]),
    ("Dell", [r"dell", r"ديل"]),
    ("Lenovo", [r"lenovo", r"لينوفو", r"thinkpad"]),
    ("Asus", [r"asus", r"أسوس", r"اوسوس", r"rog", r"tuf"]),
    ("Acer", [r"acer", r"ايسر", r"أيسر"]),
    ("MSI", [r"msi"]),
    ("بلايستيشن", [r"playstation", r"بلايستيشن", r"بلستيشن", r"\bps4\b", r"\bps5\b", r"\bps3\b", r"سوني"]),
    ("Xbox", [r"xbox", r"إكس بوكس", r"اكس بوكس"]),
    ("Nintendo", [r"nintendo", r"نينتندو", r"switch"]),
    ("DJI", [r"dji", r"\bmavic", r"\bmini\s?[234]\s?pro"]),
    ("Canon / Nikon", [r"canon", r"كانون", r"nikon", r"نيكون"]),
    ("إنفرتر شمسي", [r"انفرتر", r"إنفرتر", r"inverter", r"أوبتي", r"اوبتي", r"opf?", r"growatt", r"كانفيرتر", r"هيبرد"]),
    ("بطاريات ليثيوم", [r"ليثيوم", r"بطاري", r"lithium", r"\blifepo", r"ديب سايكل", r"tubular"]),
    ("ألواح شمسية", [r"ألواح", r"الواح", r"سولار", r"شمسي", r"panel"]),
    ("راوترات ومودمات", [r"راوتر", r"router", r"مودم", r"modem", r"starlink", r"ستارلينك", r"tp-?link", r"تيداتا"]),
    ("سماعات", [r"سماعات", r"سماعات", r"headphone", r"earbud", r"jbl", r"سبيكر", r"speaker"]),
    ("ساعات ذكية", [r"ساعة ذكية", r"ساعات ذكية", r"smart ?watch", r"سمارت", r"watch"]),
]

def detect_brand(title):
    t = title.lower()
    for name, pats in BRANDS:
        for p in pats:
            if re.search(p, t, re.I):
                return name
    return ""

def parse_price(p):
    m = re.sub(r"[^\d]", "", p or "")
    return int(m) if m else None

for a in ads:
    a["brand"] = detect_brand(a["title"])
    a["price_num"] = parse_price(a["price"])
    a["url"] = "https://ye.opensooq.com" + (a["href"] or f"/ar/search/{a['id']}")

# ---- stats ----
cat_stats = {}
for cat in sorted(set(a["category_top"] for a in ads)):
    ca = [a for a in ads if a["category_top"] == cat]
    prices = [a["price_num"] for a in ca if a["price_num"]]
    subc = Counter(a["subcat_ar"] or "غير محدد" for a in ca)
    cond = Counter(a["condition"] or "غير محدد" for a in ca)
    cat_stats[cat] = {
        "count": len(ca),
        "new": cond.get("جديد", 0), "used": cond.get("مستعمل", 0),
        "with_price": len(prices),
        "median": int(statistics.median(prices)) if prices else None,
        "min": min(prices) if prices else None, "max": max(prices) if prices else None,
        "top_subcats": subc.most_common(12),
    }

brand_counter = Counter(a["brand"] for a in ads if a["brand"])

# subcategory detail (cat2 codes)
sub_stats = defaultdict(lambda: {"n": 0, "new": 0, "used": 0})
for a in ads:
    k = (a["category_top"], a["subcat_ar"] or "غير محدد")
    sub_stats[k]["n"] += 1
    if a["condition"] == "جديد": sub_stats[k]["new"] += 1
    elif a["condition"] == "مستعمل": sub_stats[k]["used"] += 1

# ---- Excel ----
wb = Workbook()
HDR_FILL = PatternFill("solid", fgColor="1F4E5F")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
def style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HDR_FILL; cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"

# Sheet 1: raw data
ws = wb.active; ws.title = "البيانات الخام"
cols = ["رقم الإعلان", "اسم المنتج", "القسم الرئيسي", "الفئة الفرعية (بالموقع)", "رمز الفئة الفرعية",
        "السعر (ريال)", "الحالة", "المدينة", "المنطقة", "تاريخ النشر (كما هو معروض)",
        "عدد المفضلات", "المشاهدات", "رابط الإعلان"]
ws.append(cols)
for a in sorted(ads, key=lambda x: x["category_top"]):
    ws.append([a["id"], a["title"], a["category_top"], a["subcat_ar"], a["cat2"],
               a["price_num"] if a["price_num"] else (a["price"] or "غير محدد"),
               a["condition"] or "غير محدد", a["city"], a["nhood"], a["date_text"],
               "غير متاح للعامة", "غير متاح للعامة", a["url"]])
style_header(ws, len(cols))
for c, w in zip(range(1, len(cols)+1), [13, 45, 22, 24, 22, 12, 10, 10, 12, 16, 15, 12, 46]):
    ws.column_dimensions[get_column_letter(c)].width = w
ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"

# Sheet 2: category summary
ws2 = wb.create_sheet("ملخص الفئات")
ws2.append(["القسم الرئيسي", "عدد الإعلانات", "جديد", "مستعمل", "إعلانات بسعر معلن",
            "السعر الوسيط (ريال)", "أدنى سعر", "أعلى سعر", "عدد المفضلات"])
for cat, s in cat_stats.items():
    ws2.append([cat, s["count"], s["new"], s["used"], s["with_price"],
                s["median"] or "—", s["min"] or "—", s["max"] or "—", "غير متاح للعامة"])
style_header(ws2, 9)
for c, w in zip(range(1, 10), [28, 13, 9, 10, 16, 17, 11, 11, 16]):
    ws2.column_dimensions[get_column_letter(c)].width = w
ws2.auto_filter.ref = f"A1:I{ws2.max_row}"

# Sheet 3: subcategories
ws3 = wb.create_sheet("الفئات الفرعية")
ws3.append(["القسم الرئيسي", "الفئة الفرعية", "عدد الإعلانات", "جديد", "مستعمل", "عدد المفضلات"])
for (cat, sub), s in sorted(sub_stats.items(), key=lambda kv: -kv[1]["n"]):
    ws3.append([cat, sub, s["n"], s["new"], s["used"], "غير متاح للعامة"])
style_header(ws3, 6)
for c, w in zip(range(1, 7), [26, 30, 13, 9, 10, 16]):
    ws3.column_dimensions[get_column_letter(c)].width = w
ws3.auto_filter.ref = f"A1:F{ws3.max_row}"

# Sheet 4: brands
ws4 = wb.create_sheet("الماركات في العناوين")
ws4.append(["الماركة / نوع المنتج", "عدد الإعلانات التي تذكرها (من 3358)", "النسبة %", "عدد المفضلات"])
for b, n in brand_counter.most_common():
    ws4.append([b, n, round(100*n/len(ads), 1), "غير متاح للعامة"])
style_header(ws4, 4)
for c, w in zip(range(1, 5), [26, 32, 10, 16]):
    ws4.column_dimensions[get_column_letter(c)].width = w
ws4.auto_filter.ref = f"A1:D{ws4.max_row}"

# Sheet 5: methodology / limitations
ws5 = wb.create_sheet("منهجية وقيود")
notes = [
    ["البند", "التفاصيل"],
    ["تاريخ جمع البيانات", "23 نوفمبر 2026 (UTC)"],
    ["المصدر", "ye.opensooq.com — إعلانات مدينة صنعاء"],
    ["الأقسام المفحوصة", "الإلكترونيات + الموبايلات والتابلت + الكمبيوتر واللابتوب + أجهزة إلكترونية وأجهزة منزلية + ألعاب الفيديو وملحقاتها"],
    ["عدد صفوف الإعلانات المفحوصة أثناء التصفح", "4,934 صفاً عبر 139 صفحة نتائج (قبل إزالة التكرار بين الأقسام)"],
    ["عدد الإعلانات الفريدة بعد إزالة التكرار (بمعرف الإعلان)", "3,358"],
    ["عدد الإعلانات التي أمكن استخراج المفضلة منها", "0 — الموقع لا يعرض عدّاد المفضلة للزائر غير المسجّل"],
    ["دليل تقني", "حقل favorites_count موجود في بيانات صفحة الإعلان لكن قيمته فارغة لغير المسجّلين (فُحص 11 إعلاناً عشوائياً من كل الفئات)"],
    ["عدد المشاهدات", "غير معروض أيضاً للزائر العادي في صفحات الإعلانات اليمنية"],
    ["إزالة التكرار", "تمت بمعرف الإعلان data-id1؛ الإعلان الواحد الظاهر في أكثر من قسم حُسب مرة واحدة"],
    ["السعر", "كما هو معلن في الإعلان؛ بعض الإعلانات بلا سعر"],
    ["تاريخ النشر", "نص نسبي كما يعرضه الموقع (مثل: قبل ساعة) — الموقع لا يعرض تاريخاً مطلقاً للعامة"],
    ["الأخلاقيات", "التزام بسرعة تصفح مهذبة، دون تجاوز أي حماية أو تسجيل دخول"],
]
for row in notes:
    ws5.append(row)
style_header(ws5, 2)
ws5.column_dimensions["A"].width = 42; ws5.column_dimensions["B"].width = 95
for r in range(2, ws5.max_row+1):
    ws5.cell(row=r, column=2).alignment = Alignment(wrap_text=True, vertical="top")

wb.save("OpenSooq_Sanaa_Electronics_2026-11-23.xlsx")
print("Excel saved")

json.dump({"cat_stats": cat_stats, "brands": brand_counter.most_common(),
           "sub_top": [[k[0], k[1], v["n"], v["new"], v["used"]] for k, v in sorted(sub_stats.items(), key=lambda kv: -kv[1]["n"])[:30]],
           "total": len(ads)},
          open("stats.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("stats saved")
