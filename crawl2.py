# -*- coding: utf-8 -*-
"""Extension crawl: remaining electronics categories in Sanaa + finish mobiles pagination."""
import urllib.request, urllib.parse, re, html as H, json, time, gzip, random

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
BASE = "https://ye.opensooq.com"

def fetch(url, retries=2):
    last = None
    for i in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": UA, "Accept-Language": "ar,en;q=0.8",
                "Accept": "text/html,application/xhtml+xml,*/*;q=0.8"})
            with urllib.request.urlopen(req, timeout=45) as r:
                data = r.read()
                if r.headers.get("Content-Encoding") == "gzip":
                    data = gzip.decompress(data)
                return data.decode("utf-8", "ignore")
        except Exception as e:
            last = e
            time.sleep(3)
    return None

def parse_ads(page_html):
    ads = []
    for b in re.findall(r'<a class="postListItemData.*?</a>', page_html, re.S):
        def g(pat, default=""):
            m = re.search(pat, b)
            return H.unescape(m.group(1)).strip() if m else default
        ad = {
            "id": g(r'data-id1="(\d+)"'),
            "city": g(r'data-city="([^"]*)"'),
            "nhood": g(r'data-nhood="([^"]*)"'),
            "cat1": g(r'data-cat1-code="([^"]*)"'),
            "cat2": g(r'data-cat2-code="([^"]*)"'),
            "href": g(r'href="([^"]+)"'),
            "date_text": g(r'<span class="col-span-2[^"]*">([^<]+)</span>'),
            "price": g(r'text-\[#c62828\]">([^<]+)</span>'),
            "title": re.sub(r"<[^>]+>", "", g(r'<h2[^>]*>(.*?)</h2>')).strip(),
            "condition": g(r'alt="(جديد|مستعمل)"'),
            "subcat_ar": g(r'<span class="text-sm leading-4">([^<]+)</span>'),
        }
        if ad["id"]:
            ads.append(ad)
    return ads

data = json.load(open("sanaa_ads.json", encoding="utf-8"))
all_ads = {a["id"]: a for a in data["ads"]}
summary_extra = []

JOBS = [
    ("أجهزة إلكترونية وأجهزة منزلية", "صنعاء/اجهزة-الكترونيات", 60),
    ("ألعاب فيديو وملحقاتها", "صنعاء/العاب-فيديو-وملحقاتها", 8),
    ("الموبايلات والتابلت (استكمال)", "صنعاء/موبايل-تابلت", 52),
]
START_PAGE = {"صنعاء/موبايل-تابلت": 46}

for cat_name, path, maxp in JOBS:
    start = START_PAGE.get(path, 1)
    seen_local = 0
    pages = 0
    exhausted = False
    for page in range(start, maxp + 1):
        url = BASE + urllib.parse.quote("/ar/" + path, safe="/%") + (f"?page={page}" if page > 1 else "")
        htm = fetch(url)
        time.sleep(0.8 + random.random() * 0.6)
        if htm is None:
            break
        ads = parse_ads(htm)
        pages += 1
        new = 0
        for ad in ads:
            if ad["id"] not in all_ads:
                ad["category_top"] = cat_name.replace(" (استكمال)", "")
                all_ads[ad["id"]] = ad
                new += 1
            seen_local += 1
        print(f"{cat_name} p{page}: {len(ads)} ads ({new} new)", flush=True)
        if len(ads) == 0:
            exhausted = True
            break
    summary_extra.append({"category": cat_name, "pages": pages, "ads_seen": seen_local, "exhausted": exhausted})
    print(f"DONE {cat_name}: {seen_local} seen / {pages} pages, exhausted={exhausted}", flush=True)

ads_list = list(all_ads.values())
data["ads"] = ads_list
data["cat_summary"] = data.get("cat_summary", []) + summary_extra
json.dump(data, open("sanaa_ads.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("TOTAL unique ads:", len(ads_list), flush=True)
