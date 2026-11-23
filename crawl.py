# -*- coding: utf-8 -*-
"""Crawl OpenSooq Yemen - Sanaa electronics-related categories. Polite crawler."""
import urllib.request, urllib.parse, re, html as H, json, time, gzip, random, sys

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
    print("  FAIL:", url, last, flush=True)
    return None

CATS = [
    ("الإلكترونيات", "/ar/صنعاء/الكترونيات"),
    ("الموبايلات والتابلت", "/ar/صنعاء/موبايل-تابلت"),
    ("الكمبيوتر واللابتوب", "/ar/صنعاء/لابتوب-وكمبيوتر"),
    ("ألعاب الفيديو وملحقاتها", "/ar/صنعاء/ألعاب-فيديو-وملحقاتها"),
]

def parse_ads(page_html):
    ads = []
    blocks = re.findall(r'<a class="postListItemData.*?</a>', page_html, re.S)
    for b in blocks:
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
            "title": g(r'<h2[^>]*>(.*?)</h2>'),
            "condition": g(r'alt="(جديد|مستعمل)"'),
            "subcat_ar": g(r'<span class="text-sm leading-4">([^<]+)</span>'),
        }
        ad["title"] = re.sub(r"<[^>]+>", "", ad["title"]).strip()
        if ad["id"]:
            ads.append(ad)
    return ads

all_ads = {}
cat_summary = []
for cat_name, path in CATS:
    seen_local = 0
    pages_fetched = 0
    for page in range(1, 46):
        q = urllib.parse.quote(path, safe="/")
        url = BASE + q + (f"?page={page}" if page > 1 else "")
        htm = fetch(url)
        pages_fetched += 1
        time.sleep(0.8 + random.random() * 0.7)
        if htm is None:
            break
        ads = parse_ads(htm)
        new = 0
        for ad in ads:
            if ad["id"] not in all_ads:
                ad["category_top"] = cat_name
                all_ads[ad["id"]] = ad
                new += 1
        seen_local += len(ads)
        print(f"{cat_name} p{page}: {len(ads)} ads ({new} new)", flush=True)
        if not ads or new == 0:
            break
    cat_summary.append({"category": cat_name, "pages": pages_fetched, "ads_seen": seen_local})
    print(f"DONE {cat_name}: {seen_local} ads seen over {pages_fetched} pages", flush=True)

ads_list = list(all_ads.values())
json.dump({"collected_utc": "2026-11-23", "cat_summary": cat_summary, "ads": ads_list},
          open("sanaa_ads.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("TOTAL unique ads:", len(ads_list), flush=True)

# category totals shown in header "(N)"
for cat_name, path in CATS:
    q = urllib.parse.quote(path, safe="/")
    htm = fetch(BASE + q)
    time.sleep(1)
    if htm:
        txt = re.sub(r"<[^>]+>", " ", H.unescape(htm[:6000]))
        m = re.search(r"صنعاء[^0-9]{0,80}([0-9,]{2,8})", txt)
        print("TOTAL", cat_name, "->", m.group(1) if m else "?", flush=True)
