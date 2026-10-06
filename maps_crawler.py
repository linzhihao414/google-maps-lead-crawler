"""
Google Maps Lead Crawler v2
Search keywords -> collect name, phone, website, address, email -> export Excel
Perfect for WhatsApp / B2B lead generation.
Edit keywords.txt to change search terms.
"""
import asyncio
import re
from datetime import datetime
from pathlib import Path

import httpx
from openpyxl import Workbook
from playwright.async_api import async_playwright

# ============ Config ============
KEYWORDS_FILE = Path(__file__).parent / "keywords.txt"
MAX_RESULTS = 100
OUTPUT_DIR = Path(__file__).parent / "导出结果"
HEADLESS = False
EXTRACT_EMAIL = True
# ================================


def load_keywords():
    if KEYWORDS_FILE.exists():
        text = KEYWORDS_FILE.read_text(encoding="utf-8").strip()
        return [k.strip() for k in re.split(r"[,，\n]", text) if k.strip()]
    return ["leather bag manufacturer"]


def clean_text(t):
    if not t:
        return ""
    # Remove icon glyphs
    return re.sub(r"[\ue000-\uf8ff]", "", t).strip()


def clean_phone(phone):
    if not phone:
        return ""
    p = re.sub(r"[^\d+]", "", phone)
    return p


def extract_email_from_website(url):
    if not url or "google.com" in url:
        return ""
    try:
        r = httpx.get(url, timeout=8, follow_redirects=True,
                      headers={"User-Agent": "Mozilla/5.0"})
        emails = re.findall(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', r.text)
        emails = [e for e in emails if not any(
            x in e.lower() for x in ['example.', 'sentry', 'wixpress', '.png', '.jpg', '.gif', '.webp']
        )]
        return emails[0] if emails else ""
    except:
        return ""


async def crawl_map(page, keyword):
    url = f"https://www.google.com/maps/search/{keyword}"
    await page.goto(url, wait_until="domcontentloaded", timeout=30000)
    await asyncio.sleep(8)

    # Cookie consent
    try:
        btn = await page.query_selector('button[aria-label*="Accept"]')
        if btn:
            await btn.click()
            await asyncio.sleep(3)
    except:
        pass

    panel = await page.query_selector('div[role="feed"]')
    if not panel:
        return []

    # Scroll to load more
    for _ in range(25):
        await panel.evaluate("el => el.scrollTop = el.scrollHeight")
        await asyncio.sleep(1.2)

    links = await panel.evaluate("""
        () => [...document.querySelectorAll('div[role="feed"] a[href*="/maps/place/"]')].map(a=>a.href)
    """)
    links = list(dict.fromkeys(links))[:MAX_RESULTS]
    print(f"  Found {len(links)} businesses", flush=True)

    results = []
    for i, link in enumerate(links, 1):
        try:
            await page.goto(link, wait_until="domcontentloaded", timeout=20000)
            await asyncio.sleep(2.5)

            data = await page.evaluate("""() => ({
                name: document.querySelector('h1')?.innerText?.trim() || '',
                address: document.querySelector('button[data-item-id="address"]')?.innerText?.trim() || '',
                website: document.querySelector('a[data-item-id="authority"]')?.href || '',
                phone: document.querySelector('button[data-item-id^="phone"]')?.innerText?.trim() || '',
                rating: document.querySelector('div.F7nice span[aria-hidden="true"]')?.innerText?.trim() || '',
            })""")

            phone = clean_phone(clean_text(data["phone"]))
            website = data["website"]
            email = extract_email_from_website(website) if EXTRACT_EMAIL and website else ""

            results.append({
                "keyword": keyword,
                "name": clean_text(data["name"]),
                "phone": phone,
                "whatsapp": f"https://wa.me/{phone.replace('+','')}" if phone else "",
                "email": email,
                "website": website,
                "address": clean_text(data["address"]),
                "rating": data["rating"],
                "maps_url": page.url,
                "collect_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
            })

            if i % 10 == 0:
                print(f"    [{i}/{len(links)}] {data['name'][:30]} | {phone}", flush=True)

            await asyncio.sleep(0.5)
        except Exception as e:
            print(f"    Skip: {e}", flush=True)

    return results


def save_to_excel(all_data, keyword):
    OUTPUT_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"

    headers = ["Keyword", "Business Name", "Phone", "WhatsApp", "Email",
               "Website", "Address", "Rating", "Maps URL", "Collected At"]
    ws.append(headers)

    for r in all_data:
        ws.append([r["keyword"], r["name"], r["phone"], r["whatsapp"], r["email"],
                    r["website"], r["address"], r["rating"], r["maps_url"], r["collect_time"]])

    for i, w in enumerate([15, 35, 18, 28, 30, 35, 40, 8, 50, 18], 1):
        ws.column_dimensions[chr(64 + i)].width = w

    fname = OUTPUT_DIR / f"GoogleMaps_{keyword}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(str(fname))
    print(f"  Saved: {fname}", flush=True)


async def main():
    keywords = load_keywords()
    print("=" * 50, flush=True)
    print("  Google Maps Lead Crawler v2", flush=True)
    print("=" * 50, flush=True)
    print(f"Keywords: {', '.join(keywords)}", flush=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=HEADLESS, channel="chrome")
        ctx = await browser.new_context(locale="en-US", viewport={"width": 1280, "height": 800})
        page = await ctx.new_page()

        total = 0
        for kw in keywords:
            print(f"\n=== {kw} ===", flush=True)
            try:
                results = await crawl_map(page, kw)
                save_to_excel(results, kw)
                total += len(results)
            except Exception as e:
                print(f"  Error: {e}", flush=True)

        print(f"\nDone! Total {total} leads.", flush=True)
        print(f"Output: {OUTPUT_DIR.resolve()}", flush=True)
        await browser.close()
        input("Press Enter to exit...")


if __name__ == "__main__":
    asyncio.run(main())
