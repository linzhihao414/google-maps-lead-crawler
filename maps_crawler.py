"""
Google Maps Lead Crawler
Search keywords on Google Maps -> collect business name, phone, website, address -> export Excel
Edit keywords.txt to change search terms.
"""
import asyncio
import re
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from playwright.async_api import async_playwright

# ============ Config ============
KEYWORDS_FILE = Path(__file__).parent / "keywords.txt"
MAX_RESULTS = 100
OUTPUT_DIR = Path(__file__).parent / "output"
HEADLESS = False
# ================================


def load_keywords():
    if KEYWORDS_FILE.exists():
        text = KEYWORDS_FILE.read_text(encoding="utf-8").strip()
        return [k.strip() for k in re.split(r"[,，\n]", text) if k.strip()]
    return ["leather bag manufacturer"]


async def crawl_map(page, keyword):
    """Search Google Maps and collect business info"""
    url = f"https://www.google.com/maps/search/{keyword}"
    await page.goto(url, wait_until="domcontentloaded", timeout=30000)
    await asyncio.sleep(5)

    # Dismiss cookie consent if appears
    try:
        consent = await page.query_selector('button[aria-label*="Accept"], button[aria-label*="同意"]')
        if consent:
            await consent.click()
            await asyncio.sleep(2)
    except:
        pass

    # Scroll results panel to load more
    results_panel = await page.query_selector('div[role="feed"]')
    if not results_panel:
        print(f"  No results panel for: {keyword}")
        return []

    print(f"  Loading results...")
    for i in range(20):
        await results_panel.evaluate("el => el.scrollTop = el.scrollHeight")
        await asyncio.sleep(1.5)
        if i % 5 == 0:
            count = await results_panel.evaluate("el => el.children.length")
            print(f"    Loaded {count} businesses...")

    # Get all business links
    links = await results_panel.evaluate("""
        () => {
            const items = document.querySelectorAll('div[role="feed"] > div > div > a[href*="/maps/place/"]');
            return Array.from(items).map(a => a.href);
        }
    """)
    links = list(dict.fromkeys(links))[:MAX_RESULTS]
    print(f"  Found {len(links)} businesses, collecting details...")

    results = []
    for i, link in enumerate(links, 1):
        try:
            await page.goto(link, wait_until="domcontentloaded", timeout=20000)
            await asyncio.sleep(2)

            # Name
            name = await page.evaluate("""
                () => {
                    const h = document.querySelector('h1');
                    return h ? h.innerText.trim() : '';
                }
            """)

            # Rating and reviews
            rating = await page.evaluate("""
                () => {
                    const el = document.querySelector('div.F7nice span[aria-hidden="true"]');
                    return el ? el.innerText.trim() : '';
                }
            """)

            # Address
            address = await page.evaluate("""
                () => {
                    const el = document.querySelector('button[data-item-id="address"]');
                    return el ? el.innerText.trim() : '';
                }
            """)

            # Website
            website = await page.evaluate("""
                () => {
                    const el = document.querySelector('a[data-item-id="authority"]');
                    return el ? el.href : '';
                }
            """)

            # Phone
            phone = await page.evaluate("""
                () => {
                    const el = document.querySelector('button[data-item-id^="phone:tel:"]');
                    if (!el) return '';
                    return el.innerText.trim() || el.getAttribute('data-item-id').replace('phone:tel:', '');
                }
            """)

            results.append({
                "keyword": keyword,
                "name": name,
                "phone": phone,
                "website": website,
                "address": address,
                "rating": rating,
                "maps_url": page.url,
                "collect_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
            })

            if i % 10 == 0:
                print(f"    [{i}/{len(links)}] {name[:30]} | {phone}")

            await asyncio.sleep(0.5)

        except Exception as e:
            print(f"    Skip: {e}")
            continue

    return results


def save_to_excel(all_data, keyword):
    OUTPUT_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"

    headers = ["Keyword", "Business Name", "Phone", "Website", "Address",
               "Rating", "Maps URL", "Collected At"]
    ws.append(headers)

    for r in all_data:
        ws.append([
            r["keyword"], r["name"], r["phone"], r["website"], r["address"],
            r["rating"], r["maps_url"], r["collect_time"]
        ])

    widths = [15, 35, 20, 35, 40, 8, 50, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i)].width = w

    fname = OUTPUT_DIR / f"GoogleMaps_{keyword}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(str(fname))
    print(f"  Saved: {fname}")


async def main():
    keywords = load_keywords()
    print("=" * 50)
    print("  Google Maps Lead Crawler")
    print("=" * 50)
    print(f"Keywords: {', '.join(keywords)}")
    print()

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=HEADLESS, channel="chrome")
        context = await browser.new_context(
            locale="en-US",
            viewport={"width": 1280, "height": 800},
        )
        page = await context.new_page()

        total = 0
        for kw in keywords:
            print(f"\n=== Searching: {kw} ===")
            try:
                results = await crawl_map(page, kw)
                print(f"  Collected {len(results)} leads")
                save_to_excel(results, kw)
                total += len(results)
            except Exception as e:
                print(f"  Error: {e}")

        print(f"\nDone! Total {total} leads.")
        print(f"Output: {OUTPUT_DIR.resolve()}")
        await browser.close()
        input("Press Enter to exit...")


if __name__ == "__main__":
    asyncio.run(main())
