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
LOCATION_FILE = Path(__file__).parent / "位置.txt"
HISTORY_FILE = Path(__file__).parent / "历史记录.csv"
MAX_RESULTS = 500
OUTPUT_DIR = Path(__file__).parent / "导出结果"
HEADLESS = False
EXTRACT_EMAIL = True
# ================================


def load_keywords():
    if KEYWORDS_FILE.exists():
        text = KEYWORDS_FILE.read_text(encoding="utf-8").strip()
        return [k.strip() for k in re.split(r"[,，\n]", text) if k.strip()]
    return ["leather bag manufacturer"]


def load_location():
    """Read 位置.txt: if filled, search only within that area."""
    if LOCATION_FILE.exists():
        loc = LOCATION_FILE.read_text(encoding="utf-8").strip()
        # Ignore comments
        lines = [l.strip() for l in loc.splitlines() if l.strip() and not l.strip().startswith("#")]
        if lines:
            return lines[0]
    return ""


def load_history():
    """Load already-collected business URLs/phones from history file."""
    seen = set()
    if HISTORY_FILE.exists():
        import csv
        with open(HISTORY_FILE, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) >= 3:
                    if row[0]:
                        seen.add(row[0])  # maps_url
                    if row[1]:
                        seen.add("phone:" + row[1])  # phone
    return seen


def save_history(results):
    """Append new results to history file."""
    import csv
    new_file = not HISTORY_FILE.exists()
    with open(HISTORY_FILE, "a", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        if new_file:
            writer.writerow(["maps_url", "phone", "name", "collect_time"])
        for r in results:
            writer.writerow([r["maps_url"], r["phone"], r["name"], r["collect_time"]])


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


def lead_score(phone, email, website, rating):
    s = 0
    reasons = []
    if phone:
        s += 40
        reasons.append("phone")
    if email:
        s += 30
        reasons.append("email")
    if website:
        s += 20
        reasons.append("website")
    # Rating bonus
    try:
        r = float(rating) if rating else 0
        if r >= 4.5:
            s += 10
            reasons.append("high-rated")
    except:
        pass
    return min(100, s), ",".join(reasons)


async def crawl_map(page, keyword, location="", seen=None):
    if seen is None:
        seen = set()
    if location:
        query = f"{keyword} in {location}"
    else:
        query = keyword
    url = f"https://www.google.com/maps/search/{query}"
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
    for _ in range(80):
        await panel.evaluate("el => el.scrollTop = el.scrollHeight")
        await asyncio.sleep(1.0)

    links = await panel.evaluate("""
        () => [...document.querySelectorAll('div[role="feed"] a[href*="/maps/place/"]')].map(a=>a.href)
    """)
    links = list(dict.fromkeys(links))[:MAX_RESULTS]
    print(f"  Found {len(links)} businesses", flush=True)

    results = []
    skipped = 0
    for i, link in enumerate(links, 1):
        try:
            # Skip already-collected
            if link in seen:
                skipped += 1
                continue

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

            # Skip if phone already in history
            if phone and ("phone:" + phone) in seen:
                skipped += 1
                continue

            website = data["website"]
            email = extract_email_from_website(website) if EXTRACT_EMAIL and website else ""
            score, reasons = lead_score(phone, email, website, data["rating"])

            results.append({
                "keyword": keyword,
                "location": location,
                "name": clean_text(data["name"]),
                "phone": phone,
                "whatsapp": f"https://wa.me/{phone.replace('+','')}" if phone else "",
                "email": email,
                "website": website,
                "address": clean_text(data["address"]),
                "rating": data["rating"],
                "score": score,
                "reasons": reasons,
                "maps_url": page.url,
                "collect_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
            })

            # Add to seen
            seen.add(page.url)
            if phone:
                seen.add("phone:" + phone)

            if i % 10 == 0:
                print(f"    [{i}/{len(links)}] {data['name'][:30]} | {phone} (skipped {skipped} old)", flush=True)

            await asyncio.sleep(0.5)
        except Exception as e:
            print(f"    Skip: {e}", flush=True)

    print(f"  New: {len(results)} | Skipped (already collected): {skipped}", flush=True)
    results.sort(key=lambda x: x["score"], reverse=True)
    return results


def save_to_excel(all_data, keyword):
    OUTPUT_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"

    headers = ["Score", "Keyword", "Location", "Business Name", "Phone", "WhatsApp", "Email",
               "Website", "Address", "Rating", "Reasons", "Maps URL", "Collected At"]
    ws.append(headers)

    for r in all_data:
        ws.append([r["score"], r["keyword"], r.get("location",""), r["name"], r["phone"], r["whatsapp"], r["email"],
                    r["website"], r["address"], r["rating"], r["reasons"], r["maps_url"], r["collect_time"]])

    # Style
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    fill = PatternFill("solid", fgColor="4472C4")
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
    ws.freeze_panes = "A2"

    widths = [8, 15, 20, 35, 18, 28, 30, 35, 40, 8, 20, 50, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    fname = OUTPUT_DIR / f"GoogleMaps_{keyword}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(str(fname))
    print(f"  Saved: {fname}", flush=True)


async def main():
    keywords = load_keywords()
    location = load_location()
    print("=" * 50, flush=True)
    print("  Google Maps Lead Crawler v3", flush=True)
    print("=" * 50, flush=True)
    print(f"Keywords: {', '.join(keywords)}", flush=True)
    if location:
        print(f"Location: {location}", flush=True)
    else:
        print("Location: (not set - global search)", flush=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=HEADLESS, channel="chrome")
        ctx = await browser.new_context(locale="en-US", viewport={"width": 1280, "height": 800})
        page = await ctx.new_page()

        seen = load_history()
        if seen:
            print(f"Loaded history: {len(seen)} already-collected entries (will skip)", flush=True)

        total = 0
        for kw in keywords:
            print(f"\n=== {kw} ===", flush=True)
            try:
                results = await crawl_map(page, kw, location, seen)
                if results:
                    save_to_excel(results, kw)
                    save_history(results)
                total += len(results)
            except Exception as e:
                print(f"  Error: {e}", flush=True)

        print(f"\nDone! Total {total} leads.", flush=True)
        print(f"Output: {OUTPUT_DIR.resolve()}", flush=True)
        await browser.close()
        input("Press Enter to exit...")


if __name__ == "__main__":
    asyncio.run(main())
