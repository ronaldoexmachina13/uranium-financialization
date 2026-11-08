"""Add new quarters to data/raw/yellowcake/nav_quarterly.csv from Yellow Cake plc's quarterly operating updates.

Run from the repo root:  python scripts/fetch_yellowcake.py
Yellow Cake publishes NAV per share and the uranium held only in its quarterly operating updates (RNS), so the
series is quarterly. The announcements are listed on an ASP.NET page that is searched and paged with form posts.
The script searches for 'Quarterly', follows every result whose title looks like a quarterly update, reads
the quarter end, the NAV per share (GBP) and the pounds of U3O8 held, and ADDS quarters that are not yet in the file.
Existing rows are never changed. A quarter whose text the script cannot read is printed, so you can add it by hand.
New rows get lb_basis = stated, spot_usd_lb from data/raw/cameco/uranium_price_monthly.csv, and announcement_date from the article.
Wording of the announcements has changed over the years; if a new quarter is skipped, send me the printed line.
Standard library plus optional truststore (Windows certificates).
"""
import csv
import datetime as dt
import html as htmllib
import http.cookiejar
import re
import urllib.parse
import urllib.request
from pathlib import Path

try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "raw" / "yellowcake" / "nav_quarterly.csv"
CAMECO = ROOT / "data" / "raw" / "cameco" / "uranium_price_monthly.csv"
LOG = ROOT / "DATA_LOG.md"
LIST_URL = "https://otp.tools.investis.com/clients/uk/yellow_cake/rns/regulatory-news.aspx"
FIELDS = ["quarter_end", "nav_per_share_gbp", "uranium_held_million_lb", "lb_basis", "spot_usd_lb", "announcement_date"]
MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"]
DATE = r"(\d{1,2})(?:st|nd|rd|th)?\s+(" + "|".join(MONTHS) + r")\s+(20\d\d)"
MAX_PAGES = 15

jar = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
opener.addheaders = [("User-Agent", "uranium-thesis-fetch/1.0")]


def get(url, data=None):
    body = urllib.parse.urlencode(data).encode() if data else None
    return opener.open(url, body, timeout=90).read().decode("utf-8", "ignore")


def hidden(page):
    out = {}
    for name in ("__VIEWSTATE", "__VIEWSTATEGENERATOR", "__EVENTVALIDATION"):
        m = re.search(r'name="%s"[^>]*value="([^"]*)"' % name, page)
        if m:
            out[name] = htmllib.unescape(m.group(1))
    return out


def stories(page):
    """(title, absolute url) for each announcement link on a list page."""
    found = []
    for href, label in re.findall(r'<a[^>]+href="([^"]*regulatory-story[^"]*)"[^>]*>(.*?)</a>', page, flags=re.S | re.I):
        title = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", htmllib.unescape(label))).strip()
        found.append((title, urllib.parse.urljoin(LIST_URL, htmllib.unescape(href))))
    return found


def list_quarterly_updates():
    page = get(LIST_URL)
    data = {**hidden(page), "__EVENTTARGET": "", "__EVENTARGUMENT": "",
            "RNSCategoryFilter1$txtSearchFor": "Quarterly", "RNSCategoryFilter1$btnsubmit": "Search"}
    page = get(LIST_URL, data)
    links = {}
    for n in range(MAX_PAGES):
        for title, url in stories(page):
            links[url] = title
        nxt = re.search(r"__doPostBack\('(RNSViewControl1\$dpNumeric\$ctl00\$ctl0%d)'" % (n + 1), page)
        if not nxt:
            break
        data = {**hidden(page), "__EVENTTARGET": nxt.group(1), "__EVENTARGUMENT": "",
                "RNSCategoryFilter1$txtSearchFor": "Quarterly"}
        page = get(LIST_URL, data)
    return {u: t for u, t in links.items() if re.search(r"quarter|operating update|NAV", t, re.I)}


def read_article(url):
    text = re.sub(r"\s+", " ", htmllib.unescape(re.sub(r"<[^>]+>", " ", get(url))))
    q = re.search(r"(?:quarter|period)\s+ended?\s+(?:on\s+)?" + DATE, text, re.I) or re.search(r"as at\s+" + DATE, text, re.I)
    nav = re.search(r"NAV[^£]{0,120}?£\s?(\d+\.\d+)", text) or re.search(r"£\s?(\d+\.\d+)\s+per\s+share", text, re.I)
    lb = re.search(r"(\d{1,2}\.\d{1,2})\s*(?:million|m)\s*(?:lbs?|pounds)", text, re.I)
    qdate = q.group(0) if q else ""
    ann = next((m for m in re.finditer(DATE, text, re.I) if m.group(0) not in qdate), None)  # first date that is not the quarter end
    if not (q and nav):
        return None
    d, mo, y = int(q.group(1)), MONTHS.index(q.group(2).lower()) + 1, int(q.group(3))
    row = {"quarter_end": f"{y}-{mo:02d}-{d:02d}", "nav_per_share_gbp": nav.group(1),
           "uranium_held_million_lb": lb.group(1) if lb else ""}
    if ann:
        row["announcement_date"] = f"{ann.group(3)}-{MONTHS.index(ann.group(2).lower()) + 1:02d}-{int(ann.group(1)):02d}"
    return row


def main():
    have = {}
    if OUT.exists():
        for r in csv.DictReader(OUT.open(encoding="utf-8")):
            have[r["quarter_end"]] = r
    spot = {}
    if CAMECO.exists():
        for r in csv.DictReader(CAMECO.open(encoding="utf-8")):
            spot[r["date"][:7]] = r["spot_usd_lb"]
    added = skipped = 0
    for url, title in sorted(list_quarterly_updates().items(), key=lambda kv: kv[1]):
        try:
            row = read_article(url)
        except Exception as e:
            print(f"could not read '{title}' ({e})")
            skipped += 1
            continue
        if row is None:
            print(f"could not find quarter end and NAV in: {title}  [{urllib.parse.urlparse(url).path}]")
            skipped += 1
            continue
        if row["quarter_end"] in have:
            continue
        row.setdefault("announcement_date", "")
        row["lb_basis"] = "stated" if row["uranium_held_million_lb"] else ""
        row["spot_usd_lb"] = spot.get(row["quarter_end"][:7], "")
        have[row["quarter_end"]] = {k: row.get(k, "") for k in FIELDS}
        print("added", row["quarter_end"], "NAV", row["nav_per_share_gbp"], "lb", row["uranium_held_million_lb"] or "missing")
        added += 1
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with OUT.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, lineterminator="\n")
        w.writeheader()
        for q in sorted(have):
            w.writerow({k: have[q].get(k, "") for k in FIELDS})
    with LOG.open("a", encoding="utf-8") as f:
        f.write(f"| {dt.date.today().isoformat()} | Yellow Cake NAV per share and uranium held (quarterly operating updates), script run | {LIST_URL} | "
                f"quarterly, {len(have)} rows, {added} new, {skipped} announcements unread | data/raw/yellowcake/nav_quarterly.csv |\n")
    print(f"{len(have)} quarters in the file, {added} added, {skipped} announcements not read")


if __name__ == "__main__":
    main()
