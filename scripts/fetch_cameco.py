"""Download Cameco's monthly uranium prices into data/raw/cameco/uranium_price_monthly.csv and log it.

Run from the repo root:  python scripts/fetch_cameco.py
Cameco publishes the month-end spot and long-term price as an industry average of the UxC and TradeTech
month-end prices. The table is in the page HTML. The whole history on the page is kept (from 1988).
Uses only the standard library. If you get a certificate error on Windows:  pip install truststore
"""
import csv
import datetime as dt
import re
import calendar
import urllib.request
from pathlib import Path

try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "raw" / "cameco" / "uranium_price_monthly.csv"
LOG = ROOT / "DATA_LOG.md"
URL = "https://www.cameco.com/invest/markets/uranium-price"
START = "1988/01"  # the page starts in 1988; the workbook starts at the first month in the file


def main():
    req = urllib.request.Request(URL, headers={"User-Agent": "uranium-thesis-fetch/1.0"})
    html = urllib.request.urlopen(req, timeout=60).read().decode("utf-8", "ignore")
    table = html[html.index("<table"):html.index("</table>")]
    rows = {}
    for tr in re.findall(r"<tr>(.*?)</tr>", table, flags=re.S):
        cells = [re.sub(r"<[^>]+>", "", c).strip() for c in re.findall(r"<td[^>]*>(.*?)</td>", tr, flags=re.S)]
        if len(cells) >= 3 and re.fullmatch(r"\d{4}/\d\d/\d\d", cells[0]) and cells[0][:7] >= START:
            rows[cells[0][:7]] = (cells[1], cells[2])  # a repeated month in the source keeps one row
    if len(rows) < 200:
        raise RuntimeError(f"only {len(rows)} months found; the page layout may have changed")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with OUT.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["date", "spot_usd_lb", "longterm_usd_lb"])
        for ym in sorted(rows):
            y, m = int(ym[:4]), int(ym[5:7])
            w.writerow([f"{y}-{m:02d}-{calendar.monthrange(y, m)[1]:02d}", *rows[ym]])
    first, last = min(rows).replace("/", "-"), max(rows).replace("/", "-")
    line = (f"| {dt.date.today().isoformat()} | Uranium month-end spot and long-term price (Cameco industry average of UxC and "
            f"TradeTech month-end prices), script run | {URL} | monthly, {len(rows)} rows, {first} to {last} | "
            "data/raw/cameco/uranium_price_monthly.csv |\n")
    with LOG.open("a", encoding="utf-8") as f:
        f.write(line)
    print(f"{len(rows)} months, {first} to {last}")


if __name__ == "__main__":
    main()
