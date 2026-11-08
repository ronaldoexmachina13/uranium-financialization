"""Download SPUT (Sprott Physical Uranium Trust) daily data from the Sprott fund page.

Run from the repo root:  python scripts/fetch_sput.py
The page has no download button, but its charts are written into the page as data. This script reads them.
If you get a certificate error on Windows, run once:  pip install truststore

Writes to data/raw/sprott/:
  sput_daily.csv      date, price_usd (U.U close), nav_usd, premium_discount_pct   (every trading day)
  sput_holdings.csv   date, pounds_u3o8_millions                                  (weekly, Fridays)
  sput_acquired.csv   date, pounds_acquired, cumulative_pounds_acquired           (purchase days only)
Note: the acquired series excludes the 18.1M lb already held at the start (16 Jul 2021).
"""
import csv
import datetime as dt
import re
import urllib.request
from pathlib import Path

try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "raw" / "sprott"
LOG = ROOT / "DATA_LOG.md"
URL = "https://sprott.com/investment-strategies/exchange-listed-products/physical-commodity-funds/uranium/"
UTC = r"Date\.UTC\((\d+),(\d+),(\d+)\)"  # note: the month is 0-based in JavaScript


def iso(y, m, d):
    return dt.date(int(y), int(m) + 1, int(d)).isoformat()


def parse(html):
    """Return {(chart id, series name): [(date, value), ...]}."""
    out = {}
    for block in html.split("renderTo: '")[1:]:
        cid = block[: block.index("'")]
        if "Placeholder" not in cid:
            continue
        head = block[: re.search(r"series:\s*\[", block).start()]
        cats = [iso(*m) for m in re.findall(UTC, head)]
        for part in re.split(r"\bname:\s*'", block)[1:]:
            name = part[: part.index("'")]
            rest = part[part.index("'"):]
            m = re.search(r"data:\s*\[([\s\S]*?)\]\s*(?:,\s*\n|\}|\n\s*\})", rest)
            if not m:
                continue
            body = m.group(1)
            pts = [(iso(a, b, c), float(v)) for a, b, c, v in re.findall(r"\[" + UTC + r",\s*(-?[\d.]+)\]", body)]
            if not pts:
                nums = re.findall(r"-?[\d.]+", body)
                pts = [(cats[i], float(v)) for i, v in enumerate(nums)]
            out[(cid, name)] = pts
    return out


def main():
    req = urllib.request.Request(URL, headers={"User-Agent": "uranium-thesis-fetch/1.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        html = r.read().decode("utf-8", "ignore")
    s = parse(html)
    price = dict(s[("Uranium_Market_Price_Vs_NAV_Graph_Placeholder", "Share Price")])
    nav = dict(s[("Uranium_Market_Price_Vs_NAV_Graph_Placeholder", "Net Asset Value")])
    pdc = dict(s[("Uranium_PDC_Graph_Placeholder", "Premium/Discount to Net Asset Value")])
    held = s[("SPUT_Uranium_MetalsHeldChart_Placeholder", "Pounds of U308 (millions)")]
    acq = dict(s[("SPUT_Uranium_MetalsAcquiredChart_Placeholder", "Daily Pounds Acquired")])
    cum = dict(s[("SPUT_Uranium_MetalsAcquiredChart_Placeholder", "Cumulative")])
    dates = sorted(price)
    if not (dates and set(dates) == set(nav) == set(pdc) and acq and held):
        raise RuntimeError("The page layout has changed: series missing or dates do not line up")
    if int(sum(acq.values())) != int(max(cum.values())):
        raise RuntimeError("Acquired pounds do not add up to the cumulative figure")

    OUT.mkdir(parents=True, exist_ok=True)

    def write(name, header, rows):
        with open(OUT / name, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(header)
            w.writerows(rows)

    write("sput_daily.csv", ["date", "price_usd", "nav_usd", "premium_discount_pct"],
          [(d, price[d], nav[d], pdc[d]) for d in dates])
    write("sput_holdings.csv", ["date", "pounds_u3o8_millions"], held)
    adates = sorted(acq)
    write("sput_acquired.csv", ["date", "pounds_acquired", "cumulative_pounds_acquired"],
          [(d, int(acq[d]), int(cum[d])) for d in adates])

    today = dt.date.today().isoformat()
    lines = [
        f"| {today} | SPUT price, NAV and premium/discount (USD, U.U) | Sprott fund page | daily, {len(dates)} rows, {dates[0]} to {dates[-1]} | data/raw/sprott/sput_daily.csv |\n",
        f"| {today} | SPUT pounds of U3O8 held | Sprott fund page | weekly, {len(held)} rows, {held[0][0]} to {held[-1][0]} | data/raw/sprott/sput_holdings.csv |\n",
        f"| {today} | SPUT pounds acquired and cumulative | Sprott fund page | purchase days, {len(adates)} rows, {adates[0]} to {adates[-1]} | data/raw/sprott/sput_acquired.csv |\n",
    ]
    with LOG.open("a", encoding="utf-8") as f:
        f.writelines(lines)
    print(f"Daily: {len(dates)} rows ({dates[0]} to {dates[-1]}); holdings: {len(held)}; purchase days: {len(adates)}")
    print(f"Last purchase: {adates[-1]}; pounds held now: {held[-1][1]}M lb")


if __name__ == "__main__":
    main()
