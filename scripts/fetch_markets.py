"""Download daily prices for the comparison commodities and the S&P 500 from Yahoo Finance (for month-end values).

Run from the repo root:  python scripts/fetch_markets.py
If you get a certificate error on Windows, run once:  pip install truststore
Writes one file per ticker to data/raw/markets/<ticker>.csv (same columns as fetch_equities.py) and appends to DATA_LOG.md.
Futures are Yahoo's continuous front-month series (the contract changes at each roll, prices are not back-adjusted).
XAUUSD=X and XAGUSD=X are spot gold and silver, kept next to the futures so the two can be compared.
A ticker that fails is reported and skipped; the others are still saved.
"""
import csv
import datetime as dt
import time
from pathlib import Path

import fetch_equities as fe  # same folder: reuses its download and parsing code

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "raw" / "markets"
MARKETS = {
    "GC=F": "Gold futures, COMEX front month, USD per troy oz",
    "XAUUSD=X": "Gold spot, USD per troy oz",
    "SI=F": "Silver futures, COMEX front month, USD per troy oz",
    "XAGUSD=X": "Silver spot, USD per troy oz",
    "HG=F": "Copper futures, COMEX front month, USD per lb",
    "PL=F": "Platinum futures, NYMEX front month, USD per troy oz",
    "CL=F": "WTI crude futures, NYMEX front month, USD per barrel",
    "BZ=F": "Brent crude futures, ICE front month, USD per barrel",
    "NG=F": "Natural gas futures (Henry Hub), NYMEX front month, USD per MMBtu",
    "^GSPC": "S&P 500 index (full history)",
}


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    p2 = int(time.time())
    today = dt.date.today().isoformat()
    lines, failed = [], []
    for t, desc in MARKETS.items():
        try:
            rows, cur = fe.rows_from(fe.get_json(fe.URL.format(t=t.replace("^", "%5E"), p2=p2)))
        except Exception as e:
            print(f"{t}: FAILED ({e})")
            failed.append(t)
            continue
        with open(OUT / f"{t}.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["date", "open", "high", "low", "close", "adj_close", "volume"])
            w.writerows(rows)
        lines.append(f"| {today} | {t}: {desc} | Yahoo Finance | daily, {cur}, {len(rows)} rows, {rows[0][0]} to {rows[-1][0]} | data/raw/markets/{t}.csv |\n")
        print(f"{t}: {len(rows)} rows ({rows[0][0]} to {rows[-1][0]}, {cur})")
        time.sleep(1)
    with fe.LOG.open("a", encoding="utf-8") as f:
        f.writelines(lines)
    if failed:
        print("Not saved:", ", ".join(failed))


if __name__ == "__main__":
    main()
