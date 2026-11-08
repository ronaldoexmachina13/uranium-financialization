"""Download daily prices and volumes for uranium equities, ETFs and uranium vehicles from Yahoo Finance.

Run from the repo root:  python scripts/fetch_equities.py
If you get a certificate error on Windows, run once:  pip install truststore
Writes one file per ticker to data/raw/yahoo/<ticker>.csv:
  date, open, high, low, close, adj_close, volume
close is adjusted for splits only; adj_close is also adjusted for dividends.
Dates are the exchange's local trading date. Prices are in each listing's own currency (see the log).
"""
import csv
import datetime as dt
import json
import time
import urllib.error
import urllib.request
from pathlib import Path

try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "raw" / "yahoo"
LOG = ROOT / "DATA_LOG.md"
URL = "https://query1.finance.yahoo.com/v8/finance/chart/{t}?period1=0&period2={p2}&interval=1d&events=div%7Csplit"

TICKERS = {
    "URA": "Global X Uranium ETF",
    "URNM": "Sprott Uranium Miners ETF",
    "NLR": "VanEck Uranium and Nuclear ETF",
    "CCJ": "Cameco (NYSE)",
    "KAP.IL": "Kazatomprom GDR (London)",
    "U-U.TO": "Sprott Physical Uranium Trust, USD units (TSX)",
    "U-UN.TO": "Sprott Physical Uranium Trust, CAD units (TSX)",
    "YCA.L": "Yellow Cake plc (London, pence)",
    "UEC": "Uranium Energy Corp",
    "DNN": "Denison Mines (NYSE American)",
    "NXE": "NexGen Energy (NYSE)",
    "PDN.AX": "Paladin Energy (ASX)",
    "BOE.AX": "Boss Energy (ASX)",
    "DYL.AX": "Deep Yellow (ASX)",
    "SLX.AX": "Silex Systems (ASX)",
    "BWXT": "BWX Technologies",
    "1164.HK": "CGN Mining (Hong Kong)",
}


def get_json(url):
    req = urllib.request.Request(url, headers={"User-Agent": "uranium-thesis-fetch/1.0"})
    for attempt in range(4):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 429 and attempt < 3:
                time.sleep(5 * (attempt + 1))
                continue
            raise


def rows_from(payload):
    res = payload["chart"]["result"][0]
    offset = res["meta"].get("gmtoffset", 0)
    q = res["indicators"]["quote"][0]
    adj = res["indicators"]["adjclose"][0]["adjclose"]
    rows = []
    for i, ts in enumerate(res["timestamp"]):
        if q["close"][i] is None:
            continue
        day = (dt.datetime(1970, 1, 1) + dt.timedelta(seconds=ts + offset)).date().isoformat()
        rows.append([day, q["open"][i], q["high"][i], q["low"][i], q["close"][i], adj[i], q["volume"][i]])
    return rows, res["meta"].get("currency", "")


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    p2 = int(time.time())
    today = dt.date.today().isoformat()
    lines = []
    for t, desc in TICKERS.items():
        rows, cur = rows_from(get_json(URL.format(t=t, p2=p2)))
        with open(OUT / f"{t}.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["date", "open", "high", "low", "close", "adj_close", "volume"])
            w.writerows(rows)
        lines.append(f"| {today} | {t}: {desc} | Yahoo Finance | daily, {cur}, {len(rows)} rows, {rows[0][0]} to {rows[-1][0]} | data/raw/yahoo/{t}.csv |\n")
        print(f"{t}: {len(rows)} rows ({rows[0][0]} to {rows[-1][0]}, {cur})")
        time.sleep(1)
    with LOG.open("a", encoding="utf-8") as f:
        f.writelines(lines)


if __name__ == "__main__":
    main()
