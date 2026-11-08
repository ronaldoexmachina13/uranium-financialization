"""Download the FRED series used in the thesis into data/raw/fred/ and log each one.

Run from the repo root:  python scripts/fetch_fred.py
Uses only the standard library. Safe to re-run: files are overwritten, log lines are appended.
"""
import datetime as dt
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "raw" / "fred"
LOG = ROOT / "DATA_LOG.md"
URL = "https://fred.stlouisfed.org/graph/fredgraph.csv?id={id}&cosd=1990-01-01"

# id: (description, frequency / timing)
SERIES = {
    "PURANUSDM": ("Global price of uranium, USD per lb (IMF)", "monthly, period average"),
    "PCOALAUUSDM": ("Coal, Australia, USD per tonne (IMF)", "monthly, period average"),
    "PCOPPUSDM": ("Copper, USD per tonne (IMF)", "monthly, period average"),
    "PNGASUSUSDM": ("Natural gas, Henry Hub, USD per MMBtu (IMF)", "monthly, period average"),
    "DCOILWTICO": ("WTI crude oil, USD per barrel (EIA)", "daily"),
    "DCOILBRENTEU": ("Brent crude oil, USD per barrel (EIA)", "daily"),
    "VIXCLS": ("CBOE VIX", "daily close"),
    "DTWEXBGS": ("Broad US dollar index, nominal", "daily"),
    "DFF": ("Effective federal funds rate", "daily"),
    "DGS10": ("10-year Treasury yield", "daily"),
    "DFII10": ("10-year real Treasury yield (TIPS)", "daily"),
    "T10Y3M": ("10-year minus 3-month Treasury spread", "daily"),
    "BAA10Y": ("Moody's Baa corporate yield minus 10-year Treasury", "daily"),
    "T10YIE": ("10-year breakeven inflation", "daily"),
    "DEXCAUS": ("CAD per USD exchange rate", "daily"),
    "SP500": ("S&P 500 index (FRED keeps about 10 years only)", "daily close"),
}


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    today = dt.date.today().isoformat()
    lines = []
    for sid, (desc, timing) in SERIES.items():
        req = urllib.request.Request(URL.format(id=sid), headers={"User-Agent": "uranium-thesis-fetch/1.0"})
        with urllib.request.urlopen(req, timeout=60) as r:
            text = r.read().decode("utf-8")
        if not text.startswith("observation_date"):
            raise RuntimeError(f"{sid}: unexpected response")
        (OUT / f"{sid}.csv").write_text(text, encoding="utf-8")
        n = len(text.strip().splitlines()) - 1
        lines.append(f"| {today} | {sid}: {desc} | FRED | {timing}, {n} rows | data/raw/fred/{sid}.csv |\n")
        print(f"{sid}: {n} rows")
    with LOG.open("a", encoding="utf-8") as f:
        f.writelines(lines)


if __name__ == "__main__":
    main()
