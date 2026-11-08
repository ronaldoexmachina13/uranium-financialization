"""Download the World Bank Pink Sheet monthly price file into data/raw/worldbank/ and log it.

Run from the repo root:  python scripts/fetch_worldbank.py
The file's web address changes every month, so the script reads the current link from the
World Bank commodity page. If that fails, it falls back to the Nov 2026 link.
If you get a certificate error on Windows, run once:  pip install truststore
(it makes Python use the Windows certificate store).
Covers gold, silver, platinum, coal, copper, oil, natural gas and other commodities (monthly, period average).
"""
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
OUT = ROOT / "data" / "raw" / "worldbank"
LOG = ROOT / "DATA_LOG.md"
PAGE = "https://www.worldbank.org/en/research/commodity-markets"
FALLBACK = ("https://thedocs.worldbank.org/en/doc/74e8be41ceb20fa0da750cda2f6b9e4e-0050012026"
            "/related/CMO-Historical-Data-Monthly.xlsx")
HEADERS = {"User-Agent": "uranium-thesis-fetch/1.0"}


def get(url):
    with urllib.request.urlopen(urllib.request.Request(url, headers=HEADERS), timeout=120) as r:
        return r.read()


def current_link():
    try:
        html = get(PAGE).decode("utf-8", "ignore")
        m = re.search(r'https://thedocs\.worldbank\.org/[^"\']*CMO-Historical-Data-Monthly\.xlsx', html)
        return m.group(0) if m else FALLBACK
    except Exception:
        return FALLBACK


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    url = current_link()
    data = get(url)
    if data[:2] != b"PK":
        raise RuntimeError("Downloaded file is not an Excel file")
    path = OUT / "CMO-Historical-Data-Monthly.xlsx"
    path.write_bytes(data)
    today = dt.date.today().isoformat()
    line = (f"| {today} | World Bank Pink Sheet, monthly prices (gold, silver, coal, copper, oil, gas, etc.) "
            f"| World Bank | monthly, period average, {len(data) // 1024} KB | data/raw/worldbank/{path.name} |\n")
    with LOG.open("a", encoding="utf-8") as f:
        f.write(line)
    print(f"Saved {path} from {url}")


if __name__ == "__main__":
    main()
