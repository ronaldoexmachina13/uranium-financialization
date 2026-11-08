"""Collect UF6 conversion prices (quarter-end) from Cameco's quarterly reports into
data/raw/cameco/conversion_prices_quarterly.csv and log the run.

Run from the repo root:  python scripts/fetch_conversion.py
Each Cameco quarterly report has a table 'Industry prices at quarter end' (conversion spot and long-term,
North America and Europe, USD per kgU as UF6; average of TradeTech and UxC). Each report shows 4 to 6
quarter-ends, so the script reads a list of reports and merges them with the existing file.
The SEC asks for a contact in the user agent. Before running, set it once in PowerShell, with your own details:
    $env:SEC_USER_AGENT = "Your Name your.email@example.com"
Two reports are PDFs (marked below). They need:  pip install pypdf   (without it those quarters keep the values already in the file)
Add newer reports to REPORTS when Cameco publishes them. Existing values are kept unless a report gives a different value, which is printed.
If a number looks wrong, check the printed 'uranium spot' line against uranium_price_monthly.csv first.
Standard library plus optional truststore (Windows certificates) and pypdf.
"""
import csv
import datetime as dt
import html as htmllib
import io
import os
import re
import urllib.request
from pathlib import Path

try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "raw" / "cameco" / "conversion_prices_quarterly.csv"
LOG = ROOT / "DATA_LOG.md"
S = "https://www.sec.gov/Archives/edgar/data/1009001/"
REPORTS = [  # (quarter the report covers up to, url, kind)
    ("2026-09", S + "000119312526439858/d106072dex992.htm", "html"),
    ("2026-06", "https://www.cameco.com/sites/default/files/documents/2026-Q2-MDA-FS-Notes_0.pdf", "pdf"),
    ("2025-09", "https://www.cameco.com/sites/default/files/documents/2025%20Q3%20MDA%20FS%20and%20Notes.pdf", "pdf"),
    ("2024-09", S + "000119312524252580/d814682dex992.htm", "html"),
    ("2023-09", S + "000119312523266724/d542228dex992.htm", "html"),
    ("2022-06", S + "000119312522202679/d381670dex992.htm", "html"),
    ("2021-03", S + "000119312521154085/d333418dex992.htm", "html"),
    ("2020-09", S + "000119312520285174/d29588dex992.htm", "html"),
    ("2020-06", S + "000119312520201910/d42842dex992.htm", "html"),
    ("2019-03", S + "000119312519131628/d705708dex992.htm", "html"),
    ("2018-09", S + "000119312518316636/d602124dex992.htm", "html"),
    ("2017-06", S + "000119312517237934/d403558dex992.htm", "html"),
    ("2016-03", S + "000119312516568367/d186902dex992.htm", "html"),
    ("2015-09", S + "000119312515360874/d60384dex992.htm", "html"),
    ("2014-06", S + "000119312514289754/d769759dex992.htm", "html"),
    ("2012-09", S + "000119312512446266/d432032dex992.htm", "html"),
    ("2011-06", "https://www.sec.gov/Archives/edgar/data/1009001/000095012311072768/o72778exv99w2.htm", "html"),
]
FIELDS = ["conv_spot_na_usd_kgu", "conv_spot_eu_usd_kgu", "conv_lt_na_usd_kgu", "conv_lt_eu_usd_kgu"]
MONTHS = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6, "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}
DATE = re.compile(r"\b(MAR|JUN|SEP|DEC)[A-Z]*\.?\s+(\d{1,2}),?\s+(20\d\d)\b", re.I)
PRICE = re.compile(r"(?<![\d.])\d{1,3}\.\d{2}(?![\d.])")


def fetch(url):
    ua = os.environ.get("SEC_USER_AGENT", "uranium-thesis-fetch/1.0")
    return urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": ua}), timeout=120).read()


def to_text(url, kind):
    raw = fetch(url)
    if kind == "pdf":
        try:
            from pypdf import PdfReader
        except ImportError:
            return None
        return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(raw)).pages)
    text = re.sub(r"<[^>]+>", " ", raw.decode("utf-8", "ignore"))
    return re.sub(r"\s+", " ", htmllib.unescape(text))


def parse(text):
    """Return {quarter_end: [uranium spot, uranium lt, conv spot na/eu, conv lt na/eu]} from the first usable table."""
    for m in re.finditer(r"Industry prices", text, flags=re.I):
        win = text[m.start():m.start() + 2600]
        dates = [(int(y), MONTHS[mo[:3].upper()], int(d)) for mo, d, y in DATE.findall(win[:500])]
        n = len(dates)
        if n < 4:
            continue
        nums = [float(x) for x in PRICE.findall(win[win.rfind(str(dates[-1][0])) :])]
        if len(nums) < 6 * n:
            continue
        rows = [nums[i * n:(i + 1) * n] for i in range(6)]
        if not all(3 < v < 250 for v in rows[0]):
            continue
        return {f"{y}-{mo:02d}-{d:02d}": [rows[r][i] for r in range(6)] for i, (y, mo, d) in enumerate(dates)}
    return None


def main():
    have = {}
    if OUT.exists():
        for r in csv.DictReader(OUT.open(encoding="utf-8")):
            have[r["quarter_end"]] = [r[f] for f in FIELDS]
    spot = {}  # uranium spot per quarter, printed for a sanity check
    changed = skipped = 0
    for upto, url, kind in REPORTS:
        try:
            text = to_text(url, kind)
        except Exception as e:  # keep going; report at the end
            print(f"{upto}: could not download ({e})")
            skipped += 1
            continue
        if text is None:
            print(f"{upto}: PDF skipped (pip install pypdf to read it)")
            skipped += 1
            continue
        table = parse(text)
        if not table:
            print(f"{upto}: table not found, check the report by hand")
            skipped += 1
            continue
        for q, v in table.items():
            new = [f"{x:.2f}" for x in v[2:]]
            spot[q] = v[0]
            if q in have and [float(x) for x in have[q]] != [float(x) for x in new]:
                print(f"{q}: value differs from the file, report {upto} says {new}, file has {have[q]}; keeping the file")
            elif q not in have:
                have[q] = new
                changed += 1
        print(f"{upto}: {len(table)} quarters read")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with OUT.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["quarter_end", *FIELDS])
        for q in sorted(have):
            w.writerow([q, *have[q]])
    with LOG.open("a", encoding="utf-8") as f:
        f.write(f"| {dt.date.today().isoformat()} | UF6 conversion prices (Cameco quarter-end tables), script run | Cameco quarterly reports | "
                f"quarterly, {len(have)} rows, {changed} new, {skipped} reports skipped | data/raw/cameco/conversion_prices_quarterly.csv |\n")
    print(f"{len(have)} quarters in the file, {changed} added, {skipped} reports skipped")


if __name__ == "__main__":
    main()
