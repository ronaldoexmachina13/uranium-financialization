"""Phase 4: daily UX front-month futures settlement history from Stooq (symbol UX.F).

Why a browser: Stooq offers no CSV for UX.F (the /q/d/l/ endpoint returns 0 bytes
even in a normal browser) and serves pages behind a JavaScript browser check, so
plain HTTP requests receive a challenge page instead of data. With the author's
explicit approval (2026-09-29), pages are loaded in a VISIBLE Chrome window driven
by Playwright, which runs that check exactly as a manual visit would.
This is a documented exception to the pipeline's no-circumvention rule.

Stooq Terms of Service (reviewed 2026-09-29): no clause on automated access;
5.3 forbids redistribution without consent -> raw pages and the series are for
thesis analysis only and must not be published.

Usage
-----
    python src/download_stooq_ux.py            # fetch all daily pages, then parse
    python src/download_stooq_ux.py --parse    # parse stored pages only (no network)

Raw   -> data/raw/uranium/stooq_ux_pages/uranium_futures_UX_daily__stooq-UX.F__pageNNN__<YYYYMMDD>.html (rendered DOM)
Parsed-> data/processed/daily/uranium_futures_UX_price_daily__stooq.csv
"""
from __future__ import annotations

import argparse
import io
import random
import re
import sys
import time
from pathlib import Path

import pandas as pd

from common import PROCESSED, RAW, save_raw, log_retrieval

RAW_SUBDIR = "uranium/stooq_ux_pages"
PAGE_URL = "https://stooq.com/q/d/?s=ux.f&i=d&l={n}"
CHROME = Path.home() / "AppData/Local/Google/Chrome/Application/chrome.exe"
TABLE_ID = "fth1"
PAUSE_SECONDS = (3.0, 5.0)


# ----------------------------------------------------------------- fetch
def fetch_pages() -> None:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=str(CHROME), headless=False)
        page = browser.new_page()
        n, last = 1, None
        while last is None or n <= last:
            url = PAGE_URL.format(n=n)
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=60_000)
                page.wait_for_selector(f"#{TABLE_ID} tbody tr", timeout=60_000)
                html = page.content()
                if last is None:  # read page count from the '>>' pager link
                    m = re.findall(r"s=ux\.f&amp;i=d&amp;l=(\d+)", html)
                    last = max(map(int, m)) if m else 1
                    print(f"total pages: {last}")
                data = html.encode("utf-8")
                path, status = save_raw(data, RAW_SUBDIR, f"uranium_futures_UX_daily__stooq-UX.F__page{n:03d}", "html")
                log_retrieval("Uranium_Stooq_UX", "stooq (playwright, visible Chrome)", url,
                              path, data, status, f"page {n}/{last}")
                print(f"page {n}/{last}: {status}")
            except Exception as e:
                log_retrieval("Uranium_Stooq_UX", "stooq (playwright, visible Chrome)", url,
                              None, None, "failed", repr(e)[:300])
                print(f"page {n}: FAILED {e!r}", file=sys.stderr)
                break  # stop rather than hammer the site; rerun resumes safely (dedupe by hash)
            n += 1
            time.sleep(random.uniform(*PAUSE_SECONDS))
        browser.close()


# ----------------------------------------------------------------- parse
def parse_page(path: Path) -> pd.DataFrame:
    html = path.read_text(encoding="utf-8", errors="replace")
    t = pd.read_html(io.StringIO(html), attrs={"id": TABLE_ID})[0]
    t.columns = ["stooq_no", "date", "open", "high", "low", "close",
                 "change_pct", "change", "volume", "open_interest"]
    t["date"] = pd.to_datetime(t["date"], format="%d %b %Y")
    t["change_pct"] = pd.to_numeric(t["change_pct"].astype(str).str.rstrip("%"), errors="coerce")
    for c in ["stooq_no", "open", "high", "low", "close", "change", "volume", "open_interest"]:
        t[c] = pd.to_numeric(t[c], errors="coerce")
    return t.assign(source_file=path.name)


def parse_all() -> pd.DataFrame:
    files = sorted((RAW / RAW_SUBDIR).glob("uranium_futures_UX_daily__stooq-UX.F__page*.html"))
    if not files:
        raise SystemExit("no Stooq pages stored")
    # latest retrieval of each page number wins (all are kept on disk)
    latest: dict[str, Path] = {}
    for f in files:
        latest[re.search(r"__page(\d{3})__", f.name).group(1)] = f
    df = pd.concat([parse_page(f) for f in latest.values()], ignore_index=True)
    df = df.sort_values("stooq_no").reset_index(drop=True)

    n_expected = int(df["stooq_no"].max())
    checks = {
        "rows": len(df),
        "row numbers complete 1..N": set(df["stooq_no"]) == set(range(1, n_expected + 1)),
        "duplicate row numbers": int(df["stooq_no"].duplicated().sum()),
        "duplicate dates": int(df["date"].duplicated().sum()),
        "dates increase with row number": bool(df["date"].is_monotonic_increasing),
        "first date": df["date"].min().date(),
        "last date": df["date"].max().date(),
        "non-zero volume rows": int((df["volume"] > 0).sum()),
        "non-zero open-interest rows": int((df["open_interest"] > 0).sum()),
    }
    dest = PROCESSED / "daily"
    dest.mkdir(parents=True, exist_ok=True)
    df.to_csv(dest / "uranium_futures_UX_price_daily__stooq.csv", index=False)
    for k, v in checks.items():
        print(f"{k:32s} {v}")
    return df


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--parse", action="store_true", help="parse stored pages only")
    a = ap.parse_args()
    if not a.parse:
        fetch_pages()
    parse_all()
