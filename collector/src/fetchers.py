"""Source-specific fetch + light parse used by the download scripts.

fetch() returns the raw bytes exactly as received (for yahoo: the yfinance
DataFrame serialised to CSV, since yfinance does not expose the HTTP payload).
read_raw() parses a stored raw file into a tidy DataFrame [date, value, ...]
without altering values; it is reused later by clean_data.py.
"""
from __future__ import annotations

import io
import warnings

import pandas as pd
import requests

from common import load_config, http_session, save_raw, log_retrieval

FRED_URL = "https://fred.stlouisfed.org/graph/fredgraph.csv?id={id}"
EIA_URL = "https://www.eia.gov/dnav/pet/hist_xls/{id}d.xls"


def raw_stem(s: dict) -> str:
    """Readable raw-file stem: <name>__<source>-<id>, e.g. gold_futures__yahoo-GC-F."""
    return f"{s['name']}__{s['source']}-{s['id']}"


def identifier(s: dict) -> str:
    return {"fred": FRED_URL.format(id=s["id"]),
            "eia_xls": EIA_URL.format(id=s["id"]),
            "cboe": s.get("url", ""),
            "html": s.get("url", ""),
            "yahoo": f"yfinance:{s['id']}"}[s["source"]]


def fetch(s: dict, session: requests.Session, timeout: int) -> tuple[bytes, str]:
    src = s["source"]
    if src in ("fred", "cboe", "eia_xls", "html"):
        r = session.get(identifier(s), timeout=timeout)
        r.raise_for_status()
        if src in ("fred", "cboe") and r.text.lstrip().lower().startswith("<!doctype html"):
            raise ValueError("received HTML instead of CSV (series id not found?)")
        return r.content, {"eia_xls": "xls", "html": "html"}.get(src, "csv")
    if src == "yahoo":
        import yfinance as yf
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            df = yf.Ticker(s["id"]).history(period="max", auto_adjust=False, actions=False)
        if df.empty:
            raise ValueError("yfinance returned no rows")
        return df.to_csv().encode("utf-8"), "csv"
    raise ValueError(f"unknown source {src}")


def read_raw(s: dict, path) -> pd.DataFrame:
    """Parse a raw file to columns: date (datetime64, naive), value (float),
    plus any extra source columns. Missing markers become NaN; no filling."""
    src = s["source"]
    if src == "fred":
        df = pd.read_csv(path)
        df.columns = ["date", "value"]
        df["value"] = pd.to_numeric(df["value"], errors="coerce")  # FRED uses '.' / '' for missing
    elif src == "cboe":
        df = pd.read_csv(path)
        df.columns = [c.strip().lower() for c in df.columns]
        df = df.rename(columns={"close": "value"})
        df["date"] = pd.to_datetime(df["date"], format="%m/%d/%Y")
    elif src == "eia_xls":
        df = pd.read_excel(path, sheet_name="Data 1", skiprows=2, engine="xlrd")
        df.columns = ["date", "value"]
        df["value"] = pd.to_numeric(df["value"], errors="coerce")
    elif src == "html" and s["id"] == "cameco-uranium-price":
        # Table 0 = full monthly history. Early rows are dated the 1st of the month,
        # later rows the month-end; both denote that month's month-end assessment.
        t = pd.read_html(io.StringIO(open(path, encoding="utf-8").read()))[0]
        t.columns = ["date_raw", "value", "long_term"]
        df = t.assign(date=pd.to_datetime(t["date_raw"], format="%Y/%m/%d"),
                      value=pd.to_numeric(t["value"], errors="coerce"),
                      long_term=pd.to_numeric(t["long_term"], errors="coerce"))
        df["month"] = df["date"].dt.to_period("M")
    elif src == "yahoo":
        df = pd.read_csv(path)
        df = df.rename(columns={"Date": "date", "Close": "value"})
        # Yahoo stamps futures/indices at local exchange midnight (tz-aware); keep the calendar date.
        df["date"] = pd.to_datetime(df["date"].str[:10])
    else:
        raise ValueError(src)
    df["date"] = pd.to_datetime(df["date"])
    return df


def download_group(group: str) -> pd.DataFrame:
    cfg = load_config()
    session = http_session(cfg)
    timeout = cfg["http"]["timeout_seconds"]
    rows = []
    for s in (x for x in cfg["series"] if x["group"] == group):
        ident = identifier(s)
        try:
            data, ext = fetch(s, session, timeout)
            path, status = save_raw(data, s["raw_dir"], raw_stem(s), ext)
            log_retrieval(s["key"], s["source"], ident, path, data, status)
            df = read_raw(s, path).dropna(subset=["value"])
            rows.append({"key": s["key"], "id": s["id"], "status": status, "n_obs": len(df),
                         "first": df["date"].min().date(), "last": df["date"].max().date(),
                         "raw_file": path.name})
        except Exception as e:  # record the failure, keep going
            log_retrieval(s["key"], s["source"], ident, None, None, "failed", repr(e)[:300])
            rows.append({"key": s["key"], "id": s["id"], "status": f"FAILED: {e!r}"[:120]})
    return pd.DataFrame(rows)
