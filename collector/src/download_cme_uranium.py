"""Phase 5: forward collection of CME/COMEX UxC Uranium U3O8 futures (UX) from the
public CME Daily Bulletin, plus a reproducible parser.

Why forward-only: CME publishes only the *current* bulletin free of charge
(https://www.cmegroup.com/daily_bulletin/current/); history is sold via DataMine.
Missed days cannot be recovered later, so run this every day (Task Scheduler).

Usage
-----
    python src/download_cme_uranium.py            # collect today's bulletin, then re-parse all
    python src/download_cme_uranium.py --parse    # re-parse stored PDFs only (no network)

Raw PDFs  -> data/raw/cme/daily_bulletin/CME_bulletin_UX__<YYYYMMDD>__<bulletin#>_<FINAL|PRELIMINARY>__<S62|S02B>.pdf
Parsed    -> data/processed/daily/uranium_futures_UX_by_contract_daily__CME.csv   (one row per trade date x contract)
             data/processed/daily/uranium_futures_UX_volume_OI_daily__CME.csv      (one row per trade date, with checks)

Terms: CME grants a personal, non-transferable licence to use the website and
forbids copying/distribution; the raw PDFs are for private research and must not
be redistributed. /daily_bulletin/ is not disallowed in cmegroup.com/robots.txt.
"""
from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber

from common import (PROCESSED, RAW, load_config, http_session, save_raw,
                    log_retrieval, sha256)

RAW_SUBDIR = "cme/daily_bulletin"
MONTH_RE = re.compile(r"^[A-Z]{3}\d{2}$")
DATE_RE = re.compile(r"(Mon|Tue|Wed|Thu|Fri|Sat|Sun),\s+([A-Z][a-z]{2})\s+(\d{1,2}),\s+(\d{4})")
BULLETIN_RE = re.compile(r"BULLETIN\s*#\s*(\d+)")

# Column bins, calibrated on bulletin #185 (2026-09-25) and validated by header positions.
# Numeric columns are right-aligned, so they are identified by the word's right edge (x1).
S62_HEADER = {"SETT.": 294.0, "PNT/PIT": 468.0, "INTEREST": 552.0}
S62_BINS = {  # name: (edge, lo, hi)
    "globex_open":   ("x0", 110, 190),
    "high_low":      ("x0", 190, 280),
    "settle":        ("x1", 280, 314),   # also requires x0 >= 280
    "chg_sign":      ("x0", 314, 322),
    "chg":           ("x1", 322, 352),
    "globex_volume": ("x1", 395, 446),
    "pnt_volume":    ("x1", 470, 500),
    "open_interest": ("x1", 520, 560),
    "oi_chg_sign":   ("x0", 560, 572),
    "oi_chg":        ("x1", 572, 600),
}
S02B_BINS = {
    "globex_volume":    ("x1", 180, 215),
    "other_volume":     ("x1", 215, 330),   # open outcry / PNT
    "total_volume":     ("x1", 330, 352),
    "open_interest":    ("x1", 395, 412),
    "oi_chg_sign":      ("x0", 418, 434),
    "oi_chg":           ("x1", 440, 472),
    "volume_52w_ago":   ("x1", 505, 530),
    "oi_52w_ago":       ("x1", 565, 590),
}


# ----------------------------------------------------------------- helpers
def words(page):
    return page.extract_words(x_tolerance=1.5, keep_blank_chars=False)


def rows_by_line(ws, tol=1.6):
    lines: list[list[dict]] = []
    for w in sorted(ws, key=lambda w: (w["top"], w["x0"])):
        if lines and abs(lines[-1][0]["top"] - w["top"]) <= tol:
            lines[-1].append(w)
        else:
            lines.append([w])
    return [sorted(l, key=lambda w: w["x0"]) for l in lines]


def assign(line, bins):
    out: dict[str, str] = {}
    for w in line:
        for name, (edge, lo, hi) in bins.items():
            if lo < w[edge] <= hi and not (name == "settle" and w["x0"] < 280):
                out[name] = (out.get(name, "") + " " + w["text"]).strip()
                break
    return out


def num(s: str | None, dash_is_zero: bool):
    """'----' is the bulletin's 'no trade' marker: 0 for volume, NaN for prices."""
    if s is None or s == "":
        return 0.0 if dash_is_zero else float("nan")
    if set(s) == {"-"}:
        return 0.0 if dash_is_zero else float("nan")
    if s == "UNCH":
        return 0.0
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return float("nan")


def signed(sign: str | None, val: float) -> float:
    return -val if sign == "-" else val


def header_info(pdf) -> dict:
    text = pdf.pages[0].extract_text() or ""
    m = DATE_RE.search(text)
    b = BULLETIN_RE.search(text)
    if not (m and b):
        raise ValueError("could not read bulletin date/number from page 1")
    trade_date = datetime.strptime(f"{m.group(2)} {m.group(3)} {m.group(4)}", "%b %d %Y").date()
    status = "FINAL" if "FINAL" in text else ("PRELIMINARY" if "PRELIM" in text.upper() else "UNKNOWN")
    return {"trade_date": trade_date, "bulletin_no": int(b.group(1)), "status": status}


# ----------------------------------------------------------------- parsers
def parse_s62(path: Path) -> tuple[pd.DataFrame, dict]:
    with pdfplumber.open(path) as pdf:
        info = header_info(pdf)
        for page in pdf.pages:
            ws = words(page)
            texts = " ".join(w["text"] for w in ws)
            if "UXC URANIUM U3O8 FUTURES" not in texts:
                continue
            for k, x0 in S62_HEADER.items():  # layout guard
                if not any(w["text"] == k and abs(w["x0"] - x0) < 3 for w in ws):
                    raise ValueError(f"S62 layout changed: header '{k}' not at x0~{x0}")
            lines = rows_by_line(ws)
            start = next(i for i, l in enumerate(lines)
                         if [w["text"] for w in l][:2] == ["UX", "FUT"])
            recs, total = [], None
            for l in lines[start + 1:]:
                t = [w["text"] for w in l]
                if t[0] == "TOTAL":
                    if t[1:3] != ["UX", "FUT"]:
                        raise ValueError(f"unexpected TOTAL line: {t}")
                    total = assign(l, S62_BINS)
                    # the TOTAL OI can sit ~1pt off the TOTAL label line
                    break
                if not MONTH_RE.match(t[0]):
                    # TOTAL figures sometimes print on their own line just above 'TOTAL'
                    total = assign(l, S62_BINS)
                    continue
                c = assign(l, S62_BINS)
                recs.append({
                    "contract": t[0],
                    "settle": num(c.get("settle"), False),
                    "settle_chg": signed(c.get("chg_sign"), num(c.get("chg"), False)),
                    "globex_volume": num(c.get("globex_volume"), True),
                    "pnt_volume": num(c.get("pnt_volume"), True),
                    "open_interest": num(c.get("open_interest"), True),
                    "oi_chg": signed(c.get("oi_chg_sign"), num(c.get("oi_chg"), True)),
                    "globex_open_raw": c.get("globex_open", ""),
                    "high_low_raw": c.get("high_low", ""),
                    "row_text_raw": " ".join(t),
                })
            df = pd.DataFrame(recs)
            if df.empty:
                raise ValueError("UX block found but no contract rows parsed")
            df["total_volume"] = df["globex_volume"] + df["pnt_volume"]
            tot = {"s62_total_oi": num((total or {}).get("open_interest"), True),
                   "s62_total_globex_volume": num((total or {}).get("globex_volume"), True),
                   "s62_total_pnt_volume": num((total or {}).get("pnt_volume"), True)}
            return df.assign(**info), {**info, **tot}
    raise ValueError("UX not found in Section 62")


def parse_s02b(path: Path) -> dict:
    with pdfplumber.open(path) as pdf:
        info = header_info(pdf)
        for page in pdf.pages:
            for l in rows_by_line(words(page)):
                t = [w["text"] for w in l]
                if t[:3] == ["UX", "UXC", "URANIUM"]:
                    c = assign(l, S02B_BINS)
                    return {**info,
                            "s02b_total_volume": num(c.get("total_volume"), True),
                            "s02b_open_interest": num(c.get("open_interest"), True),
                            "s02b_oi_chg": signed(c.get("oi_chg_sign"), num(c.get("oi_chg"), True)),
                            "s02b_volume_52w_ago": num(c.get("volume_52w_ago"), True),
                            "s02b_oi_52w_ago": num(c.get("oi_52w_ago"), True)}
    raise ValueError("UX not found in Section 02B")


# ----------------------------------------------------------------- collect
def collect() -> None:
    cfg = load_config()
    bc = cfg["cme_bulletin"]
    session = http_session(cfg)
    for sec, fname in bc["sections"].items():
        url = bc["base_url"] + fname
        try:
            r = session.get(url, timeout=cfg["http"]["timeout_seconds"])
            r.raise_for_status()
            data = r.content
            if not data.startswith(b"%PDF"):
                raise ValueError("response is not a PDF")
            tmp = RAW / RAW_SUBDIR / f"_incoming_{sec}.pdf"
            tmp.parent.mkdir(parents=True, exist_ok=True)
            tmp.write_bytes(data)
            with pdfplumber.open(tmp) as pdf:
                info = header_info(pdf)
            tmp.unlink()
            # Skip if this exact file is already stored under any name
            known = {sha256(p.read_bytes()) for p in (RAW / RAW_SUBDIR).glob(f"*{sec}.pdf")}
            if sha256(data) in known:
                log_retrieval(f"UX_{sec}", "cme_bulletin", url, None, data, "unchanged",
                              f"{info['trade_date']} #{info['bulletin_no']} {info['status']}")
                print(f"{sec}: unchanged ({info['trade_date']} {info['status']})")
                continue
            stem = f"CME_bulletin_UX__{info['trade_date']:%Y%m%d}__{info['bulletin_no']}_{info['status']}__{sec}"
            path, status = save_raw(data, RAW_SUBDIR, stem, "pdf", dated=False)
            log_retrieval(f"UX_{sec}", "cme_bulletin", url, path, data, status,
                          f"{info['trade_date']} #{info['bulletin_no']} {info['status']}")
            print(f"{sec}: {status} -> {path.name}")
        except Exception as e:
            log_retrieval(f"UX_{sec}", "cme_bulletin", url, None, None, "failed", repr(e)[:300])
            print(f"{sec}: FAILED {e!r}", file=sys.stderr)


# ----------------------------------------------------------------- parse all
def status_rank(s: str) -> int:
    return {"FINAL": 2, "PRELIMINARY": 1}.get(s, 0)


def parse_all() -> None:
    folder = RAW / RAW_SUBDIR
    contracts, totals, s02b, errors = [], [], [], []
    for p in sorted(folder.glob("*S62.pdf")):
        try:
            df, tot = parse_s62(p)
            contracts.append(df.assign(source_file=p.name))
            totals.append({**tot, "s62_file": p.name})
        except Exception as e:
            errors.append((p.name, repr(e)))
    for p in sorted(folder.glob("*S02B.pdf")):
        try:
            s02b.append({**parse_s02b(p), "s02b_file": p.name})
        except Exception as e:
            errors.append((p.name, repr(e)))
    for name, err in errors:
        print(f"PARSE ERROR {name}: {err}", file=sys.stderr)
    if not contracts:
        print("no Section 62 files parsed")
        return

    def best(df):  # FINAL beats PRELIMINARY for the same trade date
        return (df.assign(_r=df["status"].map(status_rank))
                  .sort_values(["trade_date", "_r"]).drop_duplicates("trade_date", keep="last")
                  .drop(columns="_r"))

    c = pd.concat(contracts, ignore_index=True)
    c = c[c.set_index(["trade_date", "status"]).index.isin(
        best(c[["trade_date", "status"]].drop_duplicates()).set_index(["trade_date", "status"]).index)]
    t = best(pd.DataFrame(totals))
    # order contracts by expiry month so 'first' = nearest (front) contract
    c["contract_month"] = pd.to_datetime(c["contract"].str.title(), format="%b%y").dt.to_period("M")
    c = c.sort_values(["trade_date", "contract_month"])
    agg = (c.groupby("trade_date")
             .agg(n_contracts=("contract", "size"),
                  ux_globex_volume=("globex_volume", "sum"),
                  ux_pnt_volume=("pnt_volume", "sum"),
                  ux_total_volume=("total_volume", "sum"),
                  ux_open_interest=("open_interest", "sum"),
                  front_contract=("contract", "first"),
                  front_settle=("settle", "first"))
             .reset_index())
    out = agg.merge(t, on="trade_date", how="left")
    if s02b:
        out = out.merge(best(pd.DataFrame(s02b)).drop(columns=["bulletin_no", "status"]),
                        on="trade_date", how="left")
    # Validation flags (never used to alter values)
    out["check_oi_vs_s62_total"] = out["ux_open_interest"] == out["s62_total_oi"]
    if "s02b_open_interest" in out:
        out["check_oi_vs_s02b"] = out["ux_open_interest"] == out["s02b_open_interest"]
        out["check_volume_vs_s02b"] = out["ux_total_volume"] == out["s02b_total_volume"]
    dest = PROCESSED / "daily"
    dest.mkdir(parents=True, exist_ok=True)
    c.to_csv(dest / "uranium_futures_UX_by_contract_daily__CME.csv", index=False)
    out.sort_values("trade_date").to_csv(dest / "uranium_futures_UX_volume_OI_daily__CME.csv", index=False)
    print(out.T.to_string(header=False) if len(out) == 1 else out.tail().to_string(index=False))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--parse", action="store_true", help="parse stored PDFs only")
    a = ap.parse_args()
    if not a.parse:
        collect()
    parse_all()
