"""
fetch_futures_dated.py  (written 15 Nov 2026; replaces the one-fixed-offset date rule for the futures files)

What it does, in plain words
  Downloads the daily history of the seven futures series from Yahoo Finance and
  stores it in NEW files, with each bar given its true calendar date.
  Why: the older script added ONE fixed time-zone offset to every bar, so bars
  stamped at local midnight fell on the previous day in US summer time.
  Here each bar is converted with the exchange time zone (with summer time).

What it never does
  - It never overwrites a file (files are opened in mode "x", which fails if the file exists).
  - It never fills, interpolates, or drops a bar silently. Bars without a close are counted and reported.
  - On any surprise it stops with a message. There is no fallback.

How to run (from the repo root, with Ronaldo)
  python scripts/fetch_futures_dated.py diagnose GC=F     # prints, saves nothing
  python scripts/fetch_futures_dated.py fetch             # all seven, saves new files
"""
import bisect
import csv
import datetime as dt
import hashlib
import json
import os
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path
from zoneinfo import ZoneInfo          # needs the 'tzdata' package on Windows (it is in the lock file)

import pandas as pd

TICKERS = ["CL=F", "BZ=F", "GC=F", "HG=F", "NG=F", "SI=F", "PL=F"]
# Same request as scripts/fetch_equities.py line 27, so only the date handling differs.
URL = ("https://query1.finance.yahoo.com/v8/finance/chart/{t}"
       "?period1=0&period2={p2}&interval=1d&events=div%7Csplit")
USER_AGENT = "uranium-thesis-fetch/1.0"
OUT_DIR = Path(os.environ.get("THESIS_RAW_OUT", "data/raw/markets_dated"))
COLLECTOR_DIR = Path(os.environ.get("THESIS_COLLECTOR", "collector/data/raw/commodities"))
OLD_DIR = Path(os.environ.get("THESIS_OLD_RAW", "data/raw/markets"))     # the existing (shifted) files, read only, for sanity checks
COLLECTOR_FILES = {"CL=F": "oil_WTI_futures__yahoo-CL-F__*.csv", "GC=F": "gold_futures__yahoo-GC-F__*.csv",
                   "HG=F": "copper_futures__yahoo-HG-F__*.csv", "SI=F": "silver_futures__yahoo-SI-F__*.csv"}
EXPECTED_LOCAL_TIME = "00:00"      # ASSUMPTION A3: Yahoo stamps daily futures bars at local midnight. Confirmed or refuted by 'diagnose'.
NEGATIVE_CLOSE_ALLOWED = {("CL=F", "2020-04-20")}   # the one real negative settlement (CFTC release 8315-20)
PAUSE_SECONDS = 3                  # polite pause between tickers
# Early-close days (day after Thanksgiving, 24 Dec, 3 Jul, 31 Dec 2008): Yahoo stamps these bars at the session open
# (09:30, or 07:20 in 2013-2014) instead of midnight. They are accepted ONLY if every test in off_time_check() passes,
# and every one of them is listed on screen and in a small log file. Anything else off midnight still stops the run.
OFF_TIME_LATEST_HOUR = 12          # an accepted off-time stamp must be in the morning (before 12:00 local)
OFF_TIME_MAX_GAP_DAYS = 5          # calendar days to the previous and the next bar (long weekends with a holiday reach 4)


class Stop(Exception):
    """Raised for every problem. The top level prints it and ends the program."""


# ---------------------------------------------------------------- network (the ONLY network lines)
def http_get(url):
    """NETWORK. Returns (bytes, headers). Retries only HTTP 429 (too many requests), three times."""
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    for attempt in range(1, 4):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                body = r.read()
                headers = dict(r.headers.items())
        except urllib.error.HTTPError as e:
            if e.code == 429 and attempt < 3:
                time.sleep(10 * attempt)
                continue
            raise Stop(f"HTTP {e.code} from {url}. Stopping; nothing saved.")
        except (urllib.error.URLError, TimeoutError) as e:
            raise Stop(f"Network error for {url}: {e}. Stopping; nothing saved.")
        want = headers.get("Content-Length")
        if want is not None and int(want) != len(body):
            raise Stop(f"Partial download: header says {want} bytes, got {len(body)}. Nothing saved.")
        return body, headers
    raise Stop("Too many HTTP 429 answers. Wait a while and run again.")


def url_for(ticker):
    return URL.format(t=ticker, p2=int(time.time()))


# ---------------------------------------------------------------- parsing (no network)
def parse_chart(raw_bytes):
    """Turns Yahoo's JSON into (meta, timestamps, fields). Every assumed field is checked here (A1, A2)."""
    try:
        payload = json.loads(raw_bytes.decode("utf-8"))
    except (ValueError, UnicodeDecodeError) as e:
        raise Stop(f"The reply is not valid JSON (truncated or an error page?): {e}")
    chart = payload.get("chart")
    if not isinstance(chart, dict):
        raise Stop(f"No 'chart' key. Top-level keys found: {list(payload)}")
    if chart.get("error"):
        raise Stop(f"Yahoo reports an error: {chart['error']}")
    result = chart.get("result")
    if not result or len(result) != 1:
        raise Stop("Expected exactly one item in chart.result.")
    res = result[0]
    for key in ("meta", "timestamp", "indicators"):
        if key not in res:
            raise Stop(f"chart.result[0] lacks '{key}'. Keys found: {list(res)}")
    meta, ts = res["meta"], res["timestamp"]
    for key in ("symbol", "currency", "exchangeTimezoneName", "gmtoffset"):
        if key not in meta:
            raise Stop(f"meta lacks '{key}' (assumption A1 fails). meta keys found: {sorted(meta)}")
    try:
        ZoneInfo(meta["exchangeTimezoneName"])
    except Exception:
        raise Stop(f"meta.exchangeTimezoneName = {meta['exchangeTimezoneName']!r} is not a time zone this computer knows.")
    ind = res["indicators"]
    try:
        q = ind["quote"][0]
        adj = ind["adjclose"][0]["adjclose"]
        fields = {k: q[k] for k in ("open", "high", "low", "close", "volume")}
    except (KeyError, IndexError, TypeError) as e:
        raise Stop(f"indicators lacks the expected quote/adjclose arrays (assumption A2 fails): {e!r}")
    fields["adjclose"] = adj
    for k, v in fields.items():
        if len(v) != len(ts):
            raise Stop(f"Array '{k}' has {len(v)} items but there are {len(ts)} timestamps.")
    if len(ts) == 0:
        raise Stop("Zero bars returned.")
    return meta, ts, fields


def build_table(ticker, meta, ts, fields):
    """Converts each timestamp with the exchange time zone, per bar. Returns (DataFrame, report dict)."""
    tz = ZoneInfo(meta["exchangeTimezoneName"])
    rows, no_close = [], 0
    clock = {}
    all_dates = []                              # every bar in the reply, also the empty ones (used only to judge gaps)
    for i, t in enumerate(ts):
        local = dt.datetime.fromtimestamp(t, tz=tz)
        hhmm = local.strftime("%H:%M")
        all_dates.append(local.date())
        clock[hhmm] = clock.get(hhmm, 0) + 1
        if fields["close"][i] is None:
            no_close += 1                       # counted, reported, not hidden
            continue
        rows.append((local.date().isoformat(), hhmm, fields["open"][i], fields["high"][i], fields["low"][i],
                     fields["close"][i], fields["adjclose"][i], fields["volume"][i]))
    df = pd.DataFrame(rows, columns=["date", "_hhmm", "open", "high", "low", "close", "adj_close", "volume"])
    info = {"all_dates": all_dates, "bars_in_reply": len(ts), "bars_without_close": no_close, "clock_times": clock,
            "gmtoffset_in_meta": meta["gmtoffset"], "tz": meta["exchangeTimezoneName"], "currency": meta["currency"]}
    return df, info


def off_time_check(df, off, all_dates):
    """Accepts bars not stamped at midnight ONLY if they pass every test below; otherwise raises Stop. Returns one log row per bar."""
    last_date = df["date"].iloc[-1]
    out = []
    pos = {d: i for i, d in enumerate(df["date"])}
    for _, r in off.iterrows():
        d, hhmm = r["date"], r["_hhmm"]
        i = pos[d]
        day = dt.date.fromisoformat(d)
        prev_d = dt.date.fromisoformat(df["date"].iloc[i - 1]) if i > 0 else None      # previous bar WITH a close
        next_d = dt.date.fromisoformat(df["date"].iloc[i + 1]) if i + 1 < len(df) else None
        # Gaps are judged on every bar of the reply, also bars Yahoo returns without prices (empty placeholders).
        # A long gap between bars with a close is accepted only when the reply itself has a bar within the limit on that side;
        # a stretch where the reply has no bar at all (missing days) still stops the run.
        k = bisect.bisect_left(all_dates, day)
        prev_any = all_dates[k - 1] if k > 0 else None
        k2 = bisect.bisect_right(all_dates, day)
        next_any = all_dates[k2] if k2 < len(all_dates) else None
        problems = []
        if d == last_date and next_d is None:
            pass                                    # the newest, possibly still-open bar: allowed as before, but listed
        else:
            if int(hhmm[:2]) >= OFF_TIME_LATEST_HOUR:
                problems.append(f"stamped at {hhmm}, not in the morning (a late stamp could belong to the next trading day)")
            if day.weekday() >= 5:
                problems.append("falls on a weekend")
            if prev_any is not None and not (1 <= (day - prev_any).days <= OFF_TIME_MAX_GAP_DAYS):
                problems.append(f"previous bar in the reply is {(day - prev_any).days} days earlier")
            if next_any is not None and not (1 <= (next_any - day).days <= OFF_TIME_MAX_GAP_DAYS):
                problems.append(f"next bar in the reply is {(next_any - day).days} days later")
        if problems:
            raise Stop(f"Off-time bar {d} {hhmm} is NOT accepted: " + "; ".join(problems) + ". Nothing saved. Show this to the Coder.")
        gap_p = (day - prev_d).days if prev_d else ""
        gap_n = (next_d - day).days if next_d else ""
        long_gap = any(isinstance(g, int) and g > OFF_TIME_MAX_GAP_DAYS for g in (gap_p, gap_n))
        out.append({"date": d, "local_time": hhmm, "previous_bar_date": prev_d.isoformat() if prev_d else "",
                    "next_bar_date": next_d.isoformat() if next_d else "", "days_to_previous_bar_with_close": gap_p,
                    "days_to_next_bar_with_close": gap_n, "close": r["close"],
                    "result": "accepted: long gap between bars with a close is made of empty bars in Yahoo's own reply" if long_gap else "accepted"})
    # duplicates are caught by G2 below; the order is checked there too
    return out


def guards(ticker, df, info):
    """Every check that can stop the save. Returns a list of notes to print."""
    notes = []
    # G1: the clock time of the bars. If most bars are not at the expected local time, the date rule is wrong.
    share = df["_hhmm"].value_counts(normalize=True)
    top, top_share = share.index[0], float(share.iloc[0])
    if top != EXPECTED_LOCAL_TIME or top_share < 0.99:
        raise Stop(f"Bar clock times in {info['tz']}: {share.round(4).to_dict()}. Expected {EXPECTED_LOCAL_TIME} for >=99% of bars. "
                   "The 'local midnight' assumption (A3) is not confirmed; do NOT save. Show this to the Coder.")
    off = df[df["_hhmm"] != EXPECTED_LOCAL_TIME]
    info["off_time_log"] = []
    if len(off):
        info["off_time_log"] = off_time_check(df, off, info["all_dates"])
        shown = ", ".join(f"{r['date']} {r['local_time']}" for r in info["off_time_log"])
        notes.append(f"{len(off)} bar(s) not stamped at {EXPECTED_LOCAL_TIME} were checked and accepted (each: morning stamp on a weekday, exchange-zone date equals the stamp's own date, "
                     f"previous and next bars in the reply are 1 to {OFF_TIME_MAX_GAP_DAYS} days away, no duplicate date): {shown}. Listed in the off-time log file."
                     + (f" {sum(1 for x in info['off_time_log'] if x['result'] != 'accepted')} of them sit next to a long gap of bars with a close, made of empty bars in Yahoo's reply (marked in the log)." if any(x['result'] != 'accepted' for x in info['off_time_log']) else ""))
    # G2: sorted, unique
    d = pd.to_datetime(df["date"])
    if not d.is_monotonic_increasing:
        raise Stop("Dates are not in ascending order.")
    dups = df["date"][df["date"].duplicated(keep=False)]
    if len(dups):
        raise Stop(f"Duplicate dates after conversion: {sorted(set(dups))[:10]}. Nothing is dropped automatically.")
    # G3: weekends
    wk = d.dt.dayofweek
    if (wk >= 5).any():
        raise Stop(f"{int((wk >= 5).sum())} weekend-dated bars after conversion, e.g. {df['date'][wk >= 5].head(3).tolist()}. Not saved.")
    # G4: prices
    if (df[["open", "high", "low", "close"]].isna().any()).any():
        raise Stop("Missing open/high/low in a bar that has a close. Nothing is filled.")
    neg = df[df["close"] <= 0]
    for _, r in neg.iterrows():
        if (ticker, r["date"]) not in NEGATIVE_CLOSE_ALLOWED:
            raise Stop(f"Close <= 0 on {r['date']} for {ticker}: {r['close']}")
    if len(neg):
        notes.append(f"Negative close kept as published: {neg[['date','close']].values.tolist()}")
    if info["currency"] != "USD":
        raise Stop(f"Currency in meta is {info['currency']!r}, expected USD.")
    # G5: informational flags (never drop)
    flat = ((df.open == df.high) & (df.high == df.low) & (df.low == df.close) & (df.volume == 0)).sum()
    notes.append(f"Stale-looking bars (open=high=low=close and volume 0): {int(flat)} (kept, flagged, not removed)")
    ne = int((abs(df.adj_close - df.close) > 1e-9).sum())
    notes.append(f"Bars where adj_close differs from close: {ne} (futures should normally be 0; not verified)")
    # G6: Friday against Monday in US summer months (the shift made Fridays vanish)
    summer = d.dt.month.isin(range(3, 11))
    fri, mon = int(((wk == 4) & summer).sum()), int(((wk == 0) & summer).sum())
    if mon and not (0.85 <= fri / mon <= 1.15):
        raise Stop(f"Friday/Monday rows in March-October: {fri}/{mon}. Expected within 15% of each other.")
    notes.append(f"Friday rows {fri} against Monday rows {mon} in March to October")
    return notes


def compare_collector(ticker, df):
    """Agreement with the collector's yfinance copy (only 4 tickers have one). Returns text lines."""
    pat = COLLECTOR_FILES.get(ticker)
    if not pat:
        return [f"No collector copy for {ticker}: agreement check not possible"]
    files = sorted(COLLECTOR_DIR.glob(pat))
    if not files:
        return [f"Collector copy not found at {COLLECTOR_DIR} ({pat}): check skipped"]
    c = pd.read_csv(files[-1])
    c["date"] = c["Date"].astype(str).str[:10]
    m = df.merge(c[["date", "Close"]], on="date", how="inner")
    agree = int(((m["close"] - m["Close"]).abs() < 1e-3).sum())
    lines = [f"Collector copy {files[-1].name}: {len(c)} rows, overlap {len(m)} dates, "
             f"close agrees on {agree} ({agree / max(len(m), 1):.2%})"]
    only_c = len(set(c["date"]) - set(df["date"]))
    lines.append(f"Dates only in the collector copy: {only_c}; dates only in the new file inside the overlap window: "
                 f"{len(set(df['date'][(df['date'] >= c['date'].min()) & (df['date'] <= c['date'].max())]) - set(c['date']))}")
    if agree / max(len(m), 1) < 0.99:
        lines.append("WARNING: less than 99% agreement. Do not accept this file before the Coder has looked at it.")
    return lines


def compare_old_file(ticker, df):
    """Sanity against the OLD file of the same ticker (read only): size of history and price level (units, wrong contract)."""
    p = OLD_DIR / f"{ticker}.csv"
    if not p.exists():
        return [f"Old file {p} not found: size and level check skipped"]
    o = pd.read_csv(p)
    if len(df) < 0.98 * len(o):
        raise Stop(f"New file has {len(df)} rows, old file {len(o)}: more than 2% fewer. Possible partial download. Not saved.")
    gap = abs((pd.Timestamp(df["date"].iloc[0]) - pd.Timestamp(o["date"].iloc[0])).days)
    if gap > 5:
        raise Stop(f"First date {df['date'].iloc[0]} differs from the old file's {o['date'].iloc[0]} by {gap} days. Not saved.")
    w = df.tail(250)                                   # compare over the SAME calendar window, not the same row count
    lo = (pd.Timestamp(w["date"].iloc[0]) - pd.Timedelta(days=1)).strftime("%Y-%m-%d")
    ow = o[(o["date"] >= lo) & (o["date"] <= w["date"].iloc[-1])]
    if len(ow) < 100:
        raise Stop(f"Only {len(ow)} old-file rows fall in the comparison window {lo} to {w['date'].iloc[-1]}.")
    a, b = float(w["close"].median()), float(ow["close"].median())
    if not (0.95 <= a / b <= 1.05):
        raise Stop(f"Median close over {lo} to {w['date'].iloc[-1]}: {a:.3f} new against {b:.3f} in the old file: units or contract changed? Not saved.")
    return [f"Old file {p.name}: {len(o)} rows against {len(df)} new; first dates {o['date'].iloc[0]} / {df['date'].iloc[0]}; "
            f"median close over the same last-250-bars window {b:.3f} old, {a:.3f} new"]


def biggest_moves(df, n=5):
    """Lists the largest one-day changes for the eye (possible contract rolls or bad ticks). Nothing is removed."""
    ch = (df["close"].astype(float).diff() / df["close"].astype(float).shift(1).abs()).abs()
    top = ch.sort_values(ascending=False).head(n).index
    return ["Largest one-day moves (look for roll jumps or bad ticks; kept as published): " +
            "; ".join(f"{df['date'][i]} {df['close'][i - 1]:.4g}->{df['close'][i]:.4g}" for i in sorted(top) if i > 0)]


# ---------------------------------------------------------------- saving
def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        h.update(f.read())
    return h.hexdigest()


def save_new(path, data, mode):
    """Opens with 'x': fails if the file already exists. Never overwrites."""
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with open(path, mode) as f:
            f.write(data)
    except FileExistsError:
        raise Stop(f"{path} already exists. Raw files are never overwritten. Choose another day or stop.")


def run_ticker(ticker, getter=http_get, today=None):
    today = today or dt.date.today().strftime("%Y%m%d")
    raw_path = OUT_DIR / f"{ticker}__fetched_{today}.json"
    if raw_path.exists():                               # a run stopped after saving the raw reply: use that file, do not download or overwrite
        raw, headers = raw_path.read_bytes(), {"Date": "not downloaded again: the raw file saved earlier today is reused"}
    else:
        raw, headers = getter(url_for(ticker))
        save_new(raw_path, raw, "xb")                   # the untouched reply is the real raw file
    meta, ts, fields = parse_chart(raw)
    df, info = build_table(ticker, meta, ts, fields)
    notes = guards(ticker, df, info)
    notes += compare_old_file(ticker, df) + biggest_moves(df)
    csv_path = OUT_DIR / f"{ticker}__fetched_{today}.csv"
    out = df.drop(columns="_hhmm")
    save_new(csv_path, out.to_csv(index=False, lineterminator="\n"), "x")
    log_path = OUT_DIR / f"{ticker}__off_time_bars_{today}.csv"
    if info["off_time_log"]:
        lg = pd.DataFrame(info["off_time_log"])
        save_new(log_path, lg.to_csv(index=False, lineterminator="\n"), "x")
    now_local = dt.datetime.now(ZoneInfo(info["tz"])).date().isoformat()
    if out["date"].iloc[-1] >= now_local:
        notes.append(f"The newest bar ({out['date'].iloc[-1]}) is dated today in {info['tz']}: it may be an unfinished snapshot. "
                     "Treat the last bar as provisional (the collector copy of 28 Sep 2026 shows such a case).")
    report = {
        "ticker": ticker, "rows": len(out), "first": out["date"].iloc[0], "last": out["date"].iloc[-1],
        "no_close": info["bars_without_close"], "tz": info["tz"], "gmtoffset_meta": info["gmtoffset_in_meta"],
        "raw_file": str(raw_path), "raw_bytes": raw_path.stat().st_size, "raw_sha256": sha256_of(raw_path),
        "off_time_log_file": str(log_path) if info["off_time_log"] else "none (all bars at midnight)", "csv_file": str(csv_path), "csv_bytes": csv_path.stat().st_size, "csv_sha256": sha256_of(csv_path),
        "server_date_header": headers.get("Date", "not sent"),
        "notes": notes + compare_collector(ticker, out),
    }
    return report


def print_report(r):
    print(f"\n=== {r['ticker']} ===")
    print(f"rows {r['rows']}  first {r['first']}  last {r['last']}  bars without close (skipped, counted): {r['no_close']}")
    print(f"exchange time zone {r['tz']}; gmtoffset in meta {r['gmtoffset_meta']} (shown only, not used)")
    for n in r["notes"]:
        print("  -", n)
    print(f"raw  {r['raw_file']}  {r['raw_bytes']} bytes  sha256 {r['raw_sha256']}")
    print(f"off-time log  {r['off_time_log_file']}")
    print(f"csv  {r['csv_file']}  {r['csv_bytes']} bytes  sha256 {r['csv_sha256']}")
    print("server Date header:", r["server_date_header"], "(a stale value would hint at caching)")


# ---------------------------------------------------------------- diagnostic (saves nothing)
def diagnose(ticker, getter=http_get):
    raw, headers = getter(url_for(ticker))
    payload = json.loads(raw.decode("utf-8"))
    res = payload["chart"]["result"][0]
    print("top-level keys:", list(payload), "| chart keys:", list(payload["chart"]))
    print("result[0] keys:", list(res))
    print("meta keys:", sorted(res["meta"]))
    for k in ("symbol", "currency", "exchangeName", "instrumentType", "timezone", "exchangeTimezoneName",
              "gmtoffset", "dataGranularity", "range", "firstTradeDate"):
        print(f"  meta.{k} =", res["meta"].get(k, "<<NOT PRESENT>>"))
    print("indicators keys:", list(res["indicators"]), "| quote keys:", sorted(res["indicators"]["quote"][0]))
    print("number of timestamps:", len(res["timestamp"]), "| response header Date:", headers.get("Date", "not sent"))
    tzname = res["meta"].get("exchangeTimezoneName")
    tz = ZoneInfo(tzname) if tzname else ZoneInfo("America/New_York")
    ts = res["timestamp"]
    print("\nfirst three and last three bars (epoch seconds, UTC time, exchange-zone time):")
    for t in ts[:3] + ts[-3:]:
        print(" ", t, dt.datetime.fromtimestamp(t, tz=dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
              "|", dt.datetime.fromtimestamp(t, tz=tz).strftime("%Y-%m-%d %H:%M %Z (%a)"))
    print("\nOne winter and one summer sample (the heart of the check):")
    for want in ("2020-01-15", "2020-07-15", "2020-04-20"):
        hit = [t for t in ts if dt.datetime.fromtimestamp(t, tz=tz).strftime("%Y-%m-%d") in
               (want, (dt.date.fromisoformat(want) - dt.timedelta(days=1)).isoformat())][:3]
        for t in hit:
            u = dt.datetime.fromtimestamp(t, tz=dt.timezone.utc)
            print(f"  near {want}: epoch {t}  UTC {u:%Y-%m-%d %H:%M}  exchange-zone {dt.datetime.fromtimestamp(t, tz=tz):%Y-%m-%d %H:%M %Z %a}")
    hist = {}
    for t in ts:
        k = dt.datetime.fromtimestamp(t, tz=dt.timezone.utc).strftime("%H:%M")
        hist[k] = hist.get(k, 0) + 1
    print("\nUTC clock time of all bars (expect two values: 05:00 winter and 04:00 summer if bars are local midnight):", hist)
    loc = {}
    for t in ts:
        k = dt.datetime.fromtimestamp(t, tz=tz).strftime("%H:%M")
        loc[k] = loc.get(k, 0) + 1
    print("exchange-zone clock time of all bars (expect one value, 00:00):", loc)
    print("\nNothing was saved.")


def main(argv):
    if len(argv) >= 3 and argv[1] == "diagnose":
        diagnose(argv[2]); return
    if len(argv) == 2 and argv[1] == "fetch":
        reports = []
        today_ = dt.date.today().strftime("%Y%m%d")
        first = True
        for t in TICKERS:
            if (OUT_DIR / f"{t}__fetched_{today_}.csv").exists():
                print(f"\n=== {t} === already complete today (csv exists), skipped; nothing overwritten")
                continue
            if not (OUT_DIR / f"{t}__fetched_{today_}.json").exists():   # pause only before a real download
                if not first:
                    time.sleep(PAUSE_SECONDS)
                first = False
            r = run_ticker(t)
            print_report(r)
            reports.append(r)
        print("\nAll seven saved. Next: read the checklist, then draft the manifest lines.")
        return
    print(__doc__)


if __name__ == "__main__":
    try:
        main(sys.argv)
    except Stop as e:
        print("\nSTOPPED:", e)
        sys.exit(1)
