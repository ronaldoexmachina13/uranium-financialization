"""Recover missed CME UX daily bulletins from the Internet Archive (Wayback Machine).

CME keeps only the current bulletin on its site, but the Wayback Machine has saved
the same two files (the ones download_cme_uranium.py collects) many times. This
script lists those saved copies, downloads each one, reads the trade date and
bulletin number printed inside the PDF, and stores it under the collector's usual
file name. The date inside the PDF decides the name, not the date of the capture.

Rules (same as the collector):
  * raw PDFs are stored exactly as received and never overwritten;
  * a copy already on disk (same content, or same trade date + bulletin + status) is skipped;
  * every attempt is appended to logs/retrieval_log.csv;
  * private research use only: do not redistribute or push the PDFs anywhere (D-010).

It also recovers CME's two MONTHLY reports (NYMEX/COMEX Exchange Volume Report and
Open Interest Report). Each saved version prints the UX row for one month, the same
month a year earlier and the previous month, so every version adds up to three monthly
values. They go to data/raw/cme/monthly_reports/ and the UX numbers to
data/processed/monthly/uranium_futures_UX_monthly_volume_OI__CME_reports.csv.

Usage (from the Project folder, same Python as the collector):
    python src/recover_cme_from_wayback.py --dry-run        # list what would be fetched
    python src/recover_cme_from_wayback.py                  # daily from 2026-03-01 + all monthly reports
    python src/recover_cme_from_wayback.py --from 20261006  # daily bulletins only from this date
    python src/recover_cme_from_wayback.py --monthly-only   # only the monthly reports
    python src/recover_cme_from_wayback.py --daily-only     # only the daily bulletins
The daily files are then parsed automatically (same as `download_cme_uranium.py --parse`).
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
import time

import pdfplumber
import requests

from common import PROCESSED, RAW, http_session, load_config, log_retrieval, save_raw, sha256
from download_cme_uranium import RAW_SUBDIR, header_info, parse_all

CDX = "https://web.archive.org/cdx/search/cdx"
PAUSE_SECONDS = 1.5  # be gentle with archive.org


def list_captures(session: requests.Session, url: str, start: str) -> list[str]:
    """Timestamps of successful (HTTP 200) captures of `url` from `start` (YYYYMMDD) on."""
    r = session.get(CDX, params={"url": url.replace("https://", "").replace("http://", ""),
                                 "output": "txt", "fl": "timestamp,statuscode",
                                 "from": start, "filter": "statuscode:200"},
                    timeout=60)
    r.raise_for_status()
    stamps = [line.split()[0] for line in r.text.splitlines() if line.strip()]
    return sorted(set(stamps))


def fetch(session: requests.Session, stamp: str, url: str) -> bytes:
    """Original bytes of one capture ('id_' asks Wayback for the file without its banner)."""
    last = None
    for attempt in range(4):
        try:
            r = session.get(f"https://web.archive.org/web/{stamp}id_/{url}", timeout=90)
            if r.status_code == 200 and r.content.startswith(b"%PDF"):
                return r.content
            last = f"HTTP {r.status_code}, {len(r.content)} bytes"
        except requests.RequestException as e:
            last = repr(e)
        time.sleep(3 * (attempt + 1))
    raise RuntimeError(last or "unknown error")


# ----------------------------------------------------------------- monthly reports
MONTHLY_SUBDIR = "cme/monthly_reports"
MONTHLY_REPORTS = {  # kind: report file name on cmegroup.com
    "volume": "Web_Volume_Report_NYMEX_COMEX.pdf",
    "open_interest": "Web_OI_Report_NYMEX_COMEX.pdf",
}
MONTHLY_BASE = "https://www.cmegroup.com/daily_bulletin/monthly_volume/"
MON = {m: i for i, m in enumerate(["JAN", "FEB", "MAR", "APR", "MAY", "JUN",
                                   "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"], 1)}


def parse_monthly_report(path, kind: str):
    """UX row of one monthly report -> (records, note). Records are
    ((year, month), metric, value, how_it_appears_in_the_report). Anything that does
    not match the known layout is returned as a note and NOT guessed."""
    months, row = None, None
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or "").splitlines():
                if months is None and "Commodity" in line:
                    found = re.findall(r"\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(\d{4})\b", line)
                    if len(found) >= 3:
                        months = [(int(y), MON[m]) for m, y in found]
                if row is None and re.match(r"^\s*URANIUM\b.*\bUX\b", line):
                    row = line
    if months is None or row is None:
        return [], "UX row or month header not found"
    nums = [float(t.replace(",", "")) for t in row.split() if re.fullmatch(r"-?[\d,]+(\.\d+)?", t)]
    cur, py, pm = months[0], months[1], months[2]
    if kind == "volume":
        if len(nums) != 5:
            return [], f"unexpected layout ({len(nums)} numbers): {row.strip()}"
        c, p, pmv, ytd, ytdp = nums
        return [(cur, "volume", c, "current month"), (py, "volume", p, "same month prior year"),
                (pm, "volume", pmv, "previous month"),
                ((cur[0], 0), "volume_ytd", ytd, "year to date"),
                ((cur[0] - 1, 0), "volume_ytd", ytdp, "year to date, prior year")], "ok"
    if len(nums) != 3:
        return [], f"unexpected layout ({len(nums)} numbers): {row.strip()}"
    c, p, pmv = nums
    return [(cur, "open_interest", c, "current month"), (py, "open_interest", p, "same month prior year"),
            (pm, "open_interest", pmv, "previous month")], "ok"


def recover_monthly(session: requests.Session, dry_run: bool) -> None:
    folder = RAW / MONTHLY_SUBDIR
    folder.mkdir(parents=True, exist_ok=True)
    stats = {"saved": 0, "already_have": 0, "failed": 0}
    for kind, fname in MONTHLY_REPORTS.items():
        url = MONTHLY_BASE + fname
        stamps = list_captures(session, url, "20000101")
        print(f"monthly {kind}: {len(stamps)} captures on archive.org")
        if dry_run:
            continue
        known = {sha256(p.read_bytes()) for p in folder.glob(f"{fname[:-4]}__*.pdf")}
        for i, stamp in enumerate(stamps, 1):
            wb_url = f"https://web.archive.org/web/{stamp}id_/{url}"
            try:
                data = fetch(session, stamp, url)
                if sha256(data) in known:
                    stats["already_have"] += 1
                    continue
                path, status = save_raw(data, MONTHLY_SUBDIR, f"{fname[:-4]}__wayback{stamp}", "pdf", dated=False)
                known.add(sha256(data))
                stats["saved"] += 1
                log_retrieval(f"UX_monthly_{kind}", "cme_monthly_report_wayback", wb_url, path, data,
                              status, f"capture {stamp}")
                print(f"  [{i}/{len(stamps)}] {kind} {stamp} -> {path.name}")
            except Exception as e:
                stats["failed"] += 1
                log_retrieval(f"UX_monthly_{kind}", "cme_monthly_report_wayback", wb_url, None, None,
                              "failed", f"capture {stamp}: {e!r}"[:300])
                print(f"  [{i}/{len(stamps)}] {kind} {stamp} FAILED: {e!r}", file=sys.stderr)
            time.sleep(PAUSE_SECONDS)
    if dry_run:
        return
    print(f"monthly reports: new {stats['saved']}, already on disk {stats['already_have']}, failed {stats['failed']}")
    parse_monthly_all()


def parse_monthly_all() -> None:
    folder = RAW / MONTHLY_SUBDIR
    rows, unread = [], []
    for kind, fname in MONTHLY_REPORTS.items():
        for p in sorted(folder.glob(f"{fname[:-4]}__*.pdf")):
            try:
                recs, note = parse_monthly_report(p, kind)
            except Exception as e:
                recs, note = [], repr(e)
            if note != "ok":
                unread.append((p.name, note))
            for (y, m), metric, value, how in recs:
                rows.append({"month": f"{y}-{m:02d}" if m else f"{y}-YTD", "metric": metric, "value": value,
                             "how_in_report": how, "report_file": p.name})
    dest = PROCESSED / "monthly"
    dest.mkdir(parents=True, exist_ok=True)
    out = dest / "uranium_futures_UX_monthly_volume_OI__CME_reports.csv"
    rows.sort(key=lambda r: (r["metric"], r["month"], r["report_file"]))
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["month", "metric", "value", "how_in_report", "report_file"])
        w.writeheader()
        w.writerows(rows)
    months = sorted({r["month"] for r in rows if r["metric"] != "volume_ytd"})
    print(f"monthly UX values: {len(rows)} rows, {len(months)} distinct months"
          + (f" ({months[0]} to {months[-1]})" if months else "") + f" -> {out}")
    # Same month + metric printed in several reports: show if they disagree (nothing is altered)
    seen: dict = {}
    for r in rows:
        seen.setdefault((r["month"], r["metric"]), set()).add(r["value"])
    clash = [k for k, v in seen.items() if len(v) > 1]
    if clash:
        print(f"note: {len(clash)} month/metric pairs have different values in different reports "
              f"(later reports can revise earlier months); all are kept in the CSV, e.g. {clash[:3]}")
    for name, note in unread:
        print(f"  not read: {name}: {note}", file=sys.stderr)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="start", default="20260301", help="first capture date, YYYYMMDD")
    ap.add_argument("--dry-run", action="store_true", help="list captures only, download nothing")
    ap.add_argument("--daily-only", action="store_true", help="skip the monthly reports")
    ap.add_argument("--monthly-only", action="store_true", help="skip the daily bulletins")
    a = ap.parse_args()

    cfg = load_config()
    bc = cfg["cme_bulletin"]
    session = http_session(cfg)
    folder = RAW / RAW_SUBDIR
    folder.mkdir(parents=True, exist_ok=True)
    stats = {"saved": 0, "already_have": 0, "failed": 0}

    if not a.daily_only:
        recover_monthly(session, a.dry_run)
    if a.monthly_only:
        return

    for sec, fname in bc["sections"].items():
        url = bc["base_url"] + fname
        stamps = list_captures(session, url, a.start)
        print(f"{sec}: {len(stamps)} captures on archive.org since {a.start}")
        if a.dry_run:
            continue
        known = {sha256(p.read_bytes()) for p in folder.glob(f"*{sec}.pdf")}
        for i, stamp in enumerate(stamps, 1):
            wb_url = f"https://web.archive.org/web/{stamp}id_/{url}"
            try:
                data = fetch(session, stamp, url)
                digest = sha256(data)
                if digest in known:
                    stats["already_have"] += 1
                    continue
                tmp = folder / f"_incoming_{sec}.pdf"
                tmp.write_bytes(data)
                try:
                    with pdfplumber.open(tmp) as pdf:
                        info = header_info(pdf)
                finally:
                    tmp.unlink(missing_ok=True)
                stem = (f"CME_bulletin_UX__{info['trade_date']:%Y%m%d}__"
                        f"{info['bulletin_no']}_{info['status']}__{sec}")
                if (folder / f"{stem}.pdf").exists():  # same bulletin already stored: keep ours
                    known.add(digest)
                    stats["already_have"] += 1
                    log_retrieval(f"UX_{sec}", "cme_bulletin_wayback", wb_url, None, data,
                                  "unchanged", f"capture {stamp}: {stem} already on disk")
                    continue
                path, status = save_raw(data, RAW_SUBDIR, stem, "pdf", dated=False)
                known.add(digest)
                stats["saved"] += 1
                log_retrieval(f"UX_{sec}", "cme_bulletin_wayback", wb_url, path, data, status,
                              f"capture {stamp}: {info['trade_date']} #{info['bulletin_no']} {info['status']}")
                print(f"  [{i}/{len(stamps)}] {sec} {stamp} -> {path.name}")
            except Exception as e:  # keep going; one bad capture must not stop the run
                stats["failed"] += 1
                log_retrieval(f"UX_{sec}", "cme_bulletin_wayback", wb_url, None, None,
                              "failed", f"capture {stamp}: {e!r}"[:300])
                print(f"  [{i}/{len(stamps)}] {sec} {stamp} FAILED: {e!r}", file=sys.stderr)
            time.sleep(PAUSE_SECONDS)

    if a.dry_run:
        return
    print(f"\nnew files: {stats['saved']}, already on disk: {stats['already_have']}, failed: {stats['failed']}")
    parse_all()


if __name__ == "__main__":
    main()
