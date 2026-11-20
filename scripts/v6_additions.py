"""Workbook v6 (20 Nov 2026): what changes from v5, following Ronaldo's decisions of 20 Nov 2026.

Called by scripts/build_dataset.py. Copying, labelling, cutting by date and equality checks only: no return, ratio or
statistic is computed here (D-029). Monthly and quarterly sums of dated events are the same kind of sum the v5 build
already makes. Raw files are read, never written (D-005). The build stops (SystemExit) whenever a check fails.

What this file does
  1. Cut-off: every series ends at 31 Oct 2026. Ordinary series are cut by reference date. First prints (ALFRED and the
     Cameco contracting figures) are cut by the date they were published. The Annual forward-looking columns
     (2026 to 2035 values) are kept because they were published before 31 Oct 2026.
  2. New columns: ALFRED first prints (Monthly and Quarterly), Cameco long-term contracting first prints (Annual, Quarterly),
     SPUT issuance in month for 2026 (Monthly; light orange where the source is the unofficial tracker).
  3. New sheets: SPUT units daily, UPC holdings changes, Events, Read me.
  4. Labels: the Data Curator's label table of 20 Nov 2026 is copied onto the Variables sheet, one row per data column.
  5. 20 Nov 2026 (answers): Monthly SPUT units issued for Nov 2021 to Feb 2022 (official dated issues plus tracker sums, light orange),
     the check that Quarterly SPUT purchases equal the Monthly sums, and the label checks after the Yellow Cake ratio column was removed.
"""
import csv
import datetime as dt
import re
from collections import defaultdict
from pathlib import Path

from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
CUTOFF = dt.date(2026, 10, 31)
ORANGE = "FFD8A8"  # light orange: value from an unofficial source (the SPUT Tracker sheet), checked against Sprott where possible
NUM, INT = "#,##0.00", "#,##0"

F_EVENTS = RAW / "events" / "event_timeline_2005_2026_2026-11-20.csv"
F_TERM = RAW / "physical" / "term_contracting_volumes_quoted_2002_2026_2026-11-20.csv"
F_SPOTQ = RAW / "physical" / "spot_volume_quotes_not_in_existing_csv_2026-11-20.csv"
F_UPC_CHG = RAW / "upc" / "upc_holdings_changes_2010-12_to_2021-05_2026-11-20.csv"
F_TRACK21 = RAW / "sprott" / "sput_units_outstanding_daily_thirdparty_2021-09-27_to_2022-03-02_2026-11-20.csv"
F_UNITS26 = RAW / "sprott" / "sput_units_2026_official_and_tracker_2026-11-20.csv"
F_LABELS = RAW / "labels" / "label_table_v5_and_v6_new_2026-11-20.csv"  # labels, not data: the Data Curator's table of 20 Nov 2026
PHYSICAL_NOT_IN_LONG_TABLE = {F_TERM.name, F_SPOTQ.name}


def read_csv(p):
    with open(p, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def day(s):
    return dt.date.fromisoformat(str(s)[:10])


def num(v):
    v = (v or "").strip()
    if v in ("", "."):
        return None
    return float(v)


def is_number(v):
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


def qend(d):
    m = ((d.month - 1) // 3) * 3 + 3
    return dt.date(d.year + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)


# ---------------------------------------------------------------- 1. cut-off
def key_date(k):
    """Reference date of a row key: (year, month) -> last day of the month; a date stays; an int year is not cut here."""
    if isinstance(k, tuple):
        return dt.date(k[0] + (k[1] == 12), k[1] % 12 + 1, 1) - dt.timedelta(days=1)
    if isinstance(k, dt.datetime):
        return k.date()
    return k


def cut_columns(cols):
    """Drop every value whose reference date is after 31 Oct 2026. Returns the number of values dropped."""
    n = 0
    for c in cols:
        keep = {k: v for k, v in c["values"].items() if isinstance(k, int) or key_date(k) <= CUTOFF}
        n += len(c["values"]) - len(keep)
        c["values"] = keep
    return n


def check_annual_publication(ann):
    """Annual columns are not cut by reference year: the forward-looking ones (2026-2035) were published before the cut-off.
    The publication of each source file is stated here so the rule can be checked; the build stops if a new Annual source appears."""
    published = {  # raw file -> publication of the document the values come from (all before 31 Oct 2026)
        "data/raw/eia/eia_table3_us_uranium_production.csv": "EIA Domestic Uranium Production Report 2025 (Jun 2026)",
        "data/raw/eia/eia_table22_inventories_by_type.csv": "EIA Uranium Marketing Annual Report 2025 (2026, before the cut-off; Curator 20 Nov 2026: UMAR Aug 2026)",
        "data/raw/eia/eia_table23_inventories_by_owner.csv": "EIA Uranium Marketing Annual Report 2025",
        "data/raw/eia/eia_table11_unfilled_requirements.csv": "EIA Uranium Marketing Annual Reports 2024 and 2025",
        "data/raw/eia/eia_table12_max_requirements_end2025.csv": "EIA Uranium Marketing Annual Report 2025",
        "data/raw/eia/eia_table10_contracted_deliveries_in_effect_end2025.csv": "EIA Uranium Marketing Annual Report 2025",
        "data/raw/iaea_pris/iaea_pris_world_annual_1970_2025.csv": "IAEA PRIS, read 8 Nov 2026 (data to 2025)",
        "data/raw/iaea_pris/iaea_pris_world_performance_1970_2025.csv": "IAEA PRIS, read 9 Nov 2026 (data to 2025)",
        "data/raw/iaea_pris/iaea_rds2_2026_world_reactor_events_1954_2025.csv": "IAEA RDS-2, 46th edition (2026, data to end 2025)",
        str(F_TERM.relative_to(ROOT)).replace("\\", "/"): "Cameco filings; only first prints published by 31 Oct 2026 are used",
    }
    for c in ann:
        if c["raw"] not in published:
            raise SystemExit(f"Annual column {c['header']}: its source {c['raw']} has no stated publication date in v6_additions.check_annual_publication")
    return published


# ---------------------------------------------------------------- 2a. ALFRED first prints
ALFRED = [  # (series id, sheet, header of the current series it sits beside, new header, unit, note)
    ("CPIAUCSL", "M", "Consumer price index (CPI-U, seasonally adjusted)", "Consumer price index (CPI-U, seasonally adjusted), first print (ALFRED)", "index 1982-84=100",
     "Value as first published (ALFRED initial release), not revised. Oct 2025 was never published (BLS collected no CPI data in Oct 2025, government shutdown): left empty."),
    ("INDPRO", "M", "US industrial production", "US industrial production, first print (ALFRED)", "index, base of its own vintage",
     "Value as first published, on the index base of its own vintage, so the level jumps at each rebasing."),
    ("OECDPRINTO01GYSAM", "M", "OECD industrial production growth, year on year", "OECD industrial production growth, year on year, first print (ALFRED)", "percent",
     "Value as first published; the ALFRED archive for this series starts with the 2018 vintages."),
    ("OECDPRINTO01IXOBSAM", "M", "Industrial production, OECD total", "Industrial production, OECD total, first print (ALFRED), not for SPUT-era models", "index",
     "Value as first published. Ends 2024-03 like the revised series, so it covers only part of the SPUT era: not for SPUT-era models."),
    ("G7PRINTO01IXOBSAM", "M", "Industrial production, G7", "Industrial production, G7, earliest ALFRED vintage, DO NOT USE", "index",
     "DO NOT USE as first prints: 69 of 71 rows carry the archive start date 2024-04-10, so before Apr 2024 they are the 2024 vintage, not true first prints. Ends 2024-03."),
    ("GDPC1", "Q", "Real GDP, US", "Real GDP, US, first print (ALFRED)", "billion chained USD of its own vintage, SAAR",
     "Value as first published (advance estimate), in chained dollars of its own vintage, so the level jumps at each rebasing. ALFRED dates the quarter by its first day; here it sits on the quarter-end row."),
    ("A191RL1Q225SBEA", "Q", "Real GDP growth, US, annualised", "Real GDP growth, US, annualised, first print (ALFRED)", "percent",
     "Value as first published (advance estimate). ALFRED dates the quarter by its first day; here it sits on the quarter-end row."),
]
ALFRED_FILE = "data/raw/alfred/{sid}_initial_release_fetched_20261120.csv"


def _insert_after(cols, base_start, new_cols):
    idx = [i for i, c in enumerate(cols) if c["header"].split(" [")[0] == base_start]
    if len(idx) != 1:
        raise SystemExit(f"v6: current series '{base_start}' found {len(idx)} times; cannot place its first print beside it")
    cols[idx[0] + 1:idx[0] + 1] = new_cols


def add_alfred(M, Q, col):
    """Seven ALFRED first-print columns, each with a 'date first published' column beside. Cut by publication date."""
    info = {}
    for sid, sh, base, hdr, unit, note in ALFRED:
        raw = ALFRED_FILE.format(sid=sid)
        rows = read_csv(ROOT / raw)
        vals, dates, after, dot = {}, {}, [], []
        for r in rows:
            ref, pub = day(r["date"]), day(r["first_published_on"])
            if pub > CUTOFF:
                after.append(r["date"])
                continue
            k = (ref.year, ref.month) if sh == "M" else qend(ref)
            if k in vals:
                raise SystemExit(f"ALFRED {sid}: period {k} twice")
            v = num(r["value_first_published"])
            if v is None:
                dot.append(r["date"])
                continue
            vals[k], dates[k] = v, pub
        full_note = (f"{note} Rows first published after 31 Oct 2026 are left out (publication-date rule of 20 Nov 2026): {', '.join(after) or 'none'}. "
                     f"The next column gives the date each value was first published."
                     + (" The Monthly sheet starts in Jan 1988, so earlier first prints stay in the raw file." if sh == "M" and min(vals) < (1988, 1) else ""))
        c1 = col(f"{hdr} [{unit}]", vals, "native", "ALFRED (St. Louis Fed archive of FRED vintages), initial release", sid + " first print", "monthly" if sh == "M" else "quarterly",
                 "first print", unit, full_note, "", raw, NUM)
        c2 = col(f"{hdr}: date first published [date]", dates, "native", "ALFRED (realtime_start of the initial release)", sid + " first_published_on", "monthly" if sh == "M" else "quarterly",
                 "date first published", "date", "Date the value in the previous column was first published (ALFRED realtime_start). A dating field, not a predictor.", "", raw, "yyyy-mm-dd")
        _insert_after(M if sh == "M" else Q, base, [c1, c2])
        info[sid] = dict(n=len(vals), first=min(vals), last=max(vals), after=after, dot=dot, header=c1["header"].split(" [")[0])
    return info


# ---------------------------------------------------------------- 2b. long-term contracting, Cameco first prints
TERM_RAW = str(F_TERM.relative_to(ROOT)).replace("\\", "/")
_FORECAST = re.compile(r"expected|forecast|estimate\)", re.I)


def _term_rows():
    out = []
    for r in read_csv(F_TERM):
        if not re.search(r"1009001|cameco", r["source_file_or_url"], re.I):
            continue  # Yellow Cake quotes stay in the raw file
        if not r["first_print_or_restated"].startswith("first print"):
            continue  # later prints stay in the raw file
        out.append(r)
    return out


def term_contracting():
    """Annual: Cameco full-year first prints from 2002. Quarterly: Cameco 'long-term contracting reported by UxC', year to date, from Q3 2016.
    Forecasts, ranges and values such as 'more than 240' stay out. Year to date is NOT turned into single quarters (that is a calculation for Ronaldo)."""
    rows = _term_rows()
    ann, ann_d, q, q_d, left = {}, {}, {}, {}, []
    for r in rows:
        pe, rel = day(r["period_end"]), r["document_release_date"]
        if _FORECAST.search(r["basis"]):
            left.append((r["period_end"], r["volume_mlb"], "forecast, not an outcome"))
            continue
        if not is_number(r["volume_mlb"]):
            left.append((r["period_end"], r["volume_mlb"], "not a number"))
            continue
        if day(rel) > CUTOFF:
            left.append((r["period_end"], r["volume_mlb"], "published after 31 Oct 2026"))
            continue
        if r["period_type"] == "year" and "western world" not in r["basis"]:
            if pe.year in ann:
                raise SystemExit(f"term contracting: two Cameco first prints for {pe.year}")
            ann[pe.year], ann_d[pe.year] = float(r["volume_mlb"]), day(rel)
        elif r["period_type"] == "year-to-date" and r["provider"] == "UxC" and pe >= dt.date(2016, 9, 30):
            if pe != qend(pe):
                raise SystemExit(f"term contracting: year-to-date period end {pe} is not a quarter end")
            if pe in q:
                raise SystemExit(f"term contracting: two Cameco first prints for {pe}")
            q[pe], q_d[pe] = float(r["volume_mlb"]), day(rel)
        elif r["period_type"] in ("year", "year-to-date"):
            left.append((r["period_end"], r["volume_mlb"], "outside the two series (provider or start date)"))
    if min(ann) != 2002 or min(q) != dt.date(2016, 9, 30):
        raise SystemExit("term contracting: the series do not start where the decision says (2002 and Q3 2016)")
    return ann, ann_d, q, q_d, left


def add_term_quarterly(Q, col):
    ann, ann_d, q, q_d, left = term_contracting()
    hdr = "Long-term contracting volume, year to date, UxC as quoted by Cameco, first print"
    note = ("Cameco quarterly MD&A, 'long-term contracting reported by UxC', cumulative from 1 January to the quarter end (year to date), as first printed. "
            "Q1 to Q3 only: the full year is on the Annual sheet. NOT converted into single quarters (that would be a calculation; Ronaldo's). Later prints and "
            "Yellow Cake quotes stay in the raw file. Market Expert, 20 Nov 2026 (grade A, SEC EDGAR filings and kept files).")
    Q.append(col(f"{hdr} [million lb U3O8]", q, "native", "Cameco quarterly MD&A (UxC figures as quoted)", "term contracting YTD", "quarterly", "year to date", "million lb U3O8", note, "", TERM_RAW, NUM))
    Q.append(col(f"{hdr}: document release date [date]", q_d, "native", "SEC EDGAR filing date of the Cameco report", "term contracting YTD release", "quarterly", "release date", "date",
                 "Date the figure in the previous column was first published (filing date). A dating field, not a predictor.", "", TERM_RAW, "yyyy-mm-dd"))
    return dict(q=len(q), q_first=min(q), q_last=max(q), left=left)


def term_annual_cols(col):
    ann, ann_d, q, q_d, left = term_contracting()
    hdr = "Long-term contracting volume, year, Cameco first print"
    note = ("World long-term (term) contracting for the calendar year as first printed by Cameco: its own estimates to 2013, figures attributed to UxC from 2014. "
            "Forecasts and values printed as ranges or 'more than' (2005, 2006) are left out. No first print found for 2009, 2010, 2015 to 2018. "
            "Later prints and Yellow Cake quotes stay in the raw file. Market Expert, 20 Nov 2026.")
    return [col(f"{hdr} [million lb U3O8]", ann, "native", "Cameco annual and quarterly reports (SEC EDGAR, kept files)", "term contracting year", "annual", "calendar year", "million lb U3O8", note, "", TERM_RAW, NUM),
            col(f"{hdr}: document release date [date]", ann_d, "native", "SEC EDGAR filing date of the Cameco report", "term contracting year release", "annual", "release date", "date",
                "Date the figure in the previous column was first published (filing date). A dating field, not a predictor.", "", TERM_RAW, "yyyy-mm-dd")], dict(n=len(ann), years=sorted(ann))


# ---------------------------------------------------------------- 2c. SPUT issuance in month, 2026
def sput_units_2026():
    rows = [r for r in read_csv(F_UNITS26)]
    return rows


def sput_issuance_2026(M):
    """Monthly 'SPUT units issued in month' for 2026: tracker monthly totals Jan to Jun (light orange), 0 for Jul to Oct from official documents.
    Gross value: 0 for Jul to Oct only (no issue); Jan to Jun stay blank (no issue prices published for 26 Jan to 12 May)."""
    rows = sput_units_2026()
    off = {r["date"]: r for r in rows if r["source"] == "Sprott official" and r["units_outstanding"]}
    if int(float(off["2026-06-30"]["units_outstanding"])) != int(float(off["2026-11-06"]["units_outstanding"])):
        raise SystemExit("SPUT 2026: official units at 30 Jun and 6 Nov 2026 differ, so Jul to Oct are not zero")
    tr = [r for r in rows if r["source"] == "unofficial tracker" and day(r["date"]) <= dt.date(2026, 6, 30) and day(r["date"]) >= dt.date(2026, 1, 1)]
    by = defaultdict(int)
    for r in tr:
        by[(2026, day(r["date"]).month)] += int(float(r["units_issued"] or 0))
    h1 = sum(by.values())
    h1_off = int(float(off["2026-06-30"]["units_issued"]))
    if h1 != h1_off:
        raise SystemExit(f"SPUT 2026: tracker issuance 1 Jan to 30 Jun ({h1}) differs from the interim report ({h1_off})")
    # the 12 official dated January days must equal the tracker on the same days
    offd = defaultdict(int)
    for r in rows:
        if r["source"] == "Sprott official" and r["issue_price_if_known"]:
            offd[r["date"]] += int(float(r["units_issued"]))
    trd = {r["date"]: int(float(r["units_issued"] or 0)) for r in tr if day(r["date"]) <= dt.date(2026, 1, 19)}
    if any(trd.get(d_, 0) != u for d_, u in offd.items()) or any(u and d_ not in offd for d_, u in trd.items()):
        raise SystemExit("SPUT 2026: tracker and Prior Sales table differ on 2 to 19 Jan 2026")
    S_ = next(c for c in M if c["header"].startswith("SPUT units issued in month"))
    T_ = next(c for c in M if c["header"].startswith("SPUT gross issuance value in month"))
    orange, comments = [], {}
    for m in range(1, 7):
        k = (2026, m)
        if k in S_["values"]:
            raise SystemExit(f"SPUT 2026: {k} already filled")
        S_["values"][k] = by.get(k, 0)
        orange.append(("SPUT units issued in month", k))
        comments[("SPUT units issued in month", k)] = (
            f"Unofficial source (light orange): sum of the SPUT Tracker's daily units issued in {dt.date(2026, m, 1):%b %Y} (grade C; data/raw/sprott/{F_UNITS26.name}). "
            + ("Includes 7,378,312 units on 2 to 19 Jan that the Prior Sales table also dates (grade A, equal day by day); " if m == 1 else "")
            + "The tracker's 2026 days add up exactly to the 33,085,763 units the 30 Jun 2026 interim report gives for the half-year. "
            + ("No issue after 12 May 2026 in the tracker; units at 30 Jun (interim report) equal the tracker. " if m == 6 else "")
            + "Confirmed by Ronaldo on 20 Nov 2026 (answer 1: keep Jan to May from the tracker and June's 0 in light orange, as built).")  # 20 Nov 2026 (text fix): was "Ronaldo to confirm (Coder report 20 Nov 2026)."
    for m in range(7, 11):
        k = (2026, m)
        S_["values"][k], T_["values"][k] = 0, 0.0
        comments[("SPUT units issued in month", k)] = comments[("SPUT gross issuance value in month", k)] = (
            "0 from official documents (grade A): 344,156,215 units outstanding at 30 Jun 2026 (interim report) and at 6 Nov 2026 (kept fund page); units cannot be redeemed. "
            "Data Curator 20 Nov 2026.")
    S_["note"] += (" | 20 Nov 2026: Jan to Jun 2026 are monthly sums of the unofficial SPUT Tracker (light orange, grade C; they add up to the official half-year total); "
                   "Jul to Oct 2026 are 0 from official documents (grade A).")  # 20 Nov 2026 (answers): "Nov 2021 to Feb 2022 stay blank" dropped; see sput_issuance_2021_22
    T_["note"] += " | 20 Nov 2026: Jul to Oct 2026 are 0 (no issue; grade A). Jan to Jun 2026 stay blank: issue prices for 26 Jan to 12 May 2026 are not published."
    return orange, comments, dict(by_month={f"2026-{m:02d}": by.get((2026, m), 0) for m in range(1, 7)}, h1=h1)


# ---------------------------------------------------------------- 2d. SPUT issuance in month, Nov 2021 to Feb 2022 (20 Nov 2026 (answers))
F_ISS = RAW / "sprott" / "sput_issuances.csv"
ARRANGEMENT_UNITS = 75210456  # units issued on acquiring UPC, 19 Jul 2021 (as in build_dataset.sput_fundraising)
WIN_A, WIN_B = dt.date(2021, 11, 19), dt.date(2022, 2, 15)  # the tracker window: no Prior Sales table lists these days


def _days(ds):
    """'3, 8, 15 Dec 2021' style list of the days (all in one month)."""
    ds = sorted(ds)
    return ", ".join(str(d.day) for d in ds) + f" {ds[0]:%b %Y}" if ds else "none"


def sput_issuance_2021_22(M, iss, stm):
    """20 Nov 2026 (answers), Ronaldo's answer 3. Monthly 'SPUT units issued in month' for Nov 2021, Dec 2021, Jan 2022 and Feb 2022:
    official dated issuances in the month (Prior Sales tables, data/raw/sprott/sput_issuances.csv, grade A) plus the sum of the unofficial
    tracker's daily units issued on the days of the month inside the window 19 Nov 2021 to 15 Feb 2022 (grade C). The whole cell is light orange
    in all four months. Gross issuance value stays blank (no issue prices for the tracker days). Only sums of dated rows and equality checks.
    A day present in both sources would be counted once, from the official table (no such day exists; the build reports the count)."""
    tr_all = read_csv(F_TRACK21)
    tr = [r for r in tr_all if r["in_undated_window"] == "yes"]
    if any(not (WIN_A <= day(r["date"]) <= WIN_B) for r in tr) or any(WIN_A <= day(r["date"]) <= WIN_B and r["in_undated_window"] != "yes" for r in tr_all):
        raise SystemExit("SPUT 2021-22: the tracker rows marked in_undated_window do not match the window 19 Nov 2021 to 15 Feb 2022")
    src_of = defaultdict(set)
    for r in read_csv(F_ISS):
        src_of[day(r["date"])].add(r["source"])
    off_days = {d for d, _p, _n, _t in iss}
    overlap = sorted(day(r["date"]) for r in tr if day(r["date"]) in off_days)
    tr = [r for r in tr if day(r["date"]) not in off_days]  # never count a day twice: the official table wins (none in practice)
    tu = lambda r: int(float(r["units_issued_shown"] or 0))
    # the tracker's units outstanding run without a break: previous day + units issued = next day
    for a, b in zip(tr, tr[1:]):
        if int(a["units_outstanding"]) + tu(b) != int(b["units_outstanding"]):
            raise SystemExit(f"SPUT 2021-22: tracker units outstanding do not roll forward from {a['date']} to {b['date']}")
    byend = {r["period_end"]: r for r in stm}
    U = lambda d: int(byend[d]["units_outstanding"])
    off_sum = lambda a, b: sum(n for d, _p, n, _t in iss if a <= d <= b)
    start_units = ARRANGEMENT_UNITS + off_sum(dt.date(2021, 1, 1), WIN_A - dt.timedelta(days=1))  # 162,558,000 after the 18 Nov 2021 table
    t21 = sum(tu(r) for r in tr if day(r["date"]) <= dt.date(2021, 12, 31))
    t22 = sum(tu(r) for r in tr if day(r["date"]) >= dt.date(2022, 1, 1))
    first, dec31, last = tr[0], next(r for r in tr if r["date"] == "2021-12-31"), tr[-1]
    checks = {
        "units on 19 Nov 2021 (tracker) less that day's issue = 75,210,456 + dated issues to 18 Nov 2021": int(first["units_outstanding"]) - tu(first) == start_units,
        "units on 31 Dec 2021 (tracker) = FY2021 annual report": int(dec31["units_outstanding"]) == U("2021-12-31"),
        "tracker 19 Nov to 31 Dec 2021 = FY2021 units (annual report) less 75,210,456 and the dated issues to 18 Nov 2021": t21 == U("2021-12-31") - start_units,
        "tracker 1 Jan to 15 Feb 2022 = H1 2022 units issued (interim report) less the dated issues 16 Feb to 30 Jun 2022": t22 == int(byend["2022-06-30"]["units_issued_in_period"]) - off_sum(dt.date(2022, 1, 1), dt.date(2022, 6, 30)),
        "units on 15 Feb 2022 (tracker) + dated issues 16 Feb to 30 Jun 2022 = units at 30 Jun 2022 (interim report)": int(last["units_outstanding"]) + off_sum(WIN_B + dt.timedelta(days=1), dt.date(2022, 6, 30)) == U("2022-06-30"),
    }
    bad = [k for k, ok in checks.items() if not ok]
    if bad:
        raise SystemExit("SPUT 2021-22: does not reconcile, the months are NOT filled: " + "; ".join(bad))
    S_ = next(c for c in M if c["header"].startswith("SPUT units issued in month"))
    T_ = next(c for c in M if c["header"].startswith("SPUT gross issuance value in month"))
    orange, comments, out = [], {}, {}
    for (y, m) in [(2021, 11), (2021, 12), (2022, 1), (2022, 2)]:
        k = (y, m)
        if k in S_["values"] or k in T_["values"]:
            raise SystemExit(f"SPUT 2021-22: {k} already filled")
        od = sorted({d for d, _p, _n, _t in iss if (d.year, d.month) == k})
        ou = sum(n for d, _p, n, _t in iss if (d.year, d.month) == k)
        tdays = [r for r in tr if (day(r["date"]).year, day(r["date"]).month) == k]
        tsum = sum(tu(r) for r in tdays)
        S_["values"][k] = ou + tsum
        orange.append(("SPUT units issued in month", k))
        span = f"{min(day(r['date']) for r in tdays):%d} to {max(day(r['date']) for r in tdays):%d %b %Y}"
        txt = (("MIXED, official and unofficial" if od else "UNOFFICIAL") + f" (whole cell light orange; 20 Nov 2026, Ronaldo's answer 3). {ou + tsum:,} units = "
               + (f"{ou:,} units on the official dated days {_days(od)} (Sprott Prior Sales table, {'; '.join(sorted(set().union(*(src_of[d] for d in od))))}; grade A; data/raw/sprott/sput_issuances.csv) + " if od else "")
               + f"{tsum:,} units on the tracker days {span} (unofficial SPUT Tracker, grade C; days with an issue: {_days([day(r['date']) for r in tdays if tu(r)])}; "
               + f"data/raw/sprott/{F_TRACK21.name}, rows marked in_undated_window = yes). No day is in both sources. "
               + "Checks: the tracker adds up exactly to the FY2021 units (annual report) and to the H1 2022 units issued (interim report) once the dated issues are taken off. "
               + "Gross value left blank: no issue prices for 19 Nov 2021 to 15 Feb 2022.")
        comments[("SPUT units issued in month", k)] = txt
        out[f"{y}-{m:02d}"] = dict(official=ou, official_days=len(od), tracker=tsum, tracker_rows=len(tdays), total=ou + tsum)
    old = "19 Nov 2021 to 15 Feb 2022 is not listed, so those months are blank"
    if old not in S_["note"]:
        raise SystemExit("SPUT 2021-22: the note of 'SPUT units issued in month' has changed; update the replacement text")
    S_["note"] = S_["note"].replace(old, "19 Nov 2021 to 15 Feb 2022 is not listed in any Prior Sales table")
    S_["note"] += (" | 20 Nov 2026 (answers): Nov 2021 to Feb 2022 are the official dated issues of the month plus the monthly sums of the unofficial SPUT Tracker's "
                   "daily units for 19 Nov 2021 to 15 Feb 2022 (whole cell light orange, grade C; Nov 2021 and Feb 2022 are mixed; the cell comment says which days come from which source). "
                   "The tracker adds up exactly to the FY2021 and H1 2022 report figures.")
    T_["note"] += " | 20 Nov 2026 (answers): Nov 2021 to Feb 2022 stay blank: no issue prices are published for 19 Nov 2021 to 15 Feb 2022."
    return orange, comments, dict(months=out, overlap_days=[d.isoformat() for d in overlap], checks=list(checks), t21=t21, t22=t22)


def check_quarterly_purchases(M, Q):
    """20 Nov 2026 (answers): Quarterly 'SPUT purchases in quarter' [million lb] must equal the sum of the Monthly 'SPUT pounds bought in month' [lb]
    for the same three months, in every quarter Q3 2021 to Q3 2026 (equality in whole pounds)."""
    mc = next(c for c in M if c["header"].startswith("SPUT pounds bought in month"))
    qc = next(c for c in Q if c["header"].startswith("SPUT purchases in quarter"))
    want = [dt.date(2021, 9, 30)]
    while want[-1] < dt.date(2026, 9, 30):
        want.append(qend(want[-1] + dt.timedelta(days=1)))
    if sorted(qc["values"]) != want:
        raise SystemExit(f"Quarterly purchases: quarters {sorted(qc['values'])[:1]}..{sorted(qc['values'])[-1:]} are not Q3 2021 to Q3 2026 without a gap")
    res = {}
    for qd in want:
        ms = [(qd.year, qd.month - 2), (qd.year, qd.month - 1), (qd.year, qd.month)]
        if any(k not in mc["values"] for k in ms):
            raise SystemExit(f"Quarterly purchases: the Monthly column has no value for a month of {qd}")
        msum = sum(mc["values"][k] for k in ms)
        if round(qc["values"][qd] * 1e6) != round(msum) or abs(msum - round(msum)) > 1e-9:
            raise SystemExit(f"Quarterly purchases {qd}: {qc['values'][qd]} million lb differs from the Monthly sum {msum} lb")
        res[qd.isoformat()] = int(round(msum))
    return dict(quarters=len(res), first=want[0].isoformat(), last=want[-1].isoformat(), lb=res)


def check_removed_columns(Q):
    """20 Nov 2026 (answers): the Yellow Cake ratio column must not come back; its inputs must still be there."""
    hs = [c["header"] for c in Q]
    if any(h.startswith("Yellow Cake premium/discount") for h in hs):
        raise SystemExit("Quarterly: the removed column 'Yellow Cake premium/discount to NAV' is back")
    if not any(h.startswith("Yellow Cake NAV per share") for h in hs):
        raise SystemExit("Quarterly: Yellow Cake NAV per share (input of the removed ratio) is missing")


def check_labels(rows, info):
    """20 Nov 2026 (answers): exactly one label row per data column, none for a removed column, no data column without a label."""
    seen = defaultdict(int)
    for r in rows:
        if r[1] in ("-", "all"):
            continue
        seen[(r[0], str(r[1]).split(" ")[0])] += 1
    dup = [k for k, n in seen.items() if n > 1]
    if dup:
        raise SystemExit(f"labels: more than one label row for {dup}")
    if info["unlabelled"]:
        raise SystemExit(f"labels: data columns without a label row: {info['unlabelled']}")
    gone = [r for r in rows if (r[0], r[2]) in REMOVED or (r[0], str(r[2]).split(" [")[0]) in REMOVED]
    if gone:
        raise SystemExit(f"labels: a removed column still has a label row: {gone}")


# ---------------------------------------------------------------- 3. new record sheets
def sput_units_daily_rows():
    """Daily SPUT units, 19 Nov 2021 to 15 Feb 2022 (tracker, grade C) and 31 Dec 2025 to 31 Oct 2026 (official rows and tracker rows)."""
    out = []
    for r in read_csv(F_TRACK21):
        if r["in_undated_window"] != "yes":
            continue
        out.append([day(r["date"]), int(r["units_outstanding"]), int(float(r["units_issued_shown"] or 0)), None, "unofficial tracker, grade C", "C",
                    r["source"], r["proof"], r["note"] or None, "data/raw/sprott/" + F_TRACK21.name])
    for r in read_csv(F_UNITS26):
        if day(r["date"]) > CUTOFF:
            continue
        src = "Sprott official" if r["source"] == "Sprott official" else "unofficial tracker, grade C"
        note = r["note"] or None
        if r["date"] == "2026-01-20" and r["source"] == "Sprott official":
            note = ("FLAG: printed 318,488,764 is kept; probably 318,448,764 (digits swapped), see Curator 2026-11-20_sput_2026.md. Not corrected: no correction file approved. "
                    + (note or ""))
        out.append([day(r["date"]), int(float(r["units_outstanding"])) if r["units_outstanding"] else None, int(float(r["units_issued"])) if r["units_issued"] else None,
                    float(r["issue_price_if_known"]) if r["issue_price_if_known"] else None, src, r["grade"], r["document_or_url"], r["page_or_row"], note, "data/raw/sprott/" + F_UNITS26.name])
    return out


SPUT_UNITS_HEADS = ["Date", "Units outstanding", "Units issued", "Issue price if known [USD]", "Source", "Grade", "Document or URL", "Page or row", "Note", "Raw file"]

UPC_TYPE = {"purchase": "purchase", "conversion exchange (UF6 out, U3O8 in)": "conversion exchange", "sale": "sale",
            "purchase and conversion exchange": "mixed", "purchase, conversion exchange and location swap": "mixed",
            "conversion exchange and sale": "mixed", "unexplained (+2 lb)": "unexplained"}


def upc_changes_rows():
    out = []
    for r in read_csv(F_UPC_CHG):
        t = UPC_TYPE.get(r["type"])
        if t is None:
            raise SystemExit(f"UPC holdings change {r['change_no']}: type '{r['type']}' has no label")
        fp = day(r["first_public_date"]) if re.match(r"\d{4}-\d\d-\d\d$", r["first_public_date"]) else r["first_public_date"]
        out.append([int(r["change_no"]), day(r["as_of_date"]), day(r["release_date"]), day(r["previous_as_of_date"]), int(r["u3o8_lb_before"]), int(r["u3o8_lb_after"]), int(r["u3o8_lb_change"]),
                    int(r["uf6_kgu_before"]), int(r["uf6_kgu_after"]), int(r["uf6_kgu_change"]), t, r["type"], r["what_the_release_says"], fp, r["first_public_document"],
                    r["first_public_proof"], r["release_document"], r["release_url"], r["grade_of_change"], r["grade_of_first_public_date"]])
    if len(out) != 28:
        raise SystemExit(f"UPC holdings changes: {len(out)} rows, expected 28")
    return out


UPC_CHG_HEADS = ["Change no.", "As of", "Release date", "Previous as of", "U3O8 held before [lb]", "U3O8 held after [lb]", "Change in U3O8 held [lb]",
                 "UF6 held before [KgU]", "UF6 held after [KgU]", "Change in UF6 held [KgU]", "Type", "Type as written by the Data Curator", "What the release says",
                 "First public date", "First public document", "First public proof", "Release document", "Release URL", "Grade of change", "Grade of first public date"]


def events_rows():
    rows = read_csv(F_EVENTS)
    out = []
    for r in rows:
        if r["event_date"][:7] > CUTOFF.isoformat()[:7]:
            raise SystemExit(f"Event {r['event_id']} is after the cut-off")
        full = re.match(r"\d{4}-\d\d-\d\d$", r["event_date"])  # month-precision events stay as text 'YYYY-MM'
        out.append([r["event_id"], day(r["event_date"]) if full else r["event_date"], r["date_precision"], r["announcement_or_effective"], r["event"], r["category"], r["region"], r["source_title"],
                    r["source_url_or_path"], r["page_or_section"], day(r["retrieved_on"]) if r["retrieved_on"] else None, r["grade"], r["note"] or None])
    return out


EVENTS_HEADS = ["Event ID", "Event date", "Date precision", "Announcement or effective", "Event", "Category", "Region", "Source title", "Source (URL or path)",
                "Page or section", "Retrieved on", "Grade", "Note"]


def fill_rows(ws, first_row, pred):
    """Light orange on every cell of the rows where pred(row values) is true."""
    n = 0
    for r in range(first_row + 1, ws.max_row + 1):
        vals = [ws.cell(r, j).value for j in range(1, ws.max_column + 1)]
        if pred(vals):
            for j in range(1, ws.max_column + 1):
                ws.cell(r, j).fill = PatternFill("solid", fgColor=ORANGE)
            n += 1
    return n


# ---------------------------------------------------------------- 4. register notes and column labels
def update_register(by, extras):
    """Add where the new v6 columns are to the register rows they belong to (Variables sheet)."""
    add = {
        "V068": "v6: Monthly also holds the first print (ALFRED) and the date it was first published.",
        "V081": "v6: Monthly also holds the first print (ALFRED) and the date it was first published.",
        "V082": "v6: Monthly also holds ALFRED first prints of OECD growth, OECD level (not for SPUT-era models) and G7 level (do not use: not true first prints before Apr 2024).",
        "V083": "v6: Quarterly also holds ALFRED first prints of real GDP level and growth, with the dates first published.",
        "V102": "v6: Annual: Cameco full-year first prints from 2002; Quarterly: UxC year to date as quoted by Cameco, first prints from Q3 2016 (raw file keeps every print).",
        "V107": "v6: Events sheet: dated event timeline 2005 to 2026 (93 rows; announcement and effective rows are separate).",
        "V021": "v6: SPUT units daily sheet: daily units outstanding 19 Nov 2021 to 15 Feb 2022 and 2026 (unofficial tracker rows in light orange).",
        "V022": "v6: SPUT units daily sheet (daily units issued, tracker rows in light orange); Monthly units issued in month filled for 2026 (Jan to Jun in light orange) and, from 20 Nov 2026 (answers), for Nov 2021 to Feb 2022 (official dated issues plus tracker sums, light orange).",  # 20 Nov 2026 (answers)
        "V027": "v6 (20 Nov 2026, answers): Quarterly SPUT purchases in quarter filled for every quarter Q3 2021 to Q3 2026 from the dated purchase list (equal to the sums of the Monthly column).",  # 20 Nov 2026 (answers)
        "D44": "v6: UPC holdings changes sheet: the 28 releases with a change in holdings, with a type label and the first public date.",
        "D47": "v6: SPUT units daily sheet fills the undated windows with daily units (unofficial tracker, grade C); Monthly units issued in month carries their monthly sums for Nov 2021 to Feb 2022 and Jan to Jun 2026 (light orange); issue prices for 19 Nov 2021 to 15 Feb 2022 and 26 Jan to 12 May 2026 are not published.",  # 20 Nov 2026 (answers): monthly sums added
    }
    for k, t in add.items():
        if k in by:
            st, wh, no = by[k]
            if k == "V102" and st not in ("In workbook",):
                st = "Partly in workbook"
            if k == "V107" and st not in ("In workbook",):
                st = "Partly in workbook"
            by[k] = (st, (wh + "; " if wh else "") + t, no)
    for i, x in enumerate(extras):
        if x[0] in add:
            id_, var, variant, st, wh, no = x
            extras[i] = (id_, var, variant, st, (wh + "; " if wh else "") + add[x[0]], no)


# label rows proposed by the Curator for new columns -> (v6 sheet, v6 header)
NEW_LABEL_MAP = {
    "CPI-U, first published value (ALFRED CPIAUCSL)": ("Monthly", "Consumer price index (CPI-U, seasonally adjusted), first print (ALFRED)"),
    "US industrial production, first published value (ALFRED INDPRO)": ("Monthly", "US industrial production, first print (ALFRED)"),
    "OECD industrial production growth, first published value (ALFRED OECDPRINTO01GYSAM)": ("Monthly", "OECD industrial production growth, year on year, first print (ALFRED)"),
    "OECD industrial production level, first published value (ALFRED OECDPRINTO01IXOBSAM)": ("Monthly", "Industrial production, OECD total, first print (ALFRED), not for SPUT-era models"),
    "G7 industrial production level, first published value (ALFRED G7PRINTO01IXOBSAM)": ("Monthly", "Industrial production, G7, earliest ALFRED vintage, DO NOT USE"),
    "US real GDP, first published value (ALFRED GDPC1)": ("Quarterly", "Real GDP, US, first print (ALFRED)"),
    "US real GDP growth, annualised, first published value (ALFRED A191RL1Q225SBEA)": ("Quarterly", "Real GDP growth, US, annualised, first print (ALFRED)"),
    "CPI-U, first published on (ALFRED CPIAUCSL)": ("Monthly", "Consumer price index (CPI-U, seasonally adjusted), first print (ALFRED): date first published"),
    "US industrial production, first published on (ALFRED INDPRO)": ("Monthly", "US industrial production, first print (ALFRED): date first published"),
    "OECD industrial production growth, first published on (ALFRED OECDPRINTO01GYSAM)": ("Monthly", "OECD industrial production growth, year on year, first print (ALFRED): date first published"),
    "OECD industrial production level, first published on (ALFRED OECDPRINTO01IXOBSAM)": ("Monthly", "Industrial production, OECD total, first print (ALFRED), not for SPUT-era models: date first published"),
    "G7 industrial production level, first published on (ALFRED G7PRINTO01IXOBSAM)": ("Monthly", "Industrial production, G7, earliest ALFRED vintage, DO NOT USE: date first published"),
    "US real GDP, first published on (ALFRED GDPC1)": ("Quarterly", "Real GDP, US, first print (ALFRED): date first published"),
    "US real GDP growth, first published on (ALFRED A191RL1Q225SBEA)": ("Quarterly", "Real GDP growth, US, annualised, first print (ALFRED): date first published"),
    "Long-term contracting volume, full year, Cameco first print": ("Annual", "Long-term contracting volume, year, Cameco first print"),
    "Long-term contracting volume, year to date, UxC as quoted by Cameco, first print": ("Quarterly", "Long-term contracting volume, year to date, UxC as quoted by Cameco, first print"),
    "Event timeline (93 events; columns event_id, event_date, date_precision, announcement_or_effective, event, category, region, source fields, grade, note)": ("Events", "Event timeline, all columns (93 rows)"),
    "Spot-volume quotes not in the existing Cameco spot-volume file": ("(raw file only, not in the workbook)", "data/raw/physical/" + F_SPOTQ.name),
    "Change in U3O8 held, lb (28 releases with a change)": ("UPC holdings changes", "Change in U3O8 held [lb]"),
    "Change in UF6 held, KgU": ("UPC holdings changes", "Change in UF6 held [KgU]"),
    "Type of change (purchase / conversion exchange / sale / mixed / unexplained)": ("UPC holdings changes", "Type"),
    "First public date of the transaction": ("UPC holdings changes", "First public date"),
    "Release date of the monthly NAV that shows the change": ("UPC holdings changes", "Release date"),
    "SPUT units outstanding, daily, unofficial tracker (grade C)": ("SPUT units daily", "Units outstanding"),
    "SPUT units issued, daily, unofficial tracker (grade C)": ("SPUT units daily", "Units issued"),
    'Source column ("unofficial tracker, grade C")': ("SPUT units daily", "Source"),
}
# labels the Curator's table has no row for: the dating columns beside the contracting first prints and the descriptive columns of the two new
# record sheets. Same rules as the Curator's own rows (dating and descriptive fields: do not use; values copied from another sheet: copy of).
# (sheet, header, variable ID, model use, reason, timing, spot link)
_D, _DESC = "not applicable (descriptive field)", ("do not use", "descriptive field", "INFO")
CODER_LABELS = [
    ("Annual", "Long-term contracting volume, year, Cameco first print: document release date", "release date", "do not use", "dating field", "INFO", _D),
    ("Quarterly", "Long-term contracting volume, year to date, UxC as quoted by Cameco, first print: document release date", "release date", "do not use", "dating field", "INFO", _D),
    ("SPUT units daily", "Issue price if known [USD]", "V023", "keep", "", "DAY", "B"),
    ("SPUT units daily", "Grade", "source flag", *_DESC, _D), ("SPUT units daily", "Document or URL", "source flag", *_DESC, _D),
    ("SPUT units daily", "Page or row", "source flag", *_DESC, _D), ("SPUT units daily", "Note", "", *_DESC, _D), ("SPUT units daily", "Raw file", "", *_DESC, _D),
    ("UPC holdings changes", "Previous as of", "D44", "do not use", "dating field", "INFO", _D),
    ("UPC holdings changes", "U3O8 held before [lb]", "D44", "copy of UPC monthly NAV U3O8 held [lb] (previous as-of date)", "", "ME", "C"),
    ("UPC holdings changes", "U3O8 held after [lb]", "D44", "copy of UPC monthly NAV U3O8 held [lb] (as-of date)", "", "ME", "C"),
    ("UPC holdings changes", "UF6 held before [KgU]", "D44", "copy of UPC monthly NAV UF6 held (previous as-of date)", "", "ME", "C"),
    ("UPC holdings changes", "UF6 held after [KgU]", "D44", "copy of UPC monthly NAV UF6 held (as-of date)", "", "ME", "C"),
    ("UPC holdings changes", "Type as written by the Data Curator", "D44", *_DESC, _D), ("UPC holdings changes", "What the release says", "D44", *_DESC, _D),
    ("UPC holdings changes", "First public document", "D44", *_DESC, _D), ("UPC holdings changes", "First public proof", "D44", *_DESC, _D),
    ("UPC holdings changes", "Release document", "D44", *_DESC, _D), ("UPC holdings changes", "Release URL", "D44", *_DESC, _D),
    ("UPC holdings changes", "Grade of change", "D44", *_DESC, _D), ("UPC holdings changes", "Grade of first public date", "D44", *_DESC, _D),
]
REMOVED = {("SPUT statements", "Net proceeds, proceeds minus costs [USD thousand]"), ("SPUT statements", "Average issue price [USD]"), ("SPUT statements", "Pounds per unit [lb]"),
           ("SPUT statements", "Uranium oxide fair value per lb [USD]"), ("SPUT statements", "Cash as % of NAV"), ("Quarterly", "SPUT purchases as % of UxC spot volume, quarter"),
           ("Quarterly", "Yellow Cake premium/discount to NAV, quarter-end")}  # 20 Nov 2026 (answers): the Yellow Cake ratio column is removed (Ronaldo's answer 2)
DATA_SHEETS = ("Monthly", "Quarterly", "Annual", "Weekly", "Daily", "World Bank monthly")


def sheet_columns(wb):
    """(sheet, header) -> (letter, first key, last key) for every data column of v6. Data sheets: header row 7, data from row 13; tables: header row 5."""
    out = {}
    for ws in wb.worksheets:
        dup = set()
        if ws.title in DATA_SHEETS:
            hr, dr = 7, 13
            names = [ws.cell(hr, j).value for j in range(2, ws.max_column + 1)]
            dup = {n for n in names if n and names.count(n) > 1}  # e.g. Annual: two columns per IAEA series (units and MW(e))
        elif ws.title in ("Read me", "Contents", "Variables", "Gaps"):
            continue
        else:
            hr, dr = 5, 6
        if ws.title in ("Prices OHLCV", "Physical sources"):
            continue
        data = ws.title in DATA_SHEETS
        keys = [ws.cell(r, 1).value for r in range(dr, ws.max_row + 1)] if data else []
        for j in range(2 if data else 1, ws.max_column + 1):
            h = ws.cell(hr, j).value
            if not h:
                continue
            ks = [keys[i] for i, r in enumerate(range(dr, ws.max_row + 1)) if ws.cell(r, j).value not in (None, "")] if data else []
            hh = f"{h} [{ws.cell(12, j).value}]" if h in dup else str(h)
            out[(ws.title, hh)] = (get_column_letter(j), ks[0] if ks else None, ks[-1] if ks else None)
    return out


def _fmt(k):
    if k is None:
        return None
    if isinstance(k, dt.datetime):
        k = k.date()
    return k.isoformat() if isinstance(k, dt.date) else str(k)


def label_rows(wb):
    """The Curator's label table, one row per v6 data column, with v6 letters. Returns rows and counts."""
    L = read_csv(F_LABELS)
    cols = sheet_columns(wb)
    v5 = {}
    cnt = defaultdict(int)
    for r in L:
        cnt[(r["sheet"], r["header"])] += 1
    for r in L:
        # the same series name twice on a sheet (Annual IAEA units and MW(e)) is told apart by the unit, as in sheet_columns
        r["_key"] = f"{r['header']} [{r['unit']}]" if cnt[(r["sheet"], r["header"])] > 1 and r["column_letter_v5"] else r["header"]
        if r["column_letter_v5"]:
            v5[(r["sheet"], r["column_letter_v5"])] = r["_key"]
    v6 = {}
    for (sh, h), (let, _a, _b) in cols.items():
        v6[(sh, h)] = let
        v6.setdefault((sh, h.split(" [")[0]), let)

    def v6_of(sheet, header):
        return v6.get((sheet, header)) or v6.get((sheet, header.split(" [")[0]))

    sheets_all = ["Monthly", "Quarterly", "Annual", "Weekly", "Daily", "World Bank monthly", "SPUT purchases", "SPUT issuance", "SPUT statements", "UxC volumes",
                  "Yellow Cake reports", "UPC monthly NAV", "UPC holdings changes", "SPUT units daily"]

    def translate(sheet, text):
        """Replace v5 column letters in 'copy of E' style text by the v6 letter and the series name."""
        if " of " not in text:
            return text
        head, tail = text.split(" of ", 1)

        def sub(m):
            sh = m.group(1) or sheet
            let = m.group(2)
            h = v5.get((sh, let))
            if not h:
                return m.group(0)
            n6 = v6_of(sh, h)
            pre = "" if sh == sheet else sh + " "
            return f"{pre}{n6 or '(not in v6)'} ({h})"
        pat = r"(?:(" + "|".join(re.escape(s) for s in sheets_all) + r")!)?\b([A-Z]{1,3})\b"
        return head + " of " + re.sub(pat, sub, tail)

    out, missing, cov_diff, spot_set = [], [], 0, 0
    for r in L:
        sh, h = r["sheet"], r["_key"]
        if (sh, r["header"]) in REMOVED:
            continue
        if not r["column_letter_v5"]:
            if h not in NEW_LABEL_MAP:
                missing.append(h)
                continue
            sh, h = NEW_LABEL_MAP[h]
        let = v6_of(sh, h) if not sh.startswith("(") else "-"
        if sh == "Events":
            let = "all"
        if let is None:
            missing.append(f"{sh} / {h}")
            continue
        first, last = r["first_date_with_data"] or None, r["last_date_with_data"] or None
        c6 = cols.get((sh, h)) or next((v for k, v in cols.items() if k[0] == sh and k[1].split(" [")[0] == h), None)
        note = ""
        if c6 and sh in DATA_SHEETS and c6[1] is not None:
            f6, l6 = _fmt(c6[1]), _fmt(c6[2])
            if (f6, l6) != (first, last):
                note = f"Coverage in v6 ({f6} to {l6}) differs from the Curator's table ({first} to {last}): shown as in v6."
                first, last = f6, l6
                cov_diff += 1
        spot = r["spot_link"]
        if spot == "not in spot_link.csv":
            spot = "C (set 20 Nov 2026 by Ronaldo's rule for unmarked columns)"
            spot_set += 1
        letter = let + (f" (v5: {r['column_letter_v5']})" if r["column_letter_v5"] and r["column_letter_v5"] != let else "")
        out.append([sh, letter, h, r["variable_id"], translate(sh, r["model_use"]), r["do_not_use_reason"] or None, first, last, r["timing_code"], spot, note or None])
    for sh, h, vid, use, why, tim, spot in CODER_LABELS:
        c6 = cols[(sh, h)] if (sh, h) in cols else next(v for k, v in cols.items() if k[0] == sh and k[1].split(" [")[0] == h)
        out.append([sh, c6[0], h, vid, use, why or None, _fmt(c6[1]), _fmt(c6[2]), tim, spot,
                    "Label added by the Coder on 20 Nov 2026 (no row in the Curator's table; same rules as the Curator's rows)."])
    if missing:
        raise SystemExit("labels: no v6 column for " + "; ".join(missing))
    # every v6 data column must have a label
    labelled = {(x[0], x[2]) for x in out}
    unl = [k for k in cols if k[0] in DATA_SHEETS + ("SPUT purchases", "SPUT issuance", "SPUT statements", "UxC volumes", "Yellow Cake reports", "UPC monthly NAV", "UPC holdings changes", "SPUT units daily")
           and k not in labelled and (k[0], k[1].split(" [")[0]) not in labelled and k[1] not in ("Date", "Purchase date", "Issue date", "Period end", "Announcement date", "Release date", "As of", "Change no.")]
    return out, dict(rows=len(out), coverage_updated=cov_diff, spot_set_c=spot_set, unlabelled=unl)


LABEL_HEADS = ["Sheet", "Column in v6", "Series", "Variable ID", "Model use", "Do-not-use reason", "First date with data", "Last date with data", "Timing", "Spot link", "Note"]


def write_labels(ws, rows, font, fill, FOREST, SAGE):
    """Append the column labels below the register on the Variables sheet."""
    r = ws.max_row + 2
    ws.cell(r, 1, "Column labels for workbook v6: one row per data column (Data Curator's label table of 20 Nov 2026)").font = font(12, bold=True, color=FOREST)
    r += 1
    c = ws.cell(r, 1, ("Model use: keep; copy of (same values as another column); arithmetic of (built from other columns by a sum or difference); close relative of (the same quantity measured another way: "
                       "average against month-end, another vintage, another provider); do not use (with the reason). Column letters are those of v6 (v5 letter in brackets where it moved). "
                       "Timing codes are defined on Contents. Spot link: A = contains the uranium spot price by construction, B = partly linked, C = no construction link."))
    c.font = font(9, italic=True)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    ws.row_dimensions[r].height = 40
    r += 1
    for j, h in enumerate(LABEL_HEADS, 1):
        cc = ws.cell(r, j, h)
        cc.font, cc.fill = font(10, bold=True, color="FFFFFF"), fill(FOREST)
        cc.alignment = Alignment(wrap_text=True, vertical="center")
    # column K (Note) is hidden on Variables with the handoff block; the note is repeated as a comment on the Series cell
    r += 1
    for row in rows:
        for j, v in enumerate(row, 1):
            cc = ws.cell(r, j, v)
            cc.font = font(9, bold=(j == 3))
            cc.alignment = Alignment(wrap_text=True, vertical="top")
        if row[10]:
            ws.cell(r, 3).comment = Comment(row[10], "Dataset", width=320, height=90)
        if str(row[4]).startswith("do not use"):
            ws.cell(r, 5).fill = fill("F6D9D9")
        r += 1
    return r


# ---------------------------------------------------------------- 5. Read me
README = [
    ("What this workbook is", [
        "The data for Ronaldo Silva's Master's thesis (ISCAP) on the uranium price and the funds that hold physical uranium. Sprott Physical Uranium Trust (SPUT, from July 2021) is at the centre; Uranium Participation Corporation (UPC, 2005 to 2021, the fund SPUT took over) and Yellow Cake plc (from July 2018) are kept for comparison.",
        "Alongside them: other commodities, financial markets, macro data, the physical uranium market and a dated list of market events.",
    ]),
    ("How the sheets are organised", [
        "One sheet per frequency (decision D-024): Monthly, Quarterly, Annual, Weekly and Daily hold one column per series and one row per period. A series sits in the sheet of its own publication frequency.",
        "Record sheets hold one row per document or event: SPUT purchases, SPUT issuance, SPUT statements, SPUT units daily, UxC volumes, Yellow Cake reports, UPC monthly NAV, UPC holdings changes, Events and Physical sources.",
        "Prices OHLCV and World Bank monthly are source files kept whole. Contents lists every sheet; Variables lists every variable and, at the bottom, every column with its labels; Gaps lists what is missing.",
    ]),
    ("Dates", [
        "On the Monthly sheet the date shown is the 1st of the month, but the values are month-end values or values for the whole month: join by year and month, not by day.",
        "Quarterly rows are dated at the quarter end. Row 'Timing' on each data sheet says what kind of number a column holds (month-end, average, flow and so on); the codes are defined on Contents.",
    ]),
    ("Cut-off: 31 October 2026", [
        "Every series stops at 31 October 2026. Values after that date stay in the raw files in the repo (for example SPUT daily prices to 13 November 2026).",
        "First prints (the ALFRED macro columns and Cameco's long-term contracting figures) are cut by the date they were published: a value first published after 31 October 2026 is left out even if it refers to an earlier month.",
        "Forward-looking columns (the EIA figures for 2026 to 2035 on the Annual sheet) are kept, because they were published before 31 October 2026.",
    ]),
    ("How to read the labels", [
        "At the bottom of the Variables sheet each data column has: model use (keep; copy of another column; arithmetic of other columns; close relative of another column; do not use, with the reason), its first and last date, its timing code and its spot link.",
        "Spot link: A = the column contains the uranium spot price by construction (never use it to predict the spot), B = partly linked (use only lagged and with care), C = no construction link.",
        "Row 'Release' says whether the dates on which figures became public are in data/release_dates.csv, so a forecast never uses a figure before it was known. The first-print columns carry their own publication date in the column beside them.",
    ]),
    ("Colours", [
        "Light orange cell or row = value from an unofficial source: the public 'SPUT Tracker' Google Sheet, which copies the Sprott fund page every day. It was checked against Sprott's own documents wherever they overlap (every overlap matched, apart from one printed figure flagged on 20 Jan 2026). Grade C.",
        "Row 'Type · coverage' is coloured by type: native, derived (month-end of a daily series), average, gap fill, reconstructed.",
    ]),
    ("Inputs only", [
        "The workbook holds inputs, not results: no returns, ratios or statistics (decision D-029). The calculations are Ronaldo's.",
        "A few columns are sums or month-end values of other published figures (they say so in their notes). The six ratio columns of version 5 were removed on 20 November 2026, and so was the Yellow Cake premium/discount to NAV column (Quarterly) later the same day; their source figures are all still here.",  # 20 Nov 2026 (answers)
    ]),
    ("Sources", [
        "Every value has a source document or saved page, listed in data/source_manifest.csv in the repo (decision D-026); every raw file is listed with its SHA-256 hash in data/raw_manifest.csv. The exceptions are listed on the Gaps sheet.",
        "Hover over a series name for its note and raw file.",
        "For supervision only: some sources (Sprott, CME, the Baltic Exchange) restrict redistribution.",
    ]),
]


def write_readme(wb, built, font, fill, FOREST, GREY, CHAR):
    ws = wb.create_sheet("Read me")
    ws.sheet_properties.tabColor = FOREST
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Read me"
    ws["A1"].font = font(18, bold=True, color=FOREST)
    ws["A2"] = f"Uranium financialization data set, version 6. Built {built}. Cut-off 31 October 2026."
    ws["A2"].font = font(10, italic=True, color=GREY)
    r = 4
    for title, paras in README:
        ws.cell(r, 1, title).font = font(12, bold=True, color=FOREST)
        r += 1
        for p in paras:
            c = ws.cell(r, 1, p)
            c.font = font(10, color=CHAR)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 15 * max(1, len(p) // 120 + 1)
            r += 1
        r += 1
    ws.column_dimensions["A"].width = 130
    return ws
