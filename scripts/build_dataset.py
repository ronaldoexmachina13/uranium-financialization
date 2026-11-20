"""Build data/processed/thesis_dataset.xlsx from the raw files in data/raw/.

Run from the repo root after the fetch scripts:  python scripts/build_dataset.py      (needs openpyxl)

Layout: one sheet per native frequency, nothing aggregated away.
  Daily, Weekly, Monthly, Quarterly, Events_SPUT (irregular), Daily_OHLCV (long table), Monthly_Pink (World Bank averages),
  plus README, Catalogue (every column described), Analysis_Map and Gaps.
Rules:
  - A series sits in the sheet of its own frequency. Month-end versions of daily series are added to Monthly, labelled "derived".
  - One source per column; sources are never mixed or averaged. Gap-fill and reconstructed values have their own columns.
  - Month-end means the last available observation in the month. Monthly averages are used only where no daily source exists.
  - Blank cell = no observation (market closed, series not started, source missing). Nothing is interpolated or forward-filled.
"""
import calendar
import csv
import os
import re
import datetime as dt
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet._reader import WorksheetReader

import fetch_equities as fe
import fetch_markets as fm
from coverage_map import ROWS as COVERAGE
from gaps import GAPS
from handoff_table import CHAT
import workbook_layout as wb_layout

WorksheetReader.bind_tables = lambda self: None  # the World Bank file has table links openpyxl cannot read

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
DATED = RAW / "markets_dated"   # futures with repaired dates (15 Nov 2026); the older files in data/raw/markets stay as evidence
DATED_DAY = "20261115"
OUT = Path(os.environ["THESIS_OUT"]) if os.environ.get("THESIS_OUT") else ROOT / "data" / "processed" / "thesis_dataset_v4.xlsx"

NUM, NUM4, INT, PCT = "#,##0.00", "#,##0.00##", "#,##0", "0.00"
FILL = {"constructed": "E4DFEC", "native": "D9E2D5", "derived": "DDEBF7", "average": "EDEDED", "gap-fill": "FFF2CC", "reconstructed": "FCE4D6"}
FONT, BOLD = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)

FRED_DAILY = [  # (header, FRED id, unit, role)
    ("VIX", "VIXCLS", "index points", "H4"), ("Broad dollar index", "DTWEXBGS", "index", "H4"),
    ("Fed funds rate", "DFF", "percent", "H4"), ("US 10y yield", "DGS10", "percent", "H4"),
    ("US 10y real yield (TIPS)", "DFII10", "percent", "H4"), ("Term spread 10y-3m", "T10Y3M", "percentage points", "H4"),
    ("Baa-10y credit spread", "BAA10Y", "percentage points", "H4"), ("Breakeven inflation 10y", "T10YIE", "percent", "H4"),
    ("USD/CAD", "DEXCAUS", "CAD per USD", "H4 / control"), ("S&P 500 (FRED, 10 years only)", "SP500", "index", "H4 / control"),
]
FRED_DAILY_EXTRA = [  # fetched by fetch_fred.py (added 8 Nov 2026); used only if the file exists
    ("3-month Treasury yield", "DGS3MO", "percent", "H4"), ("Henry Hub natural gas, spot", "DHHNGSP", "USD per MMBtu", "H1, H2"),
    ("GBP/USD", "DEXUSUK", "USD per GBP", "control"), ("Corporate option-adjusted spread (ICE BofA)", "BAMLC0A0CM", "percent", "H4"),
]
FRED_MONTHLY_EXTRA = [  # (header, id, unit, role): monthly values as published
    ("Fed funds rate, monthly", "FEDFUNDS", "percent", "H4"), ("Consumer price index (CPI-U, seasonally adjusted)", "CPIAUCSL", "index 1982-84=100", "control"),
    ("Baa corporate yield", "BAA", "percent", "H4"), ("Aaa corporate yield", "AAA", "percent", "H4"),
    ("Broad dollar index, real", "RTWEXBGS", "index", "H4"), ("US industrial production", "INDPRO", "index 2017=100", "H4 / control"),
    ("Economic policy uncertainty, US", "USEPUINDXM", "index", "H4"), ("Economic policy uncertainty, global", "GEPUCURRENT", "index", "H4"),
]
OIL = [("WTI crude, spot", "DCOILWTICO", "H1, H2"), ("Brent crude, spot", "DCOILBRENTEU", "H1, H2")]
FRED_MONTHLY = [  # monthly averages from the IMF, via FRED
    ("Uranium (IMF)", "PURANUSDM", "USD per lb", "cross-check of the Cameco price; runs about 20% below it from 2021"),
    ("Coal, Australia (IMF)", "PCOALAUUSDM", "USD per tonne", "average; no month-end source retrieved"),
    ("Copper (IMF)", "PCOPPUSDM", "USD per tonne", "average; use the COMEX month-end column for analysis"),
    ("Natural gas, Henry Hub (IMF)", "PNGASUSUSDM", "USD per MMBtu", "average; use the NYMEX month-end column for analysis"),
]
MARKET_NAMES = {  # ticker -> (header, unit, role)
    "GC=F": ("Gold futures (GC=F)", "USD per oz", "H1, H2"), "XAUUSD=X": ("Gold spot (XAUUSD=X)", "USD per oz", "H1, H2"),
    "SI=F": ("Silver futures (SI=F)", "USD per oz", "H1, H2"), "XAGUSD=X": ("Silver spot (XAGUSD=X)", "USD per oz", "H1, H2"),
    "HG=F": ("Copper futures (HG=F)", "USD per lb", "H1, H2"), "PL=F": ("Platinum futures (PL=F)", "USD per oz", "H1, H2"),
    "CL=F": ("WTI futures (CL=F)", "USD per barrel", "H1, H2"), "BZ=F": ("Brent futures (BZ=F)", "USD per barrel", "H1, H2"),
    "NG=F": ("Natural gas futures (NG=F)", "USD per MMBtu", "H1, H2"), "^GSPC": ("S&P 500 (^GSPC)", "index", "H4 / control"),
}
EQUITIES = [("URA", "USD", "ETF"), ("URNM", "USD", "ETF"), ("NLR", "USD", "ETF"), ("CCJ", "USD", "producer"), ("KAP.IL", "USD", "producer"),
            ("U-U.TO", "USD", "physical fund"), ("U-UN.TO", "CAD", "physical fund"), ("YCA.L", "GBp", "physical fund"), ("UEC", "USD", "developer"),
            ("DNN", "USD", "developer"), ("NXE", "USD", "developer"), ("PDN.AX", "AUD", "producer"), ("BOE.AX", "AUD", "producer"),
            ("DYL.AX", "AUD", "developer"), ("SLX.AX", "AUD", "enrichment tech"), ("BWXT", "USD", "nuclear supply"), ("1164.HK", "HKD", "producer")]
SPLIT_DATE = dt.date(2021, 7, 26)  # first trading day of SPUT; before it the Yahoo U-UN.TO series is Uranium Participation Corp
EQ_NOTE = {"U-UN.TO": "SPUT only, from 26 Jul 2021 (first SPUT row: close 11.14, volume 106,200; the ticker's earlier history is Uranium Participation Corp, see the UPC column; the five stale rows of 19 to 23 Jul 2021 between the two are blank here and kept in Prices OHLCV)", "U-U.TO": "USD units of the Sprott trust",
           "YCA.L": "prices in pence (GBp)", "KAP.IL": "London GDR in USD"}


# ---------------------------------------------------------------- helpers
def read_csv(path):
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def num(v):
    if v is None or str(v).strip() in ("", ".", "None", "nan", "…"):
        return None
    return float(v)


def day(s):
    return dt.date.fromisoformat(str(s)[:10])


def ym(d):
    return (d.year, d.month)


def month_end(d):
    out = {}
    for k in sorted(d):
        out[ym(k)] = d[k]
    return out


def col(header, values, kind, source, sid, freq, timing, unit, note="", role="", raw="", fmt=NUM4):
    return dict(header=header, values={k: v for k, v in values.items() if v is not None and v != ""}, kind=kind, source=source, sid=sid,
                freq=freq, timing=timing, unit=unit, note=note, role=role, raw=raw, fmt=fmt)


def daily_series(path, datecol, valcol):
    return {day(r[datecol]): num(r[valcol]) for r in read_csv(path) if num(r[valcol]) is not None}


def months_between(a, b):
    y, m = a
    while (y, m) <= b:
        yield (y, m)
        m += 1
        if m == 13:
            y, m = y + 1, 1


def quarters_between(a, b):
    out, y, m = [], a.year, a.month
    while (y, m) <= (b.year, b.month):
        out.append(dt.date(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1))
        m += 3
        if m > 12:
            y, m = y + 1, m - 12
    return out


ISS_FIRST, ISS_LAST = (2022, 3), (2025, 12)  # months fully covered by the Prior Sales tables of five prospectus documents (16 Feb 2022 to 19 Jan 2026, no gap between them)


stmt_src_name = "SPUT annual report 2024, Q2 2025, Q4 2025 and Q2 2026 reports; prospectus 22 Jan 2026"


def sput_fundraising(S, M, equities):
    """SPUT units, issuance, proceeds, cash and TSX statistics.
    The three CSVs in data/raw/sprott were typed from Sprott documents (see DATA_LOG.md); the build stops if one of the reconciliations below fails."""
    sp = RAW / "sprott"
    if not all((sp / f).exists() for f in ("sput_issuances.csv", "sput_statements.csv", "sput_tsx_monthly.csv")):
        return
    iss = sorted((day(r["date"]), num(r["price_usd_per_unit"]), int(r["units_issued"]), r["issue_type"]) for r in read_csv(sp / "sput_issuances.csv"))
    stm = read_csv(sp / "sput_statements.csv")
    tsx = read_csv(sp / "sput_tsx_monthly.csv")
    S["issuances"], S["statements"] = iss, stm
    S["issuance_src"] = {(day(r["date"]), num(r["price_usd_per_unit"]), int(r["units_issued"])): r["source"] for r in read_csv(sp / "sput_issuances.csv")}
    byend = {r["period_end"]: r for r in stm}
    U = lambda d: int(byend[d]["units_outstanding"])
    iu = lambda d: int(byend[d]["units_issued_in_period"] or 0)

    def need(ok, msg):
        if not ok:
            raise SystemExit("SPUT reconciliation failed: " + msg)
    need(U("2022-12-31") + iu("2023-12-31") == U("2023-12-31"), "units 2022 + issued 2023 = units 2023")
    need(U("2023-12-31") + iu("2024-12-31") == U("2024-12-31"), "units 2023 + issued 2024 = units 2024")
    need(U("2023-12-31") + iu("2024-06-30") == U("2024-06-30"), "units 2023 + issued H1 2024 = units Jun 2024")
    need(U("2024-12-31") + iu("2025-06-30") == U("2025-06-30"), "units 2024 + issued H1 2025 = units Jun 2025")
    need(U("2024-12-31") + iu("2025-12-31") == U("2025-12-31"), "units 2024 + issued 2025 = units 2025")
    need(U("2025-12-31") + iu("2026-06-30") == U("2026-06-30"), "units 2025 + issued H1 2026 = units Jun 2026")
    h2 = [(d, pr, n) for d, pr, n, _ in iss if dt.date(2025, 6, 30) < d <= dt.date(2025, 12, 31)]
    need(sum(n for _, _, n in h2) == U("2025-12-31") - U("2025-06-30"), "H2 2025 issuance list = change in reported units")
    need(abs(sum(pr * n for _, pr, n in h2) / 1000 - (float(byend["2025-12-31"]["proceeds_from_issuance_usd_thousand"]) - float(byend["2025-06-30"]["proceeds_from_issuance_usd_thousand"]))) < 1, "H2 2025 price x units = reported proceeds")
    def win(a, b):
        w = [(pr, n) for d, pr, n, _ in iss if a < d <= b]
        return sum(n for _, n in w), sum(pr * n for pr, n in w) / 1000
    D = dt.date
    for (a, b, key, exact) in [((2022, 12, 31), (2023, 6, 30), "2023-06-30", True), ((2023, 12, 31), (2024, 6, 30), "2024-06-30", True),
                               ((2023, 12, 31), (2024, 12, 31), "2024-12-31", True), ((2024, 12, 31), (2025, 6, 30), "2025-06-30", True)]:
        un, vk = win(D(*a), D(*b))
        need(un == iu(key), f"dated issuances to {key} = reported units issued")
        need(abs(vk - float(byend[key]["proceeds_from_issuance_usd_thousand"])) < 1, f"dated price x units to {key} = reported proceeds")
    h2_23 = win(D(2023, 6, 30), D(2023, 12, 31))[0]
    h2_23_rep = iu("2023-12-31") - iu("2023-06-30")
    pre = win(D(2021, 12, 31), D(2022, 12, 31))[0]
    need(iu("2022-12-31") - pre == iu("2022-06-30") - win(D(2021, 12, 31), D(2022, 6, 30))[0], "units not listed in 2022 all fall in the first half of 2022 (before 16 Feb 2022)")
    print(f"SPUT reconciliations OK: unit roll-forward; dated issuances equal the reported units and proceeds for H1 2023, H1 2024, FY2024, H1 2025 and H2 2025. Known differences: H2 2023 listed {h2_23:,} vs reported {h2_23_rep:,} units ({h2_23_rep - h2_23:,} not listed); {iu('2022-12-31') - pre:,} units issued 1 Jan-15 Feb 2022 are before the first table")

    mkey = lambda d: (d.year, d.month)
    me = lambda k: dt.date(k[0] + (k[1] == 12), k[1] % 12 + 1, 1) - dt.timedelta(days=1)
    full = list(months_between((2021, 8), (2021, 10))) + list(months_between(ISS_FIRST, ISS_LAST))  # Aug-Oct 2021: Prior Sales table of the 22 Nov 2021 supplement (17 Aug to 18 Nov 2021, reconciled to units outstanding); 19 Nov 2021 to 15 Feb 2022 not listed
    recm = list(months_between((2022, 6), ISS_LAST))
    um, vm = defaultdict(int), defaultdict(float)
    for d, pr, n, _ in iss:
        if mkey(d) in full:
            um[mkey(d)] += n
            vm[mkey(d)] += pr * n
    anchors = sorted(dt.date.fromisoformat(r["period_end"]) for r in stm if r["period_type"] in ("annual", "interim") and r["units_outstanding"] and r["period_end"] >= "2022-06-30")
    rec = {}
    for k in recm:
        a_d = max(a for a in anchors if a <= me(k))  # latest reported date on or before the month-end
        rec[k] = U(a_d.isoformat()) + sum(n for d, _, n, _ in iss if a_d < d <= me(k))
    need(rec[(2025, 12)] == U("2025-12-31"), "reconstructed Dec 2025 units = reported")
    # 2021: units from the arrangement (75,210,456 on 19 Jul 2021) plus dated issuances; the table total reconciles to the 162,558,000 units stated on 22 Nov 2021
    need(75210456 + sum(n for d, _, n, _ in iss if d <= dt.date(2021, 11, 18)) == 162558000, "2021 Prior Sales table sums to 162,558,000 units outstanding on 22 Nov 2021")
    for k in months_between((2021, 8), (2021, 10)):
        rec[k] = 75210456 + sum(n for d, _, n, _ in iss if d <= me(k))
    stmt_src = "SPUT annual and interim reports (Sprott)"
    raw_s = "data/raw/sprott/sput_statements.csv"
    ends = [r for r in stm if r["period_type"] in ("annual", "interim")]

    def rep_col(field, hdr, unit, fmt):
        return col(hdr, {mkey(day(r["period_end"])): num(r[field]) for r in ends}, "native", stmt_src, field, "monthly (reporting dates only)", "reporting date", unit,
                   "blank in months without an annual or interim report; Statements_SPUT has the full table", "H3", raw_s, fmt)
    M.append(rep_col("units_outstanding", "SPUT units outstanding, reported", "units", INT))
    M.append(rep_col("nav_usd_thousand", "SPUT total NAV, reported", "USD thousand", INT))
    M.append(rep_col("cash_usd_thousand", "SPUT cash, reported", "USD thousand", INT))
    M.append(col("SPUT units outstanding, month-end, reconstructed", rec, "reconstructed", "SPUT annual and interim reports and the prospectus Prior Sales tables", "units_outstanding", "monthly", "month-end",
                 "units", "units at the latest reported date (Jun/Dec) plus the dated issuances since; equals the reported figure at each reporting month. Jun 2022 to Dec 2025, and Aug to Oct 2021 (75,210,456 arrangement units plus dated issuances); in H2 2023 it omits 33,603 units issued that half-year that the prospectus table does not list", "H3",
                 "data/raw/sprott/sput_issuances.csv", INT))
    M.append(col("SPUT units issued in month", {k: um.get(k, 0) for k in full}, "derived", "SPUT prospectus supplements and base shelf prospectus (Prior Sales tables)", "units_issued", "monthly", "sum of issue dates in the month", "units",
                 "complete months only: Aug to Oct 2021 (17 Aug to 18 Nov 2021 listed) and Mar 2022 to Dec 2025 (16 Feb 2022 to 19 Jan 2026 listed without a gap); 19 Nov 2021 to 15 Feb 2022 is not listed, so those months are blank; 0 = no issue listed. H2 2023 lacks 33,603 units that the reports show as issued", "H3", "data/raw/sprott/sput_issuances.csv", INT))
    M.append(col("SPUT gross issuance value in month [USD]", {k: vm.get(k, 0.0) for k in full}, "derived", "SPUT prospectus supplements and base shelf prospectus (Prior Sales tables)", "price x units", "monthly", "sum of price x units in the month",
                 "USD", "an estimate of gross proceeds: it equals the cash-flow proceeds in the reports for H1 2023, H1 2024, FY2024, H1 2025 and H2 2025; net proceeds per issue are not published", "H3", "data/raw/sprott/sput_issuances.csv", INT))
    ts = [r for r in tsx if r["month"] < "2026-01"]
    for field, hdr, unit, fmt in [("uu_high_usd", "SPUT TSX high, U.U", "USD", NUM), ("uu_low_usd", "SPUT TSX low, U.U", "USD", NUM), ("uu_avg_daily_volume", "SPUT TSX average daily volume, U.U", "units", INT),
                                  ("uun_high_cad", "SPUT TSX high, U.UN", "CAD", NUM), ("uun_low_cad", "SPUT TSX low, U.UN", "CAD", NUM), ("uun_avg_daily_volume", "SPUT TSX average daily volume, U.UN", "units", INT)]:
        M.append(col(f"{hdr} [{unit}]", {(int(r["month"][:4]), int(r["month"][5:])): num(r[field]) for r in ts}, "native", "TMX via SPUT prospectus supplements (Feb 2023, Jan 2024, Dec 2024, Jun 2025) and base shelf prospectus 22 Jan 2026", field, "monthly", "calendar month", unit,
                     "Feb 2022 to Dec 2025; partial-month rows (1-14 Feb 2023, 1-2 Jan 2024, 1-5 Dec 2024, 1-17 Jun 2025, 1-20 Jan 2026) are left out. Yahoo daily volumes agree within 0.2% for every month", "H3", "data/raw/sprott/sput_tsx_monthly.csv", fmt))
    vol = defaultdict(float)
    for t in ("U-U.TO", "U-UN.TO"):
        for r in equities[t]:
            d = day(r["date"])
            if d >= dt.date(2021, 8, 1) and num(r["volume"]) is not None:
                vol[mkey(d)] += num(r["volume"])
    M.append(col("SPUT TSX volume, month total, U.U + U.UN [units]", dict(vol), "derived", "Yahoo Finance", "U-U.TO + U-UN.TO volume", "monthly", "sum of daily volume in the month", "units",
                 "both TSX lines; from Aug 2021 (first full month); averages of Yahoo daily volumes match the TSX monthly averages in the prospectuses within 0.2% for every month from Feb 2022 to Dec 2025", "H3", "data/raw/yahoo/U-U.TO.csv, U-UN.TO.csv", INT))
    # weekly holdings against the pounds in the reports (information only)
    wk = sorted(S["weekly"][0]["values"].items())
    for r in ends:
        if r["pounds_u3o8_held"]:
            d = day(r["period_end"])
            near = [v for k, v in wk if k <= d]
            if near:
                print(f"  holdings check {r['period_end']}: report {float(r['pounds_u3o8_held']) / 1e6:.3f} m lb vs last weekly value {near[-1]:.3f} m lb")


def uxc_volume(S, M, Q):
    """UxC spot-market volumes as quoted in Yellow Cake's quarterly updates (typed from the PDFs, 8 Nov 2026). First-reported value per period; all vintages stay in UxC_Volume_YC."""
    p = RAW / "yellowcake" / "uxc_spot_volume_reported.csv"
    if not p.exists():
        return
    rows = read_csv(p)
    S["uxc_rows"] = rows
    first = {}
    for r in sorted(rows, key=lambda r: r["announcement_date"]):
        if r["flag"].startswith("INCONSISTENT"):
            continue
        first.setdefault((r["period_type"], r["period_start"], r["period_end"]), r)
    src = "UxC figures quoted in Yellow Cake plc quarterly operating updates"
    raw = "data/raw/yellowcake/uxc_spot_volume_reported.csv"
    mon = {(int(k[1][:4]), int(k[1][5:7])): float(r["volume_mlb"]) for k, r in first.items() if k[0] == "month"}
    M.append(col("UxC spot volume, month, first reported [million lb]", mon, "native", src, "volume_mlb", "monthly", "calendar month", "million lb U3O8",
                 "only the months an update states (about 40 of 80 since 2020); first-reported value kept, revisions are in UxC_Volume_YC; whether both sides of a trade are counted is not stated", "H3", raw, NUM))
    qt = {day(k[2]): float(r["volume_mlb"]) for k, r in first.items() if k[0] == "quarter"}
    Q.append(col("UxC spot volume, quarter, first reported [million lb]", qt, "native", src, "volume_mlb", "quarterly", "calendar quarter", "million lb U3O8",
                 "quarters stated in an update; Q2 2025 left out because the figure in the update conflicts with the later updates (see UxC_Volume_YC); some are rounded or estimates", "H3", raw, NUM))
    cp = RAW / "physical" / "cameco_mda_spot_market_volume_reported_2006_2026.csv"
    if cp.exists():
        cq = {}
        for r in read_csv(cp):
            if "period reported in this filing" in r["country_or_scope"] and "-Q" in r["year"]:
                y_, q_ = r["year"].split("-Q")
                cq[day(f"{y_}-{int(q_) * 3:02d}-" + ("31" if int(q_) in (1, 4) else "30"))] = float(r["value"])
        Q.append(col("World spot market volume, quarter, as reported in Cameco's MD&A [million lb]", cq, "native", "Cameco quarterly MD&A (volume reported by UxC from 2015)", "value", "quarterly", "calendar quarter", "million lb U3O8",
                     "value printed in the MD&A of that quarter only (Q4 is not in any MD&A on file); rounded as printed ('about', 'just over'); comparison figures printed in later reports are in Physical sources", "H3", "data/raw/physical/cameco_mda_spot_market_volume_reported_2006_2026.csv", NUM))
    ev = S["events"]
    last_ev = max(e[0] for e in ev)
    sq, share = {}, {}
    for qd, v in qt.items():
        if qd < dt.date(2021, 9, 30) or qd > last_ev:
            continue
        a = sum(e[1] or 0 for e in ev if dt.date(qd.year, qd.month - 2, 1) <= e[0] <= qd)
        sq[qd] = a / 1e6
        share[qd] = a / (v * 1e6) * 100
    Q.append(col("SPUT purchases in quarter [million lb]", sq, "derived", "Sprott fund page", "pounds_acquired", "quarterly", "sum of purchase days in the quarter", "million lb U3O8", "for quarters with a UxC volume", "H3", "data/raw/sprott/sput_acquired.csv", NUM))
    Q.append(col("SPUT purchases as % of UxC spot volume, quarter", share, "derived", "Sprott fund page and UxC via Yellow Cake", "SPUT lb / UxC volume", "quarterly", "quarter", "percent",
                 "numerator from Sprott's purchase list (trade dates), denominator the first-reported UxC quarter volume; UxC volume is a provider-defined total and may include SPUT's own purchases", "H3", raw, NUM))
    # cross-checks against SPUT purchases quoted in the same updates
    qsum = lambda a, b: sum(e[1] or 0 for e in ev if a <= e[0] <= b) / 1e6
    print("UxC / SPUT checks (text says -> purchase list):",
          f"Q1 2023 2.38 -> {qsum(dt.date(2023,1,1), dt.date(2023,3,31)):.2f};", f"Mar 2022 5.4 -> {qsum(dt.date(2022,3,1), dt.date(2022,3,31)):.2f};",
          f"Q1 2026 ~5.3 -> {qsum(dt.date(2026,1,1), dt.date(2026,3,31)):.2f};", f"Q2 2026 1.35 -> {qsum(dt.date(2026,4,1), dt.date(2026,6,30)):.2f}")



def macro_cols(S):
    """Macro, uncertainty and activity series retrieved 8 Nov 2026 (data/raw/macro): V076, V082-V084, V088-V090, D21. Values as published; blanks stay blank."""
    d = RAW / "macro"
    if not (d / "macro_index.csv").exists():
        return
    M, Q, D = S["monthly"], S["quarterly"], S["daily"]
    def series(fname):
        return {day(r["date"]): num(r["value"]) for r in read_csv(d / fname) if num(r["value"]) is not None}
    def qend(k):
        m = ((k.month - 1) // 3) * 3 + 3
        return dt.date(k.year, m, calendar.monthrange(k.year, m)[1])
    def mon(fname):
        return {ym(k): v for k, v in series(fname).items()}
    raw = lambda f: f"data/raw/macro/{f}"
    # quarterly
    for f, hdr, unit, sid, note in [("fred_GDPC1.csv", "Real GDP, US", "billion chained 2017 USD, SAAR", "GDPC1", "BEA via FRED; US only"),
                                    ("fred_A191RL1Q225SBEA.csv", "Real GDP growth, US, annualised", "percent", "A191RL1Q225SBEA", "BEA via FRED, as published (not computed here)")]:
        Q.append(col(f"{hdr} [{unit}]", {qend(k): v for k, v in series(f).items()}, "native", "FRED (BEA)", sid, "quarterly", "quarter, dated at quarter end", unit, note, "", raw(f), NUM))
    # monthly
    for f, hdr, unit, sid, src, note in [
        ("fred_OECDPRINTO01IXOBSAM.csv", "Industrial production, OECD total", "index", "OECDPRINTO01IXOBSAM", "FRED (OECD)", "FRED series ends 2024-03"),
        ("fred_G7PRINTO01IXOBSAM.csv", "Industrial production, G7", "index", "G7PRINTO01IXOBSAM", "FRED (OECD)", "FRED series ends 2024-03"),
        ("fred_OECDPRINTO01GYSAM.csv", "OECD industrial production growth, year on year", "percent", "OECDPRINTO01GYSAM", "FRED (OECD)", "runs to 2026-07 although the level series ends 2024-03: check against the OECD portal before use"),
        ("epu_monetary_policy.csv", "Categorical EPU, monetary policy", "index", "EPU monetary", "policyuncertainty.com", ""),
        ("epu_fiscal_policy_taxes_or_spending.csv", "Categorical EPU, fiscal policy (taxes or spending)", "index", "EPU fiscal", "policyuncertainty.com", ""),
        ("epu_taxes.csv", "Categorical EPU, taxes", "index", "EPU taxes", "policyuncertainty.com", ""),
        ("epu_government_spending.csv", "Categorical EPU, government spending", "index", "EPU spending", "policyuncertainty.com", ""),
        ("epu_health_care.csv", "Categorical EPU, health care", "index", "EPU health", "policyuncertainty.com", ""),
        ("epu_national_security.csv", "Categorical EPU, national security", "index", "EPU security", "policyuncertainty.com", ""),
        ("epu_entitlement_programs.csv", "Categorical EPU, entitlement programs", "index", "EPU entitlements", "policyuncertainty.com", ""),
        ("epu_regulation.csv", "Categorical EPU, regulation", "index", "EPU regulation", "policyuncertainty.com", ""),
        ("epu_financial_regulation.csv", "Categorical EPU, financial regulation", "index", "EPU finreg", "policyuncertainty.com", ""),
        ("epu_trade_policy.csv", "Categorical EPU, trade policy", "index", "EPU trade", "policyuncertainty.com", ""),
        ("epu_sovereign_debt_currency_crises.csv", "Categorical EPU, sovereign debt and currency crises", "index", "EPU sovereign", "policyuncertainty.com", ""),
        ("gpr_GPR_monthly.csv", "Geopolitical risk index, overall (GPR)", "index 1985-2019=100", "GPR", "Caldara and Iacoviello", "recent GPR, 10 newspapers; 1985 onwards"),
        ("gpr_GPRT_monthly.csv", "Geopolitical risk index, threats (GPRT)", "index 1985-2019=100", "GPRT", "Caldara and Iacoviello", ""),
        ("gpr_GPRA_monthly.csv", "Geopolitical risk index, acts (GPRA)", "index 1985-2019=100", "GPRA", "Caldara and Iacoviello", ""),
        ("gpr_GPRH_monthly.csv", "Geopolitical risk index, historical overall (GPRH)", "index", "GPRH", "Caldara and Iacoviello", "from 1900; the file in data/raw/macro holds the full history"),
        ("dallasfed_IGREA_kilian.csv", "Kilian global real economic activity index", "percent deviation from trend", "IGREA", "Dallas Fed (Kilian)", "last 19 rows were stored with the wrong date in the source file and were set to the first of the month; values untouched"),
        ("googletrends_uranium_ETF.csv", "Google Trends search interest, 'uranium ETF'", "0-100, peak = 100", "uranium ETF", "Google Trends", "worldwide web search; one request for the whole period so no stitching; Nov 2026 incomplete; early years mostly 0")]:
        M.append(col(f"{hdr} [{unit}]", mon(f), "native", src, sid, "monthly", "monthly value as published", unit, note, "", raw(f), NUM))
    # 19 Nov 2026: the two global energy-uncertainty series, read from the dated file with corrected months (the original files store the month in the day field and stay unchanged).
    ef = d / "epu_energy_EUI_global_dates_corrected_2026-11-19.csv"
    if ef.exists():
        er = read_csv(ef)
        for s_, orig, lab in [("GDP_weighted", "epu_energy_EUI_global_EUI_GDP_weighted.csv", "GDP-weighted"), ("equally_weighted", "epu_energy_EUI_global_EUI_equally_weighted.csv", "equally weighted")]:
            rows_, ov, vals = [r for r in er if r["series"] == s_], series(orig), {}
            for r in rows_:
                if num(r["value"]) != ov.get(day(r["stored_date"])):
                    raise SystemExit(f"Energy uncertainty {s_}: value at stored date {r['stored_date']} differs from {orig}")
                k = ym(day(r["corrected_month"]))
                if k in vals:
                    raise SystemExit(f"Energy uncertainty {s_}: corrected month {k} appears twice")
                vals[k] = num(r["value"])
            if len(rows_) != len(ov) or sorted(vals) != list(months_between(min(vals), max(vals))):
                raise SystemExit(f"Energy uncertainty {s_}: corrected months are not one complete monthly run matching {orig}")
            M.append(col(f"Energy uncertainty index, global, {lab} (EUI) [index]", vals, "native", "Dang et al. (2023) via policyuncertainty.com", "EUI " + s_, "monthly", "monthly index value as published", "index",
                         "Energy-related uncertainty index of Dang, Nguyen, Lee, Nguyen and Le (2023, Energy Economics 124:106817), as named in data/raw/macro/macro_index.csv. The source file stores the month in the day field (YYYY-01-MM); "
                         f"the months here are corrected_month of the dated file (stored date beside it; every value equals {orig}, which is unchanged, D-005). The country-level file is not used: its dates cannot be read without the source.",
                         "", raw(ef.name), NUM))
    # daily (published on weekdays, or on every calendar day for the GPR)
    for f, hdr, unit, sid, src, note, me in [
        ("fred_TEDRATE.csv", "TED spread (discontinued)", "percent", "TEDRATE", "FRED", "3-month LIBOR minus 3-month T-bill; discontinued by FRED, ends 2022-01-21", True),
        ("fred_DCPF3M.csv", "Commercial paper, 90-day AA financial", "percent", "DCPF3M", "FRED (Federal Reserve)", "candidate replacement for the TED spread when paired with the T-bill column; the spread is not computed here", True),
        ("fred_DCPN3M.csv", "Commercial paper, 90-day AA nonfinancial", "percent", "DCPN3M", "FRED (Federal Reserve)", "alternative to the financial commercial paper column", True),
        ("fred_DTB3.csv", "T-bill, 3-month, secondary market, discount basis", "percent", "DTB3", "FRED (Federal Reserve)", "discount basis; the 3-month Treasury yield used elsewhere (DGS3MO) is constant-maturity", True),
        ("fred_SOFR.csv", "SOFR, secured overnight financing rate", "percent", "SOFR", "FRED (NY Fed)", "starts 2018-04-03", True),
        ("gpr_GPRD_daily.csv", "Geopolitical risk index, daily overall (GPRD)", "index 1985-2019=100", "GPRD", "Caldara and Iacoviello", "calendar days including weekends; the Daily date list has weekend rows from 1990 (the daily US EPU is published on every calendar day), so weekend values show from 1990 only; earlier weekend values are in the full file in data/raw/macro", False),
        ("gpr_GPRD_THREAT_daily.csv", "Geopolitical risk index, daily threats (GPRD_THREAT)", "index 1985-2019=100", "GPRD_THREAT", "Caldara and Iacoviello", "calendar days including weekends; the Daily date list has weekend rows from 1990 (the daily US EPU is published on every calendar day), so weekend values show from 1990 only; earlier weekend values are in the full file in data/raw/macro", False),
        ("gpr_GPRD_ACT_daily.csv", "Geopolitical risk index, daily acts (GPRD_ACT)", "index 1985-2019=100", "GPRD_ACT", "Caldara and Iacoviello", "calendar days including weekends; the Daily date list has weekend rows from 1990 (the daily US EPU is published on every calendar day), so weekend values show from 1990 only; earlier weekend values are in the full file in data/raw/macro", False)]:
        v = series(f)
        D.append(col(f"{hdr} [{unit}]", v, "native", src, sid, "daily", "daily value", unit, note, "", raw(f), NUM))
        if me:
            M.append(col(f"{hdr}, month-end [{unit}]", month_end(v), "derived", src, sid, "monthly", "last observation of the month (from Daily)", unit, note, "", raw(f), NUM))

    # Funding spread (V076): decision 9 Nov 2026 - commercial paper 90-day AA financial minus T-bill 3-month, on dates where both exist; not spliced to the TED spread
    cp, tb = series("fred_DCPF3M.csv"), series("fred_DTB3.csv")
    sp = {k: round(cp[k] - tb[k], 4) for k in cp if k in tb}
    D.append(col("Funding spread, commercial paper 90-day AA financial minus T-bill 3-month [percentage points]", sp, "derived", "FRED (Federal Reserve): DCPF3M minus DTB3", "CP-TBILL", "daily", "daily difference on days both rates exist",
                 "percentage points", "Replacement for the discontinued TED spread (chosen 9 Nov 2026). Differs in construction (commercial paper vs LIBOR; discount-basis bill): not spliced to TED, and the level is not comparable across the break", "", "data/raw/macro/fred_DCPF3M.csv; data/raw/macro/fred_DTB3.csv", NUM))
    M.append(col("Funding spread, commercial paper 90-day AA financial minus T-bill 3-month, month-end [percentage points]", month_end(sp), "derived", "FRED (Federal Reserve): DCPF3M minus DTB3", "CP-TBILL", "monthly",
                 "last observation of the month (from Daily)", "percentage points", "Replacement for the discontinued TED spread (chosen 9 Nov 2026); not spliced to TED", "", "data/raw/macro/fred_DCPF3M.csv; data/raw/macro/fred_DTB3.csv", NUM))
    # Baltic Dry Index (D18): last daily observation of each month, from the stock-marketdata.com series (1985 onwards)
    bdi_f = "baltic/baltic_dry_index_monthly.csv"
    bdi = {(int(r["month"][:4]), int(r["month"][5:7])): num(r["bdi_last_obs_of_month"]) for r in read_csv(RAW / bdi_f) if not r["note"]}
    M.append(col("Baltic Dry Index, month-end [index]", bdi, "derived", "stock-marketdata.com daily BDI series (Baltic Exchange index), read from its chart", "BDI", "monthly",
                 "last daily observation of the month", "index, 1985-01-04 = 1000", "Month-end values from the daily series; Jun to Oct 2026 agree with the Investing.com monthly table. Partial month (Nov 2026) left out. Redistribution of Baltic Exchange data is restricted: for thesis use only", "context",
                 "data/raw/" + bdi_f, NUM))

# ---------------------------------------------------------------- load
def load():
    S = {}
    S["daily"], S["weekly"], S["monthly"], S["quarterly"] = [], [], [], []
    cam = read_csv(RAW / "cameco" / "uranium_price_monthly.csv")
    # 19 Nov 2026: the raw file holds Cameco's Q4 2017 average in the Dec 2017 slot; the month-end printed in Cameco's 6-K exhibits is shown instead (raw file unchanged, D-005).
    cam_fix = {}
    cfx = RAW / "cameco" / "uranium_price_monthly_correction_2017-12.csv"
    if cfx.exists():
        camd = {r["date"]: r for r in cam}
        for r in read_csv(cfx):
            if r["date"] not in camd or num(camd[r["date"]][r["field"]]) != num(r["raw_value"]):
                raise SystemExit(f"Cameco correction {r['date']} {r['field']}: the raw file no longer holds {r['raw_value']}")
            cam_fix[(r["date"], r["field"])] = r
    S["cam_fix"] = cam_fix
    camv = lambda r, k: num(cam_fix[(r["date"], k)]["corrected_value"]) if (r["date"], k) in cam_fix else num(r[k])
    spot = {ym(day(r["date"])): camv(r, "spot_usd_lb") for r in cam}
    lt = {ym(day(r["date"])): camv(r, "longterm_usd_lb") for r in cam}
    M = S["monthly"]
    cam_raw = "data/raw/cameco/uranium_price_monthly.csv"
    cam_fix_note = (" Dec 2017: the raw file holds Cameco's Q4 2017 average in the December slot; this column shows the month-end printed in Cameco's 6-K exhibits (cell comment). Raw file unchanged; correction file data/raw/cameco/uranium_price_monthly_correction_2017-12.csv." if cam_fix else "")
    M.append(col("Uranium spot, month-end [USD/lb]", spot, "native", "Cameco (industry average of UxC and TradeTech month-end prices)", "spot_usd_lb", "monthly",
                 "month-end", "USD per lb", "MAIN PRICE." + cam_fix_note, "H1-H4 (dependent variable)", cam_raw, NUM))
    M.append(col("Uranium long-term, month-end [USD/lb]", lt, "native", "Cameco (industry average of UxC and TradeTech month-end prices)", "longterm_usd_lb", "monthly",
                 "month-end", "USD per lb", cam_fix_note.strip(), "context", cam_raw, NUM))
    for hdr, sid, unit, note in FRED_MONTHLY:
        d = {ym(day(r["observation_date"])): num(r[sid]) for r in read_csv(RAW / "fred" / f"{sid}.csv")}
        M.append(col(f"{hdr}, monthly avg [{unit}]", d, "average", "FRED (IMF)", sid, "monthly", "period average", unit, note, "cross-check", f"data/raw/fred/{sid}.csv", NUM))

    for hdr, sid, unit, role in FRED_MONTHLY_EXTRA:
        p = RAW / "fred" / f"{sid}.csv"
        if p.exists():
            d = {ym(day(r["observation_date"])): num(r[sid]) for r in read_csv(p)}
            M.append(col(f"{hdr} [{unit}]", d, "native", "FRED", sid, "monthly", "monthly value as published", unit, "", role, f"data/raw/fred/{sid}.csv", NUM))

    stq = read_csv(RAW / "stooq" / "ux_futures_monthly.csv")
    stq_note = "monthly bar; 11 months missing at source (2021-10 to 2022-03, 2023-04 to 2023-08); Nov 2026 is a partial month (not in this sheet)"
    for k, hdr, f in [("open", "open", NUM), ("high", "high", NUM), ("low", "low", NUM), ("close", "close", NUM)]:
        M.append(col(f"UX futures {hdr} (Stooq) [USD/lb]", {ym(day(r["date"])): num(r[k]) for r in stq}, "native", "Stooq UX.F", k, "monthly", "monthly bar", "USD per lb",
                     stq_note, "H1, H2", "data/raw/stooq/ux_futures_monthly.csv", NUM))
    M.append(col("UX futures volume (Stooq)", {ym(day(r["date"])): num(r["volume"]) for r in stq}, "native", "Stooq UX.F", "volume", "monthly", "monthly bar", "contracts",
                 "0 throughout: not usable as a participation measure", "H3", "data/raw/stooq/ux_futures_monthly.csv", INT))
    M.append(col("UX futures open interest (Stooq)", {ym(day(r["date"])): num(r["open_interest"]) for r in stq}, "native", "Stooq UX.F", "open_interest", "monthly", "monthly bar",
                 "contracts", "310 for Jan-Apr 2026 and 0 otherwise: not usable as a participation measure", "H3", "data/raw/stooq/ux_futures_monthly.csv", INT))
    inv = read_csv(RAW / "investing" / "ux_futures_monthly.csv")
    M.append(col("UX futures close (Investing, gap fill) [USD/lb]", {ym(day(r["date"])): num(r["close"]) for r in inv}, "gap-fill", "Investing.com (UXXc2 page)", "close", "monthly",
                 "monthly close", "USD per lb", "close only (its open/high/low are unusable); fills Jan-Mar 2022 that Stooq lacks; on shared months it differs from Stooq by at most 0.15",
                 "H1, H2", "data/raw/investing/ux_futures_monthly.csv", NUM))
    rec = read_csv(RAW / "uxc" / "ux_final_settlement_reconstructed.csv")
    M.append(col("UX final settlement, reconstructed from UxC benchmark [USD/lb]", {ym(day(r["month"] + "-01")): num(r["price_usd_lb"]) for r in rec}, "reconstructed",
                 "UxC benchmark via Cameco and Yellow Cake documents (user's handoff note)", "price_usd_lb", "monthly", "UxC reporting month-end (last Monday)", "USD per lb",
                 "8 months Stooq lacks (2021-10 to 2021-12, 2023-04 to 2023-08); not checked against CME bulletins; April 2023 calendar month-end would be 53.85",
                 "H1, H2", "data/raw/uxc/ux_final_settlement_reconstructed.csv", NUM))

    cm = read_csv(RAW / "cme" / "ux_volume_oi_monthly_cme_reports.csv")
    order = {"current month": 0, "previous month": 1, "same month prior year": 2}
    for metric, hdr, fmtx in [("volume", "UX futures volume, CME monthly report [contracts]", INT), ("open_interest", "UX futures open interest, CME monthly report [contracts]", INT)]:
        best = {}
        for r in cm:
            if r["metric"] != metric or "-" not in r["month"] or r["month"].endswith("YTD"):
                continue
            y, mth = r["month"].split("-")
            k = (int(y), int(mth))
            rk = order.get(r["how_in_report"], 3)
            if k not in best or rk < best[k][0]:
                best[k] = (rk, num(r["value"]))
        M.append(col(hdr, {k: v[1] for k, v in best.items()}, "native", "CME NYMEX/COMEX monthly Exchange Volume and Open Interest reports (Wayback Machine copies)", "cme_" + metric, "monthly",
                     "month total" if metric == "volume" else "month-end", "contracts",
                     "only the months the archived reports print (current, previous or same month a year earlier); no continuous series; all other months are blank, not zero. "
                     "Each month comes from the report that shows it as the current month where possible", "H3", "data/raw/cme/ux_volume_oi_monthly_cme_reports.csv", fmtx))

    mp = read_csv(RAW / "macro" / "sf_fed_monetary_policy_surprises_2012_2023.csv")
    for k, hdr, sid in [("Surprise", "Monetary-policy surprise, raw, sum over FOMC events in month [basis points]", "mps_raw"),
                        ("Orthogonalized Surprise", "Monetary-policy surprise, orthogonalized, sum over FOMC events in month [basis points]", "mps_orth")]:
        acc = {}
        for r in mp:
            kk = ym(day(r["Date"]))
            acc[kk] = acc.get(kk, 0) + num(r[k])
        M.append(col(hdr, {a: round(b, 2) for a, b in acc.items()}, "native", "San Francisco Fed, Monetary Policy Surprises (Bauer and Swanson high-frequency measure)", sid, "monthly", "sum of event values in the month", "basis points (as published)",
                     "event-level values (Jan 2012 to Dec 2023, 101 FOMC events incl. unscheduled 2020 meetings) summed per month; months without an event are blank, not zero. The page's chart file ends Dec 2023; the full Excel file was not read", "H4",
                     "data/raw/macro/sf_fed_monetary_policy_surprises_2012_2023.csv", NUM))
    nc = read_csv(RAW / "investing" / "newcastle_coal_futures_monthly.csv")
    M.append(col("Newcastle coal futures close (Investing NCFMc1), monthly bar [USD/tonne]", {ym(day(r["month"] + "-01")): num(r["close_usd_per_tonne"]) for r in nc if not r["note"]}, "native",
                 "Investing.com Newcastle Coal Futures (NCFMc1, ICE)", "ncf_close", "monthly", "monthly bar close", "USD per tonne",
                 "continuous front-month series from Jan 2009; Nov 2026 is a partial month and is left out; roll rule not documented by the source. Rotterdam API2 coal futures not retrieved", "H1 comparator",
                 "data/raw/investing/newcastle_coal_futures_monthly.csv", NUM))
    # daily series
    D = S["daily"]
    oil_me = []
    for hdr, sid, unit, role in FRED_DAILY + [x for x in FRED_DAILY_EXTRA if (RAW / "fred" / f"{x[1]}.csv").exists()]:
        d = daily_series(RAW / "fred" / f"{sid}.csv", "observation_date", sid)
        D.append(col(f"{hdr} [{unit}]", d, "native", "FRED", sid, "daily", "daily close" if sid != "DFF" else "daily rate", unit,
                     "FRED keeps about 10 years of the S&P 500; the full history is in the ^GSPC column" if sid == "SP500" else "", role, f"data/raw/fred/{sid}.csv", NUM))
    for hdr, sid, role in OIL:
        D.append(col(f"{hdr} [USD/barrel]", daily_series(RAW / "fred" / f"{sid}.csv", "observation_date", sid), "native", "FRED (EIA)", sid, "daily", "daily spot price", "USD per barrel",
                     "", role, f"data/raw/fred/{sid}.csv", NUM))
    if (RAW / "fred" / "USEPUINDXD.csv").exists():
        D.append(col("Economic policy uncertainty, US, daily [index]", daily_series(RAW / "fred" / "USEPUINDXD.csv", "observation_date", "USEPUINDXD"), "native", "FRED", "USEPUINDXD", "daily",
                     "daily index", "index", "", "H4", "data/raw/fred/USEPUINDXD.csv", NUM))
    sq = read_csv(RAW / "stooq" / "ux_futures_daily.csv")
    D.append(col("UX futures close, daily (Stooq) [USD/lb]", {day(r["date"]): num(r["close"]) for r in sq}, "native", "Stooq UX.F", "ux_close_d", "daily", "daily close", "USD per lb",
                 "continuous front contract, roll undocumented; no rows 28 Sep 2021 to 21 Jan 2022, 31 Mar to 11 Sep 2023 and 15 Sep to 2 Oct 2023; Stooq volume and open interest are not reliable and are not used", "H1, H2",
                 "data/raw/stooq/ux_futures_daily.csv", NUM))
    ig = read_csv(RAW / "investing" / "ux_daily_gapfill_2022_investing_UXXc2.csv")
    D.append(col("UX futures close, daily (Investing UXXc2, gap fill) [USD/lb]", {day(r["date"]): num(r["close"]) for r in ig}, "gap-fill", "Investing.com Uranium Futures UXXc2 (second continuous contract)", "ux_close_inv_d",
                 "daily", "daily close", "USD per lb", "24 Jan to 1 Apr 2022 only; a different continuous series from Stooq (roll differs after Mar 2022), kept in its own column", "H1, H2",
                 "data/raw/investing/ux_daily_gapfill_2022_investing_UXXc2.csv", NUM))
    cd = read_csv(RAW / "cme" / "ux_volume_oi_daily_cme_bulletins.csv")
    for k, hdr, unit, fmtx in [("ux_total_volume", "UX futures volume, CME bulletin [contracts]", "contracts", INT), ("ux_open_interest", "UX futures open interest, CME bulletin [contracts]", "contracts", INT),
                               ("front_settle", "UX front-month settlement, CME bulletin [USD/lb]", "USD per lb", NUM)]:
        D.append(col(hdr, {day(r["trade_date"]): num(r[k]) for r in cd}, "native", "CME daily bulletin, Section 62 (Wayback Machine copies)", "cme_" + k, "daily", "trade date", unit,
                     "20 Mar to 9 Nov 2026 only (archive copies); earlier days are blank, not zero. Volume is the sum over all listed contract months; open interest checked against Section 02B (13 days differ in the 02B check and are kept as in Section 62)" if k != "front_settle" else
                     "20 Mar to 9 Nov 2026 only; settlement of the nearest listed contract month", "H3" if k != "front_settle" else "H1, H2", "data/raw/cme/ux_volume_oi_daily_cme_bulletins.csv", fmtx))
    sd = read_csv(RAW / "sprott" / "sput_daily.csv")
    for k, hdr, unit, f in [("price_usd", "SPUT price", "USD", NUM), ("nav_usd", "SPUT NAV per unit", "USD", NUM), ("premium_discount_pct", "SPUT premium/discount", "percent", NUM)]:
        D.append(col(f"{hdr} [{unit}]", {day(r["date"]): num(r[k]) for r in sd}, "native", "Sprott fund page", k, "daily", "daily close", unit, "starts Jul 2021" + (" | FLAG: the NAV per unit is the same as on the previous day on 216 of 1,331 days (16.2%), so it is stale on those days (the NAV is rounded to the cent, and a stale NAV can make the prior NAV of an issuance row stale); the provider of the NAV price is not named in the Sprott documents" if k == "nav_usd" else ""), "H3", "data/raw/sprott/sput_daily.csv", f))
    mdir = RAW / "markets"
    markets = {}
    for t in fm.MARKETS:
        dated = t.endswith("=F")   # the seven futures series come from the date-repaired files; ^GSPC stays on the old file
        p = (DATED / f"{t}__fetched_{DATED_DAY}.csv") if dated else (mdir / f"{t}.csv")
        if dated and not p.exists():
            raise SystemExit(f"Missing date-repaired file {p}. The build does not fall back to the old files.")
        if p.exists():
            hdr, unit, role = MARKET_NAMES[t]
            markets[t] = read_csv(p)
            note = "continuous front-month; the contract changes at each roll, not back-adjusted" if t.endswith("=F") else ("spot quote from Yahoo" if t.endswith("=X") else "")
            if dated:
                note += ("; dates repaired on 15 Nov 2026 (each bar dated with the exchange time zone, file in data/raw/markets_dated; the older file in data/raw/markets had 28,459 of 44,275 rows one day early across the seven futures series)")
                stale = sum(1 for r in markets[t] if num(r["volume"]) == 0 and len({r["open"], r["high"], r["low"], r["close"]}) == 1)
                if stale:
                    note += f"; {stale} stale-looking days (open = high = low = close, volume 0) are kept and flagged, not removed"
                neg = [(r["date"], num(r["close"])) for r in markets[t] if num(r["close"]) is not None and num(r["close"]) <= 0]
                if neg:
                    note += "; negative close kept as published: " + ", ".join(f"{d_} {v_:.2f}" for d_, v_ in neg)
                if t == "PL=F":
                    note += "; THIN MARKET: Yahoo's reply holds 796 empty bars with no prices (mostly 2002 to 2009, not in this column) and long stretches between bars with prices, so month-end values can be stale"
            D.append(col(f"{hdr} close [{unit}]", {day(r["date"]): num(r["close"]) for r in markets[t]}, "native", "Yahoo Finance", t, "daily", "daily close", unit, note, role,
                         f"data/raw/markets_dated/{t}__fetched_{DATED_DAY}.csv" if dated else f"data/raw/markets/{t}.csv", NUM))
    equities = {}
    for t, cur, grp in EQUITIES:
        equities[t] = read_csv(RAW / "yahoo" / f"{t}.csv")
        # Yahoo repeats the previous price on days with no trading (volume 0): those are not observations, so they stay blank here and are kept in Prices OHLCV
        nz = sum(1 for r in equities[t] if (num(r["volume"]) or 0) == 0)
        note = EQ_NOTE.get(t, "")
        if nz:
            note = (note + "; " if note else "") + f"{nz} days with zero volume (Yahoo repeats the previous price) are left blank"
        vals = {day(r["date"]): num(r["adj_close"]) for r in equities[t] if (num(r["volume"]) or 0) > 0}
        if t == "U-UN.TO":
            upc = {d: v for d, v in vals.items() if d < SPLIT_DATE}
            vals = {d: v for d, v in vals.items() if d >= SPLIT_DATE}
            D.append(col(f"UPC adj. close [{cur}]", upc, "native", "Yahoo Finance", "UPC", "daily", "daily close (adjusted for splits and dividends)", cur,
                         f"Uranium Participation Corp, the fund SPUT replaced: the Yahoo U-UN.TO series up to {max(upc):%d %b %Y} (last trade); level agrees with UPC's reported NAV per share (e.g. CAD 4.87 at May 2021 against a price of about CAD 5); "
                         "no overlap with SPUT, so do not compute a return across the switch to the next column; days with zero volume are left blank. The five rows 19 to 23 Jul 2021 in data/raw/yahoo/U-UN.TO.csv are stale (open = high = low = close = 5.11, volume 0: the last real UPC close, 16 Jul 2021); they are blank here and kept in Prices OHLCV", "H3 context", "data/raw/yahoo/U-UN.TO.csv", NUM4))
        D.append(col(f"{t} adj. close [{cur}]", vals, "native", "Yahoo Finance", t, "daily", "daily close (adjusted for splits and dividends)", cur,
                     note, "H3 context" if grp in ("ETF", "physical fund") else "context", f"data/raw/yahoo/{t}.csv", NUM4))

    # weekly and events
    W = S["weekly"]
    hold = read_csv(RAW / "sprott" / "sput_holdings.csv")
    W.append(col("SPUT holdings [million lb U3O8]", {day(r["date"]): num(r["pounds_u3o8_millions"]) for r in hold}, "native", "Sprott fund page", "pounds_u3o8_millions", "weekly",
                 "weekly (Friday)", "million lb U3O8", "starts Jul 2021", "H3", "data/raw/sprott/sput_holdings.csv", NUM))
    S["events"] = [(day(r["date"]), num(r["pounds_acquired"]), num(r["cumulative_pounds_acquired"])) for r in read_csv(RAW / "sprott" / "sput_acquired.csv")]

    # monthly derived from daily / weekly / events
    drv = lambda c, hdr, extra="": col(hdr, month_end(c["values"]), "derived", c["source"], c["sid"], "monthly", "last observation of the month (from Daily)", c["unit"],
                                      (extra + " " + c["note"]).strip(), c["role"], c["raw"], c["fmt"])
    order = ["WTI crude, spot", "Brent crude, spot"] + [MARKET_NAMES[t][0] for t in fm.MARKETS if t in markets and t != "^GSPC"]
    for c in D:
        base = c["header"].split(" [")[0].replace(" close", "").replace(" adj. close", "")
        if base in order:
            M.append(drv(c, f"{base}, month-end [{c['unit']}]"))
    for c in D:
        if c["sid"] in [x[1] for x in FRED_DAILY + FRED_DAILY_EXTRA] + ["^GSPC"] and not c["header"].startswith("S&P 500 (FRED"):
            M.append(drv(c, c["header"].split(" [")[0].replace(" close", "") + f", month-end [{c['unit']}]"))
        if c["header"].startswith("S&P 500 (FRED"):
            M.append(drv(c, f"S&P 500 (FRED), month-end [index]"))
    for c in D:
        if c["sid"] in ("price_usd", "nav_usd", "premium_discount_pct"):
            M.append(drv(c, c["header"].split(" [")[0] + f", month-end [{c['unit']}]"))
    M.append(drv(W[0], "SPUT holdings, month-end [million lb U3O8]", "last weekly value of the month."))
    acq, last_acq = defaultdict(float), max(e[0] for e in S["events"])
    for d, p, _ in S["events"]:
        acq[ym(d)] += p or 0.0
    sput_months = list(months_between(min(M_ for M_ in month_end(D[[c["sid"] for c in D].index("price_usd")]["values"])), ym(max(D[[c["sid"] for c in D].index("price_usd")]["values"]))))
    M.append(col("SPUT pounds bought in month [lb U3O8]", {k: acq.get(k, 0.0) for k in sput_months}, "derived", "Sprott fund page", "pounds_acquired", "monthly", "sum of purchases in the month", "lb U3O8",
                 f"0 = no purchase listed; the purchase list ends {last_acq.isoformat()}", "H3", "data/raw/sprott/sput_acquired.csv", INT))
    sput_fundraising(S, M, equities)
    for t, cur, grp in EQUITIES:
        if t == "U-UN.TO":
            M.append(drv(next(c for c in D if c["sid"] == "UPC"), f"UPC adj. close, month-end [{cur}]"))
        c = next(c for c in D if c["sid"] == t)
        M.append(drv(c, f"{t} adj. close, month-end [{cur}]"))

    # quarterly
    Q = S["quarterly"]
    yc = read_csv(RAW / "yellowcake" / "nav_quarterly.csv")
    yraw = "data/raw/yellowcake/nav_quarterly.csv"
    ysrc = "Yellow Cake plc quarterly operating updates"
    Q.append(col("Yellow Cake NAV per share [GBP]", {day(r["quarter_end"]): num(r["nav_per_share_gbp"]) for r in yc}, "native", ysrc, "nav_per_share_gbp", "quarterly", "quarter-end", "GBP", "no monthly NAV is published", "H3", yraw, NUM))
    # 15 Nov 2026: 12.96 at 2021-03-31 is a pro forma figure. It stays as published; the corrected figure sits in the next column.
    held_fix = {}
    hc = RAW / "yellowcake" / "yellowcake_holdings_correction_2021-03-31.csv"
    if hc.exists():
        for r in read_csv(hc):
            held_fix[day(r["quarter_end"])] = int(re.match(r"\d+", r["corrected_value"]).group())
    held_flag = ("FLAG 31 Mar 2021: the value 12.96 is a pro forma figure (after the sale to Uranium Royalty Corp. and the Kazatomprom delivery), not the pounds held; it is left as published. "
                 "Actually held: 9,856,385 lb (grade A), shown in the next column. Do not use 12.96 as held units. Source: data/raw/yellowcake/yellowcake_holdings_correction_2021-03-31.csv.")
    # 19 Nov 2026: 13 quarters marked "derived" in nav_quarterly.csv have their pounds stated in the kept updates; basis "stated (document)" and the exact lb in its own column (raw file unchanged).
    yc_st = {}
    sfx = RAW / "yellowcake" / "yellowcake_holdings_stated_2026-11-19.csv"
    if sfx.exists():
        ycq = {r["quarter_end"]: r for r in yc}
        for r in read_csv(sfx):
            q_ = ycq.get(r["quarter_end"])
            if not q_ or q_["lb_basis"] != "derived":
                raise SystemExit(f"Yellow Cake stated holdings: {r['quarter_end']} is not a 'derived' row of nav_quarterly.csv")
            mlb = num(q_["uranium_held_million_lb"])
            # 0.0051 allows two-decimal rounding (24.374996 is printed as 24.38)
            if abs(mlb - int(r["exact_lb"]) / 1e6) > 0.0051 or abs(mlb - float(r["stated_value"].split()[0])) > 0.005:
                raise SystemExit(f"Yellow Cake stated holdings: {r['quarter_end']} {mlb} million lb differs from the stated {r['stated_value']} / {r['exact_lb']} lb by more than 0.005 million lb")
            yc_st[day(r["quarter_end"])] = r
    S["yc_stated"] = yc_st
    st_note = ("13 quarters from 30 Jun 2023 that nav_quarterly.csv marks 'derived' have the pounds stated in the kept updates (lb basis 'stated (document)'; exact lb in the column 'exact as stated'); no quarter is left derived. " if yc_st else "")
    Q.append(col("Yellow Cake uranium held [million lb U3O8]", {day(r["quarter_end"]): num(r["uranium_held_million_lb"]) for r in yc}, "native", ysrc, "uranium_held_million_lb", "quarterly", "quarter-end", "million lb U3O8", "see lb_basis. " + st_note + held_flag, "H3", yraw, NUM))
    if held_fix:
        Q.append(col("Yellow Cake uranium held, corrected [lb U3O8]", held_fix, "native", "Yellow Cake update of 20 Apr 2021 (page references in the correction file), read by the Data Curator on 15 Nov 2026", "holdings_correction", "quarterly", "quarter-end", "lb U3O8",
                     "Only the quarters where the published holdings are not what was held. 31 Mar 2021: 9,856,385 lb held; the published 12.96 million lb is pro forma.", "H3", "data/raw/yellowcake/yellowcake_holdings_correction_2021-03-31.csv", INT))
    if yc_st:
        Q.append(col("Yellow Cake uranium held, exact as stated [lb U3O8]", {k: int(r["exact_lb"]) for k, r in yc_st.items()}, "native", "Yellow Cake quarterly operating updates (kept PDFs; document and page in each cell comment), read 19 Nov 2026", "exact_lb", "quarterly", "quarter-end", "lb U3O8",
                     "Only the 13 quarters (30 Jun 2023 to 30 Sep 2026, except 31 Mar 2024) where nav_quarterly.csv says lb_basis = derived. Where an update prints only million lb, the exact figure follows the note in the source file (e.g. carried from the previous exact figure with no delivery in between); the note is repeated in the cell comment. Other quarters were not read for an exact figure.",
                     "H3", "data/raw/yellowcake/yellowcake_holdings_stated_2026-11-19.csv", INT))
    Q.append(col("Yellow Cake lb basis", {day(r["quarter_end"]): ("stated (document)" if day(r["quarter_end"]) in yc_st else r["lb_basis"]) for r in yc}, "native", ysrc, "lb_basis", "quarterly", "quarter-end", "text",
                 "stated = in the announcement; derived = holdings value divided by quarter-end spot; carried = previous quarter"
                 + ("; stated (document) = 'derived' in nav_quarterly.csv, but the pounds are stated in the kept update (19 Nov 2026; document and page in the cell comment; data/raw/yellowcake/yellowcake_holdings_stated_2026-11-19.csv). No quarter is left derived" if yc_st else ""), "H3", yraw, "@"))
    Q.append(col("Yellow Cake quarter-end spot used [USD/lb]", {day(r["quarter_end"]): num(r["spot_usd_lb"]) for r in yc}, "native", ysrc, "spot_usd_lb", "quarterly", "quarter-end", "USD per lb", "as given or derived for the announcement; compare with the Cameco quarter-end column", "H3", yraw, NUM))
    Q.append(col("Yellow Cake announcement date", {day(r["quarter_end"]): day(r["announcement_date"]) for r in yc if r["announcement_date"]}, "native", ysrc, "announcement_date", "quarterly", "date of announcement", "date", "", "H3", yraw, "yyyy-mm-dd"))
    conv = read_csv(RAW / "cameco" / "conversion_prices_quarterly.csv")
    craw = "data/raw/cameco/conversion_prices_quarterly.csv"
    csrc = "Cameco quarterly reports (average of TradeTech and UxC)"
    for hdr, k in [("Conversion spot, North America", "conv_spot_na_usd_kgu"), ("Conversion spot, Europe", "conv_spot_eu_usd_kgu"),
                   ("Conversion long-term, North America", "conv_lt_na_usd_kgu"), ("Conversion long-term, Europe", "conv_lt_eu_usd_kgu")]:
        Q.append(col(f"{hdr} [USD/kgU]", {day(r["quarter_end"]): num(r[k]) for r in conv}, "native", csrc, k, "quarterly", "quarter-end", "USD per kgU as UF6",
                     "complete quarter-ends from 2005-12; read from Cameco filings (see data/source_manifest.csv); no SWU prices", "context", craw, NUM))
    yca = sorted((day(r["date"]), num(r["close"])) for r in equities["YCA.L"] if num(r["close"]))
    prem = {}
    for r in yc:
        qd, nav = day(r["quarter_end"]), num(r["nav_per_share_gbp"])
        px = [c for d_, c in yca if d_ <= qd]
        if nav and px and (qd - max(d_ for d_, _ in yca if d_ <= qd)).days <= 7:
            prem[qd] = round((px[-1] / 100.0 / nav - 1) * 100, 2)
    Q.append(col("Yellow Cake premium/discount to NAV, quarter-end [percent]", prem, "derived", "YCA.L close (pence) divided by 100, against the announced NAV", "close / nav_per_share_gbp", "quarterly",
                 "last close on or before the quarter end (within 7 days)", "percent", "price in GBP vs NAV in GBP; negative = discount", "H3", yraw + " and data/raw/yahoo/YCA.L.csv", NUM))
    for k, hdr in [("spot_usd_lb", "Uranium spot"), ("longterm_usd_lb", "Uranium long-term")]:
        d = {day(r["date"]): camv(r, k) for r in cam if day(r["date"]).month in (3, 6, 9, 12)}
        Q.append(col(f"{hdr}, quarter-end (Cameco) [USD/lb]", d, "derived", "Cameco", k, "quarterly", "month-end value of Mar, Jun, Sep, Dec", "USD per lb", "from the Monthly sheet source" + ("; Dec 2017 corrected as on the Monthly sheet (cell comment)" if cam_fix else ""), "context", cam_raw, NUM))
    uxc_volume(S, M, Q)
    macro_cols(S)
    S["markets_found"] = list(markets)
    S["equities_raw"], S["markets_raw"] = equities, markets
    return S


def pink():
    wb_idx = openpyxl.load_workbook(RAW / "worldbank" / "CMO-Historical-Data-Monthly.xlsx", data_only=True)
    ws = wb_idx["Monthly Prices"]
    cols = []
    for c in range(2, ws.max_column + 1):
        name, unit = ws.cell(5, c).value, ws.cell(6, c).value
        if not name:
            continue
        name, unit = name.strip(), (unit or "").strip().strip("()")
        d = {}
        for r in range(7, ws.max_row + 1):
            p, v = ws.cell(r, 1).value, ws.cell(r, c).value
            if p and isinstance(v, (int, float)):
                d[(int(p[:4]), int(p[5:7]))] = float(v)
        cols.append(col(f"{name} [{unit}]", d, "average", "World Bank Pink Sheet", name, "monthly", "period average", unit, "monthly average as published; no month-end values exist",
                        "H1 cross-check" if name in ("Gold", "Silver", "Copper", "Platinum", "Crude oil, Brent", "Crude oil, WTI") else "", "data/raw/worldbank/CMO-Historical-Data-Monthly.xlsx", NUM))
    wi = wb_idx["Monthly Indices"]
    for c in range(2, wi.max_column + 1):
        parts = [str(wi.cell(r, c).value).strip() for r in range(6, 10) if wi.cell(r, c).value and str(wi.cell(r, c).value).strip()]
        if not parts:
            continue
        name = parts[-1].replace("**", "").strip()
        d = {}
        for r in range(10, wi.max_row + 1):
            p, v = wi.cell(r, 1).value, wi.cell(r, c).value
            if p and isinstance(v, (int, float)):
                d[(int(p[:4]), int(p[5:7]))] = float(v)
        cols.append(col(f"{name} (WB index) [2010=100]", d, "average", "World Bank Pink Sheet (Monthly Indices)", name, "monthly", "monthly index of period averages", "index 2010=100",
                        "index built from monthly averages", "H1 context", "data/raw/worldbank/CMO-Historical-Data-Monthly.xlsx", NUM))
    return cols


def annual_eia():
    """Annual EIA series (V092, V095-V100), read from data/raw/eia. W (withheld) and missing cells stay blank."""
    d = RAW / "eia"
    if not (d / "eia_table3_us_uranium_production.csv").exists():
        return []
    src = "EIA Forms EIA-851A and EIA-858"
    def val(x):
        try:
            return float(x)
        except (TypeError, ValueError):
            return None
    def mk(header, vals, unit, note, fmt=INT, raw=""):
        vals = {k: v for k, v in vals.items() if v is not None}
        return col(header, vals, "native", src, header, "annual", "calendar or year-end, as published", unit, note, "", raw, fmt)
    cols = []
    t3 = read_csv(d / "eia_table3_us_uranium_production.csv")
    f3 = "data/raw/eia/eia_table3_us_uranium_production.csv"
    for field, name, unit, fmt in [("production_thousand_lb_u3o8", "US uranium concentrate production", "thousand lb U3O8", INT), ("shipments_thousand_lb_u3o8", "US uranium concentrate shipments", "thousand lb U3O8", INT),
                                   ("producer_sales_deliveries_thousand_lb_u3o8", "US producer sales deliveries", "thousand lb U3O8", INT), ("weighted_avg_price_usd_per_lb", "US producer weighted-average sales price", "USD per lb U3O8", NUM)]:
        cols.append(mk(name, {int(r["year"]): val(r[field]) for r in t3}, unit, "Table 3, Domestic Uranium Production Report. W = withheld to avoid disclosing company data (left blank).", fmt, f3))
    inv = read_csv(d / "eia_table22_inventories_by_type.csv")
    f22 = "data/raw/eia/eia_table22_inventories_by_type.csv"
    names = {"Owners and operators of U.S. civilian nuclear power reactors": "US reactor owners inventory, total", "Uranium concentrate (U3O8), reactor owners": "US reactor owners inventory, uranium concentrate",
             "Natural UF6, reactor owners": "US reactor owners inventory, natural UF6", "Enriched UF6, reactor owners": "US reactor owners inventory, enriched UF6",
             "Fabricated fuel, reactor owners": "US reactor owners inventory, fabricated fuel", "U.S. supplier inventories": "US supplier inventory, total",
             "Uranium concentrate (U3O8), suppliers": "US supplier inventory, uranium concentrate", "Natural UF6, suppliers": "US supplier inventory, natural UF6",
             "Enriched UF6, suppliers": "US supplier inventory, enriched UF6", "Total commercial inventories": "US total commercial inventory"}
    for item, name in names.items():
        cols.append(mk(name, {int(r["year_end"]): val(r["thousand_lb_u3o8_eq"]) for r in inv if r["item"] == item}, "thousand lb U3O8 equivalent, end of year",
                       "Table 22, Uranium Marketing Annual Report 2025. Year-end 2021-2025; 2025 preliminary; W = withheld (blank).", INT, f22))
    own = read_csv(d / "eia_table23_inventories_by_owner.csv")
    for o, name in [("U.S. brokers and traders", "US brokers and traders inventory"), ("U.S. converter, enrichers, fabricators, and producers", "US converters, enrichers, fabricators and producers inventory")]:
        cols.append(mk(name, {int(r["year_end"]): val(r["thousand_lb_u3o8_eq"]) for r in own if r["owner"] == o}, "thousand lb U3O8 equivalent, end of year",
                       "Table 23, Uranium Marketing Annual Report 2025. 2025 preliminary.", INT, "data/raw/eia/eia_table23_inventories_by_owner.csv"))
    un = read_csv(d / "eia_table11_unfilled_requirements.csv")
    for fld, name in [("asof_2025_annual", "US unfilled requirements, annual, as of end 2025"), ("asof_2024_annual", "US unfilled requirements, annual, as of end 2024")]:
        cols.append(mk(name, {int(r["year"]): val(r[fld]) for r in un}, "thousand lb U3O8 equivalent", "Table 11. Reactor owners' requirements not yet under contract; forward-looking, as of the date named.", INT, "data/raw/eia/eia_table11_unfilled_requirements.csv"))
    mx = read_csv(d / "eia_table12_max_requirements_end2025.csv")
    cols.append(mk("US reactor requirements, maximum anticipated, as of end 2025", {int(r["year"]): val(r["max_anticipated_requirements"]) for r in mx}, "thousand lb U3O8 equivalent",
                   "Table 12. Maximum anticipated market requirements of US reactor owners, 2026-2035; forward-looking.", INT, "data/raw/eia/eia_table12_max_requirements_end2025.csv"))
    cd = read_csv(d / "eia_table10_contracted_deliveries_in_effect_end2025.csv")
    for fld, name in [("all_suppliers_minimum", "US contracted deliveries, minimum, in effect end 2025"), ("all_suppliers_maximum", "US contracted deliveries, maximum, in effect end 2025")]:
        cols.append(mk(name, {int(r["delivery_year"]): val(r[fld]) for r in cd}, "thousand lb U3O8 equivalent", "Table 10. Contracted purchases by US reactor owners, by delivery year; forward-looking; W = withheld (blank).", INT, "data/raw/eia/eia_table10_contracted_deliveries_in_effect_end2025.csv"))
    return [c for c in cols if c["values"]]


# ---------------------------------------------------------------- writing (layout lives in workbook_layout.py)
RENAME = [("Events_SPUT", "SPUT purchases"), ("Issuance_SPUT", "SPUT issuance"), ("Statements_SPUT", "SPUT statements"), ("UxC_Volume_YC", "UxC volumes"),
          ("Daily_OHLCV", "Prices OHLCV"), ("Monthly_Pink", "World Bank monthly"), ("Variable_Coverage", "Variables")]
DERIVED = "Derived, not computed here"
OVERRIDE = {  # id -> (status, where, note); the workbook holds the inputs of derived variables, not the result
    "V030": ("Partly in workbook", "SPUT statements: Pounds per unit, at reporting dates", "Weekly holdings and units outstanding are not on the same date, so no continuous series."),
    "V031": (DERIVED, "Inputs: Monthly, SPUT TSX volume (V014) and SPUT units outstanding (V021)", "Units outstanding are known at reporting dates and reconstructed month by month from Jun 2022 to Dec 2025."),
    "V032": ("Partly in workbook", "Quarterly: SPUT purchases as % of UxC spot volume; UxC volumes", "Denominator is the UxC spot volume quoted in Yellow Cake updates (quarterly, from 2021 Q3); it is not continuous and may include SPUT itself."),
    "V101": ("Partly in workbook", 'Monthly and Quarterly: UxC spot volume, first reported (Yellow Cake) and world spot volume as reported in Cameco MD&A; UxC volumes (every figure and vintage); Physical sources: Cameco spot volumes 2006-2026 incl. comparison figures, US utility spot-contract purchases and prices 2003-2025 (UMAR S1a/S1b), EU spot share 1980-2012 (Euratom)', 'No continuous history: Cameco gives 53 own-quarter or cumulative values (Q4 is never given on file), Yellow Cake from 2020. Cameco names UxC only from 2015; later reports revise earlier quarters (e.g. Q3 2025: 13, then 15). EIA and Euratom figures are US and EU utility purchases, not the world market.'),
    "V035": ("In workbook", "Yellow Cake reports: Shares in issue", "As printed in each report, including treasury shares."),
    "V036": ("Partly in workbook", "Yellow Cake reports: New shares issued; Gross capital raised", "Filled where a report states one placing; multiple placings are described in the note, not summed."),
    "V037": ("Partly in workbook", "Yellow Cake reports: Uranium purchased in period", "Only where the report states the quarter's purchases; elected but undelivered purchases left blank."),
    "V038": ("Partly in workbook", "Yellow Cake reports: of which from Kazatomprom", ""),
    "V039": ("Partly in workbook", "Yellow Cake reports: of which other", ""),
    "V041": ("Partly in workbook", "Yellow Cake reports: Cash; Cash and other net current assets", "Cash is stated separately only in some reports."),
    "V044": ("Partly in workbook", "Yellow Cake reports: Shares bought back; Buyback cost", "Only the periods with a buyback programme; early-2020 counts are cumulative and left blank."),
    "V093": ("Partly in workbook", 'Physical sources: WNA 2015-2024; Red Book 2018 edition (2014-2017), 2022 edition (2018-2021) and 2024 edition (2020-2023) production by country; Red Book world production by method 2013-2023 (per cent)', 'Tonnes U. WNA values re-read against the WNA page on 8 Nov 2026 (70 values, no differences); Red Book country rows sum to their totals except the 2018 edition cumulative column (+165). Editions overlap (2020 and 2021) and revise each other: both are kept, the note names the edition. The 2021 WNA tU3O8 total (46,374) is printed so at source: use tonnes U. WNA and Red Book differ for China, India, Ukraine, South Africa and the USA. Annual country production before 2014 is in no document on file (the 2018 Red Book gives only the pre-2014 cumulative).'),
    "V097": ("Partly in workbook", 'Physical sources: IAEA TECDOC-2030 global and regional inventories and 2021 demand at end 2021 (Tables 16-23: natural, enriched, fabricated, value, years of cover); Euratom EU utility stocks 2007-2012 (2010 revised from 45,272 to 45,354 tU)', 'No continuous global series exists; as-of dates, owners and material forms differ. TECDOC global figures are IAEA estimates for one date (end 2021).'),
    "V076": ('In workbook', 'Daily and Monthly: funding spread = commercial paper 90-day AA financial minus T-bill 3-month (1997 onwards); raw TED spread (to Jan 2022), both rates and SOFR also kept', 'Decision 9 Nov 2026: commercial paper minus T-bill is the TED replacement. Not spliced to the TED spread; construction differs (commercial paper vs LIBOR).'),
    "V082": ('In workbook', 'Monthly: OECD industrial production growth, year on year (to 2026-07), the main series; OECD and G7 levels (to 2024-03)', 'Decision 9 Nov 2026: the growth series is the main one. Checked against growth computed from the level series over 579 overlapping months: mean absolute difference 0.28 points, maximum 1.19 (different vintages). US industrial production is V081.'),
    "V083": ("In workbook", "Quarterly: real GDP, US, level and annualised growth", "US only; OECD and other national GDP not retrieved."),
    "V084": ("In workbook", "Monthly: Kilian global real economic activity index", "From the Dallas Fed file."),
    "V088": ("In workbook", "Monthly: categorical EPU, 11 categories", "No energy category in the categorical file."),
    "V089": ("In workbook", "Monthly and Daily: geopolitical risk index, overall", "Daily file covers calendar days; weekends are not shown on the Daily sheet."),
    "V090": ("In workbook", "Monthly and Daily: geopolitical risk, threats and acts", "Same note as V089."),
    "D21": ("In workbook", "Monthly: Google Trends 'uranium ETF'", "Relative interest, peak = 100; carries measurement error."),
    "D44": ("Partly in workbook", "UPC monthly NAV, Dec 2010 to May 2021 (126 months; May 2018 from the quarterly filings) nine 2009 months (Jan to Apr, Aug to Dec) and Jan and Feb 2010; Daily and Monthly: UPC adj. close, 2006 to Jul 2021", "May to Jul 2009, Mar to Nov 2010 and June 2021 not retrieved; shares outstanding not printed in the releases."),
    "V103": ("In workbook", "Annual: world nuclear electricity produced, IAEA PRIS, 1970-2025", "World total only; country detail sits in the PRIS Data Browser (read on screen, no export). Values match RDS-2 for the six years checked."),
    "V106": ("In workbook", "Annual: world nuclear construction starts, grid connections, permanent shutdowns and reactors in operation, IAEA RDS-2 (2026 edition), 1954-2025", "World totals per year; the reactor-level shutdown list (220 reactors) is in the raw CSV. Operating status is the year-end count of reactors in operation."),
    "V105": ("In workbook", "Annual: world nuclear energy availability, load, unit capability factors and forced loss rate, IAEA PRIS, 1970-2025", "World totals only; PRIS reports the energy availability factor (EAF) as its headline and the load factor (LF) as the realised measure. Six more indicators sit in the raw CSV."),
    "V104": ("In workbook", "Annual: world nuclear operated capacity, IAEA PRIS, 1970-2025", "World total only. PRIS operated capacity differs from the year-end net capacity printed in RDS-2."),
    "V092": ("Partly in workbook", 'Annual: US uranium concentrate production, 2011-2025; Physical sources: US summary 1993-2010 and 2011-2025 (mine production, concentrate production and shipments, drilling, expenditures, employment) and number of operating mines 2011-2025', 'US only. Mine production by method is withheld by EIA (W) for every year; 2020 mine and concentrate production are withheld. Global and country production is V093.'),
    "V095": ("Partly in workbook", "Annual: US reactor owners inventory (total and by material), year-end 2021-2025; Physical sources: 2011-2015 (UMAR 2015), 2020 and 2024 preliminary (UMAR 2024), 2016-2019 as rounded totals (Table S3a)", "US only. 2016-2019 exist only as totals rounded to 0.1 million lb (EIA does not give the detailed tables for those years). The 2024 preliminary value (126.4 million lb) was revised to 115.1 in the 2025 edition."),
    "V096": ("Partly in workbook", "Annual: US supplier inventory (uranium concentrate; UF6 withheld), year-end 2021-2025; Physical sources: 2011-2015, 2020 and 2024 preliminary, 2016-2019 as rounded totals", "Natural and enriched UF6 held by suppliers are withheld by EIA."),
    "V098": ("Partly in workbook", 'Annual: US reactor requirements, maximum anticipated, 2026-2035; Physical sources: Red Book 2018, 2022 and 2024 world requirements (actual 2016, 2020, 2022; projections to 2050); EU fuel loaded into reactors 1980-2012 (Euratom); 2021 world demand by region and form (IAEA TECDOC-2030); US maximum anticipated requirements as of end 2015 and end 2024 (UMAR)', 'US figures forward-looking (EIA). World Red Book figures are scenarios, not observations, and the editions revise each other (the 2035 low case is 53,010, 60,544 and 77,242 tU in the 2018, 2022 and 2024 editions). EU fuel loaded and the 2021 TECDOC demand are observed values.'),
    "V042": ("In workbook", "Quarterly: Yellow Cake NAV per share (with the quarter-end spot price used); Yellow Cake reports: total NAV and NAV per share as stated, per report", "Total NAV is the figure each of the 48 reports states (read from the PDFs), never shares times NAV per share. Three reports state a pro forma or later figure; the basis column says which."),
    "D17": ('In workbook', 'Monthly: S&P 500 (^GSPC, from 1988) and S&P 500 (FRED, last ten years)', 'Decision 9 Nov 2026: the S&P 500 is the global-equity proxy (US index; longest history). MSCI World is licensed and not retrieved.'),
    "V099": ("Partly in workbook", 'Annual: US unfilled requirements, as of end 2024 and end 2025; Physical sources: as of end 2014, 2015 (UMAR 2015) and end 2023, 2024 (UMAR 2024)', 'Forward-looking; four year-end vintages (2014, 2015, 2023, 2024) plus 2025. Years 2016-2022 are missing.'),
    "V100": ("Partly in workbook", 'Annual: US contracted deliveries, minimum and maximum, in effect end 2025; Physical sources: in effect end 2015 (UMAR 2015) and end 2024 (UMAR 2024)', 'Forward-looking; three year-end vintages (2015, 2024, 2025); US/foreign split withheld by EIA for later delivery years.'),
    "D18": ("In workbook", "Monthly: Baltic Dry Index, month-end, Jan 1985 to Oct 2026", "Month-end values of a free daily series; agrees with the Investing.com monthly table for Jun to Oct 2026. Daily history not stored."),
    "D47": ("Partly in workbook", "SPUT issuance (17 Aug to 18 Nov 2021 and 16 Feb 2022 to 19 Jan 2026, 466 dated rows); SPUT statements (period totals)", "Dated issuances: 17 Aug to 18 Nov 2021 and 16 Feb 2022 onwards. 19 Nov 2021 to 15 Feb 2022 is not listed in any retrieved document (913,400 units in 19 Nov to 31 Dec 2021 by difference, and about 18.6 million units in 1 Jan to 15 Feb 2022 by the reports). Sprott publishes only annual and June reports, so no quarterly flows exist."),
}

OVERRIDE.update({
    "V060": ("Partly in workbook", "Monthly: Newcastle coal futures close (Investing NCFMc1)", "Newcastle only, Jan 2009 to Oct 2026, monthly bar close. API2 Rotterdam not retrieved."),
    "V075": ("Partly in workbook", "Monthly: Monetary-policy surprise, raw and orthogonalized, sum over FOMC events in month", "San Francisco Fed series, Jan 2012 to Dec 2023 only (chart file); 2024 onward not retrieved."),
    "V118": ("Partly in workbook", "Daily: UX volume and open interest, CME bulletin (20 Mar to 9 Nov 2026); Monthly: UX volume and open interest, CME monthly report (volume 42 months from Jun 2014, open interest 30 months from Aug 2014 to May 2024)",
             "CME publishes no free history. The workbook holds only what archived CME files show: 124 trading days in 2026 (open interest 250 to 350 contracts, volume above zero on 6 days) and scattered months from 2014. Every other period is blank. The Stooq volume and open interest columns are not used. Full daily history is only in the paid CME DataMine product."),
})


# 19 Nov 2026: register fixes from the Data Curator's ML-readiness check (report of 19 Nov 2026, section 3), checked against workbook v4.
OVERRIDE.update({
    "V047": ("In workbook", "World Bank monthly: Gold (Pink Sheet monthly average, the column tagged V047). The gold month-end is Monthly: Gold futures (GC=F), month-end, tagged V059; there is no gold spot column", "Monthly average as published. Futures dates repaired 15 Nov 2026 (data/raw/markets_dated)"),
    "V051": ("In workbook", "Monthly: Natural gas, Henry Hub (IMF), monthly avg; World Bank monthly: Natural gas, US (both period averages, tagged V051). The Henry Hub spot month-end on Monthly is tagged V057", "Period averages as published. The spot month-end comes from FRED DHHNGSP (scripts/fetch_fred.py), see V057"),
    "V054": ("In workbook", "Monthly: Copper (IMF), monthly avg; World Bank monthly: Copper (both period averages, tagged V054). The copper month-end is Monthly: Copper futures (HG=F), month-end, tagged V059", "Period averages as published. Futures dates repaired 15 Nov 2026 (data/raw/markets_dated)"),
    "V058": ("In workbook", "Daily: Gold futures (GC=F) close and Copper futures (HG=F) close, tagged V059; there is no gold or copper spot column", "Yahoo continuous front-month futures (scripts/fetch_markets.py); the gold spot ticker XAUUSD=X is listed in the fetch script but no file for it is in data/raw/markets. Futures dates repaired 15 Nov 2026 (data/raw/markets_dated)"),
    "V073": ("Partly in workbook", "Daily and Monthly: Corporate option-adjusted spread (ICE BofA), from 7 Nov 2023 (Monthly 2023-11 to 2026-10, 36 months)", "From scripts/fetch_fred.py; FRED now serves only about 3 years of this ICE BofA series, so the earlier history is not in the workbook"),
    "V030": (DERIVED, "Inputs: Weekly and Monthly: SPUT holdings (V029); Monthly: SPUT units outstanding (V021). SPUT statements: Pounds per unit at reporting dates is a ratio built in this workbook (not in the raw file)", "Weekly holdings and units outstanding are not on the same date, so no continuous series. Whether the ratio column on SPUT statements stays is for Ronaldo (D-029)."),
    "V032": (DERIVED, "Inputs: SPUT purchases (V027) and UxC volumes (V101). Quarterly: SPUT purchases as % of UxC spot volume is a ratio built in this workbook (15 quarters)", "Denominator is the UxC spot volume quoted in Yellow Cake updates (quarterly, from 2021 Q3); it is not continuous and may include SPUT itself. Whether the ratio column stays is for Ronaldo (D-029)."),
    "V043": (DERIVED, "Quarterly: Yellow Cake premium/discount to NAV, quarter-end (a ratio built in this workbook from the YCA.L close and the announced NAV, 33 quarter-ends)", "quarter-end price vs quarter-end NAV. The column is built in this workbook although the variable is a derived one; whether it stays is for Ronaldo (D-029)."),
    "V108": (DERIVED, "", "needs production data (V093). Derived variable: not built in this workbook."),
    "V109": (DERIVED, "", "needs inventory and requirements data (V097, V098). Derived variable: not built in this workbook."),
    "D27": ("Not retrieved", "", "dates only; sanctions date not listed. Event dates have to be compiled from documents; they cannot be derived from workbook data."),
    "V026": ("Partly in workbook", "SPUT statements: cash; Monthly: SPUT cash, reported", "dated anchors only (Jun 2022, Jun 2023, Dec 2023, Jun 2024, Dec 2024, Jun 2025, Dec 2025, Jun 2026); daily cash is not published"),
    "V089": ("In workbook", "Monthly and Daily: geopolitical risk index, overall", "Daily file covers calendar days; the Daily sheet shows weekend values from 1990 (its date list has weekend rows from then); earlier weekend values are only in the raw file."),
    "V090": ("In workbook", "Monthly and Daily: geopolitical risk, threats and acts", "Daily files cover calendar days; the Daily sheet shows weekend values from 1990 (its date list has weekend rows from then); earlier weekend values are only in the raw file."),
})


def clean(t):
    t = t or ""
    for a, b in RENAME:
        t = t.replace(a, b)
    t = re.sub(r"\s*\((?:H[1-4])(?:[,/ ]+H?[1-4])*\)", "", t)
    return t.strip()


def coverage_by_id():
    by, extras = {}, []
    for (i, cat, var, variant, smst, src, st, where, note, req, step, purpose, issue, route) in COVERAGE:
        if req:
            st = ("Constructed" if st == "Constructed" else "In workbook") if (RAW / req).exists() else "Script written, not run"
        if st in ("Constructed", "Analysis stage (inputs in workbook)"):
            st, where, note = DERIVED, "", (note + " " if note else "") + "Derived variable: not built in this workbook."
        if i in OVERRIDE:
            st, where, note = OVERRIDE[i]
        if st == "In workbook":  # retrieved since the coverage map was written
            note = re.sub(r"\bneeds (scripts/fetch_\w+\.py)", r"from \1", note)
            note = re.sub(r"\bmonth-end needs ", "month-end from ", note)
        rec = (st, clean(where), clean(note))
        if i.startswith("V"):
            by[i] = rec
        else:
            extras.append((i, var, variant) + rec)
    assert len(by) == 119 and all(k in by for k in CHAT), "coverage map does not match the handoff table"
    return by, extras


def rng(keys):
    k0, k1 = min(keys), max(keys)
    f = lambda k: str(k) if isinstance(k, int) else (f"{k[0]}-{k[1]:02d}" if isinstance(k, tuple) else k.isoformat())
    return f"{f(k0)} to {f(k1)}"


def main():
    S = load()
    pk = pink()
    D, W, M, Q = S["daily"], S["weekly"], S["monthly"], S["quarterly"]

    spot = next(c for c in M if c["header"].startswith("Uranium spot, month-end"))["values"]
    months = list(months_between(min(spot), max(spot)))
    dates = set()
    for c in D:
        if c["sid"] not in ("DFF", "GPRD", "GPRD_THREAT", "GPRD_ACT"):
            dates |= set(c["values"])
    dates = sorted(dates)
    qkeys_all = [k for c in Q for k in c["values"] if c["freq"] == "quarterly"]
    quarters = quarters_between(min(qkeys_all), max(qkeys_all))
    wkeys = sorted(W[0]["values"])
    pink_months = list(months_between(min(min(c["values"]) for c in pk), max(max(c["values"]) for c in pk)))

    by, extras = coverage_by_id()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheets = []  # (group, name, description, frequency, covers, rows, columns)

    # ---- frequency sheets
    def freq(name, title, sub, keys, cols, keyhdr, keyfmt, keyval, desc, fr, **kw):
        ws = wb_layout.write_data_sheet(wb, name, title, sub, keys, cols, keyhdr, keyfmt, keyval, **kw)
        sheets.append(("Data by publication frequency", name, desc, fr, rng(keys), len(keys), len(cols)))
        return ws

    mk = lambda k: dt.date(k[0], k[1], 1)
    # 19 Nov 2026: timing codes and release-date labels (rows Timing and Release, data from row 13); the Monthly date header says the 1st is shown (Curator 7.11).
    wb_layout.load_release(ROOT / "data" / "release_dates.csv")
    wsm = freq("Monthly", "Monthly series", "One row per calendar month. Prices and indicators are month-end values; period averages appear only where no daily source exists (type colour: grey).",
         months, M, "Month (date shown is the 1st; values at month-end or for the month)", "yyyy-mm", mk, "Uranium prices and the month-end value of every daily series, plus monthly-published series.", "Monthly", width=17)
    wsq = freq("Quarterly", "Quarterly series", "One row per quarter-end. Series published once a quarter, plus the quarter-end value of monthly series.",
         quarters, Q, "Quarter end", "yyyy-mm-dd", lambda k: k, "Conversion prices, Yellow Cake NAV and holdings, UxC spot volume and SPUT purchases per quarter.", "Quarterly", width=18)
    # visible flag on the pro forma holdings cell (31 Mar 2021)
    from openpyxl.comments import Comment
    jh = [j for j in range(2, wsq.max_column + 1) if "uranium held" in str(wsq.cell(7, j).value or "").lower() and "corrected" not in str(wsq.cell(7, j).value or "").lower() and "exact" not in str(wsq.cell(7, j).value or "").lower()]  # 19 Nov 2026: skip the new exact-lb column
    rh = [r for r in range(wb_layout.DATA_ROW, wsq.max_row + 1) if (wsq.cell(r, 1).value.date() if hasattr(wsq.cell(r, 1).value, "date") else wsq.cell(r, 1).value) == dt.date(2021, 3, 31)]  # 19 Nov 2026: data start below the new Timing and Release rows
    if len(jh) != 1 or len(rh) != 1:
        raise SystemExit(f"cannot find the Yellow Cake holdings cell for 31 Mar 2021 on the Quarterly sheet: columns {jh}, rows {rh}")
    cell_ = wsq.cell(rh[0], jh[0])
    cell_.fill = PatternFill("solid", fgColor=FILL["gap-fill"])
    cell_.comment = Comment("PRO FORMA, not pounds held. Actually held at 31 Mar 2021: 9,856,385 lb (grade A; see the next column). Left as published. Source: data/raw/yellowcake/yellowcake_holdings_correction_2021-03-31.csv", "Dataset", width=340, height=110)
    # 19 Nov 2026: cell comments on the Dec 2017 Cameco cells, the Yellow Cake stated-holdings cells and the empty Oct 2025 CPI cell, so a reader sees why a cell differs from the raw file or is empty.
    def cell_at(ws_, key, start):
        jj = [j for j in range(2, ws_.max_column + 1) if str(ws_.cell(7, j).value or "") == start]
        rr = [r for r in range(wb_layout.DATA_ROW, ws_.max_row + 1) if (ws_.cell(r, 1).value.date() if hasattr(ws_.cell(r, 1).value, "date") else ws_.cell(r, 1).value) == key]  # 19 Nov 2026: data start below the new header rows
        if len(jj) != 1 or len(rr) != 1:
            raise SystemExit(f"cannot find the cell {start} / {key} on {ws_.title}: columns {jj}, rows {rr}")
        return ws_.cell(rr[0], jj[0])
    for (d_, k_), r_ in S["cam_fix"].items():
        txt_ = (f"CORRECTED (19 Nov 2026, grade A). Raw file value: {r_['raw_value']} (Cameco's Q4 2017 average in the December slot). Shown here: {r_['corrected_value']}, the month-end printed in "
                f"{r_['source_document']} ({r_['page_or_table']}): \"{r_['quote']}\". {r_['note']} Correction file: data/raw/cameco/uranium_price_monthly_correction_2017-12.csv")
        nm_ = "Uranium spot" if k_ == "spot_usd_lb" else "Uranium long-term"
        for ws_, key_, hd_ in [(wsm, day(d_).replace(day=1), f"{nm_}, month-end"), (wsq, day(d_), f"{nm_}, quarter-end (Cameco)")]:
            c_ = cell_at(ws_, key_, hd_)
            if c_.value != num(r_["corrected_value"]):
                raise SystemExit(f"{ws_.title} {hd_} {key_} holds {c_.value}, not the corrected {r_['corrected_value']}")
            c_.comment = Comment(txt_, "Dataset", width=380, height=170)
    for q_, r_ in S["yc_stated"].items():
        src_ = f"{r_['source_document']}, p{r_['page']}: \"{r_['quote']}\" (grade {r_['grade']}, read {r_['read_on']})" + (f". {r_['note']}" if r_["note"] else "")
        cell_at(wsq, q_, "Yellow Cake lb basis").comment = Comment("Stated in the kept update (nav_quarterly.csv says 'derived'; raw file unchanged). " + src_, "Dataset", width=380, height=150)
        cell_at(wsq, q_, "Yellow Cake uranium held, exact as stated").comment = Comment(src_, "Dataset", width=380, height=150)
    cpi_ = cell_at(wsm, dt.date(2025, 10, 1), "Consumer price index (CPI-U, seasonally adjusted)")
    if cpi_.value is not None:
        raise SystemExit("CPI 2025-10 is no longer empty: check the comment text")
    cpi_.comment = Comment("Empty at the source: the US BLS did not collect CPI data from 1 Oct to 12 Nov 2025 because of the 2025 federal government shutdown, and FRED CPIAUCSL has no value for 2025-10 (raw file blank). "
                           "Left empty, never filled (D-005). BLS page https://www.bls.gov/cpi/additional-resources/2025-federal-government-shutdown-impact-cpi.htm (read 19 Nov 2026)", "Dataset", width=380, height=130)
    freq("Weekly", "Weekly series", "One row per Friday.", wkeys, W, "Date (Friday)", "yyyy-mm-dd", lambda k: k, "SPUT holdings from the Sprott fund page.", "Weekly", width=18)
    freq("Daily", "Daily series", "One row per trading day. Blank cell = market closed or no observation; nothing is carried forward.",
         dates, D, "Date", "yyyy-mm-dd", lambda k: k, "Market, rate, risk and fund series at their daily frequency.", "Daily", width=16)

    ann = annual_eia()
    pf = RAW / "iaea_pris" / "iaea_pris_world_annual_1970_2025.csv"
    if pf.exists():
        pr = read_csv(pf)
        pris_note = ("PRIS Analytics Data Browser, Operational > Electricity production, all countries; read from the on-screen table on 8 Nov 2026 (the page has no export). "
                     "Generation matches the IAEA RDS-2 (2024 edition) world totals for 1995, 2000, 2005, 2010, 2015 and 2020; later years differ from that report where PRIS has revised them (2022: 2,444.9 against 2,486.8 TWh)")
        ann.append(col("World nuclear electricity produced, IAEA PRIS [MWh]", {int(r["year"]): num(r["electricity_produced_mwh"]) for r in pr}, "native", "IAEA PRIS", "PRIS electricity produced", "annual",
                       "calendar year, as published", "MWh", pris_note, "", "data/raw/iaea_pris/iaea_pris_world_annual_1970_2025.csv", INT))
        ann.append(col("World nuclear operated capacity, IAEA PRIS [MW(e)]", {int(r["year"]): num(r["operated_capacity_mwe"]) for r in pr}, "native", "IAEA PRIS", "PRIS operated capacity", "annual",
                       "calendar year, as published", "MW(e)", "Net capacity of the reactors PRIS counts as operated in the year; it differs from the year-end net capacity in RDS-2 (e.g. 2015: 364,367 against 360,504 MW(e)). " + pris_note, "", "data/raw/iaea_pris/iaea_pris_world_annual_1970_2025.csv", INT))
    pp = RAW / "iaea_pris" / "iaea_pris_world_performance_1970_2025.csv"
    if pp.exists():
        pc = read_csv(pp)
        perf_note = ("PRIS Analytics Data Browser, Performance indicators, table view, all countries; read from the on-screen table on 9 Nov 2026 (no export). Annual world value as published by PRIS. "
                     "Checks: EAF + EUF = 100, UCF = 100 - PUF - UCL and EUF = PUF + UCL + XUF hold within rounding for all 56 years. 1970 has too few reactors to be meaningful (operating factor 31.5% below load factor 59.7%, as published). "
                     "The other PRIS indicators (EUF, XUF, OF, PUF, UCL, UUF) are in data/raw/iaea_pris/iaea_pris_world_performance_1970_2025.csv")
        for field, hdr, what in [("energy_availability_factor_pct", "World nuclear energy availability factor, IAEA PRIS [%]", "Energy available as a share of reference energy (PRIS EAF, the capacity-factor measure PRIS now headlines)"),
                                 ("load_factor_pct", "World nuclear load factor, IAEA PRIS [%]", "Energy produced as a share of reference energy (PRIS LF); closest to a realised capacity factor"),
                                 ("unit_capability_factor_pct", "World nuclear unit capability factor, IAEA PRIS [%]", "Energy available excluding external losses (PRIS UCF)"),
                                 ("forced_loss_rate_pct", "World nuclear forced loss rate, IAEA PRIS [%]", "Share of energy lost to unplanned outages (PRIS FLR)")]:
            ann.append(col(hdr, {int(r["year"]): num(r[field]) for r in pc}, "native", "IAEA PRIS", field, "annual", "calendar year, as published", "%", what + ". " + perf_note, "",
                           "data/raw/iaea_pris/iaea_pris_world_performance_1970_2025.csv", NUM))
    ep = RAW / "iaea_pris" / "iaea_rds2_2026_world_reactor_events_1954_2025.csv"
    if ep.exists():
        ec = read_csv(ep)
        ev_note = ("IAEA Reference Data Series No. 2 (RDS-2), 46th edition, 2026: Table 7 (construction starts, grid connections, reactors in operation at year-end; a blank in the report is entered as 0) and, for shutdowns, Table 16 (220 reactors permanently shut down by 31 Dec 2025, counted by year of shutdown; "
                   "their net capacity adds to 110,812 MW(e), the figure the report prints). Typed from the PDF text by script. Operating count does not equal the previous year plus connections minus shutdowns in 18 years, because PRIS also moves reactors in and out of suspended operation (e.g. Japan after 2011); the raw file keeps all columns. "
                   "Reactor-level shutdown list: data/raw/iaea_pris/iaea_rds2_2026_permanently_shut_down_reactors.csv")
        for field, hdr, unit in [("construction_starts_units", "World nuclear construction starts, IAEA RDS-2 [units]", "units"),
                                 ("grid_connections_units", "World nuclear grid connections, IAEA RDS-2 [units]", "units"),
                                 ("grid_connections_mwe", "World nuclear grid connections, IAEA RDS-2 [MW(e)]", "MW(e)"),
                                 ("permanent_shutdowns_units", "World nuclear permanent shutdowns, IAEA RDS-2 [units]", "units"),
                                 ("permanent_shutdowns_net_mwe", "World nuclear permanent shutdowns, IAEA RDS-2 [MW(e)]", "MW(e)"),
                                 ("reactors_in_operation_units", "World nuclear reactors in operation, year-end, IAEA RDS-2 [units]", "units")]:
            ann.append(col(hdr, {int(r["year"]): num(r[field]) for r in ec}, "native", "IAEA RDS-2 (2026 edition)", field, "annual", "calendar year, as published", unit, ev_note, "",
                           "data/raw/iaea_pris/iaea_rds2_2026_world_reactor_events_1954_2025.csv", INT))
    if ann:
        ayears = list(range(min(min(c["values"]) for c in ann), max(max(c["values"]) for c in ann) + 1))
        freq("Annual", "Annual series", "One row per year. Realised values come first; 2026-2035 columns are the reactor owners' own forward-looking figures from the EIA survey, as of the date in the column name.",
             ayears, ann, "Year", "0", lambda k: k, "US production, inventories, requirements and contracted deliveries (EIA).", "Annual", width=18)

    # ---- record tables
    def table(name, title, sub, headers, rows, widths, desc, fr, covers, **kw):
        wb_layout.write_table_sheet(wb, name, title, sub, headers, rows, widths, **kw)
        sheets.append(("Event and document tables", name, desc, fr, covers, len(rows), len(headers)))

    ev = sorted(S["events"])
    table("SPUT purchases", "SPUT purchases of physical uranium", "V027 (pounds bought) and V029 (holdings). One row per purchase day, from the Sprott fund page.",
          ["Purchase date", "Pounds acquired [lb U3O8]", "Cumulative pounds acquired [lb U3O8]"], [list(r) for r in ev], [16, 26, 32], "Every SPUT purchase day with pounds acquired.", "Irregular",
          f"{ev[0][0]:%Y-%m-%d} to {ev[-1][0]:%Y-%m-%d}", fmts=["yyyy-mm-dd", INT, INT])
    if "issuances" in S:
        FLAG_UW = "Underwritten deal priced against the 13 Jun 2025 NAV; not an ATM sale"
        cum, rows = 0, []
        for d_, pr_, n_, t_ in S["issuances"]:
            cum += n_
            flag_ = FLAG_UW if (str(d_)[:10] == "2025-06-20" and n_ == 11600000 and abs(pr_ - 17.25) < 1e-9) else None
            rows.append([d_, pr_, n_, pr_ * n_, t_, cum, S["issuance_src"][(d_, pr_, n_)], flag_])
        assert sum(1 for r_ in rows if r_[7]) == 1, "20 Jun 2025 underwritten row not found exactly once"
        table("SPUT issuance", "SPUT unit issuance", "V022 (new units), V023 (issue price) and V024 (gross proceeds). Dates from the Prior Sales tables of the prospectus supplements of 22 Nov 2021 (17 Aug to 18 Nov 2021), 15 Feb 2023, 3 Jan 2024, Dec 2024 and 18 Jun 2025 and the 22 Jan 2026 base shelf prospectus (16 Feb 2022 to 19 Jan 2026, no gap); gross value = price × units. Issues from 19 Nov 2021 to 15 Feb 2022 are not listed. Jul to Dec 2023 is 33,603 units short of the 2023 report: that amount has no dates and is not spread over any day. By the daily-file test, 44 issuance rows are below the previous day's NAV (41 before 2026 including 20 Jun 2025, so 40 excluding it; 3 in Jan 2026). One of them, 20 Jun 2025 (11,600,000 units at 17.25 against a prior-day NAV of 18.33, 5.89%), is a timing effect: it is an underwritten bought deal priced on 13 Jun 2025 against that day's NAV of 17.0376 (net to the Trust 17.0376), not an ATM sale (see the flag column). The other 43 are shortfalls of 0.007% to 0.73%: 18 are shallower than 0.15% and 25 are between 0.15% and 0.75%; the deepest are 10 and 12 Apr 2024. The reason for those 43 is not in the documents. The daily NAV is rounded to the cent (five rows are within half a cent of the prior NAV), and the 216 days on which the NAV is unchanged may make the prior NAV stale. The 2021 table plus the 75,210,456 arrangement units equals the 162,558,000 units stated on 22 Nov 2021.",
              ["Issue date", "Price per unit [USD]", "Units issued", "Gross value, price × units [USD]", "Issue type", "Cumulative units in this list", "Source document", "Flag, 15 Nov 2026"], rows, [14, 18, 16, 26, 44, 22, 52, 60],
              "Dated unit issues of SPUT (private placement, bought deal, at-the-market).", "Irregular", f"{rows[0][0]:%Y-%m-%d} to {rows[-1][0]:%Y-%m-%d}",
              fmts=["yyyy-mm-dd", "0.0000", INT, INT, None, INT, None, None])
        heads = ["Period end", "Period type", "Units outstanding", "Total NAV [USD thousand]", "NAV per unit [USD]", "TSX closing price, U.U [USD]", "Cash [USD thousand]",
                 "Uranium oxide at fair value [USD thousand]", "Pounds U3O8 held", "Units issued in the period", "Proceeds from issuance [USD thousand]", "Issue costs, cash flow [USD thousand]",
                 "Purchases of uranium oxide [USD thousand]", "Source", "Note", "Net proceeds, proceeds minus costs [USD thousand]", "Average issue price [USD]", "Pounds per unit [lb]",
                 "Uranium oxide fair value per lb [USD]", "Cash as % of NAV"]
        fields = ["period_end", "period_type", "units_outstanding", "nav_usd_thousand", "nav_per_unit_usd", "tsx_close_usd", "cash_usd_thousand", "uranium_oxide_fv_usd_thousand",
                  "pounds_u3o8_held", "units_issued_in_period", "proceeds_from_issuance_usd_thousand", "issue_costs_cashflow_usd_thousand", "purchases_of_uranium_usd_thousand"]
        rows = []
        unit_flags = {}
        uf = RAW / "sprott" / "sput_unit_count_note_2026-01.csv"
        if uf.exists():
            unit_flags = {x["date"]: x for x in read_csv(uf)}
        # 15 Nov 2026: cash, flows, fair value, pounds and units that the kept annual reports print but sput_statements.csv leaves blank.
        # They are shown in extra columns at the right; the original cells are never overwritten.
        CORR = [("cash_usd_thousand", "Cash, from the kept report [USD thousand]", INT), ("uranium_oxide_fv_usd_thousand", "Uranium oxide at fair value, from the kept report [USD thousand]", INT),
                ("pounds_u3o8_held", "Pounds U3O8 held, from the kept report", INT), ("units_issued_in_period", "Units issued for cash, from the kept report", INT),
                ("proceeds_from_issuance_usd_thousand", "Proceeds from issuance, from the kept report [USD thousand]", INT), ("issue_costs_cashflow_usd_thousand", "Issue costs, from the kept report [USD thousand]", INT),
                ("purchases_of_uranium_usd_thousand", "Purchases of uranium oxide, from the kept report [USD thousand]", INT)]
        cfix = {}
        cf = RAW / "sprott" / "sput_statements_cash_correction_2026-11-15.csv"
        for x in (read_csv(cf) if cf.exists() else []):
            cfix.setdefault(x["period_end"], {})[x["field"]] = x
        for r in S["statements"]:
            v = [num(r[f]) if f not in ("period_end", "period_type") else r[f] for f in fields]
            v[0] = day(v[0])
            units_, nav_, cash_, fv_, lb_, iss_, pr_, co_ = v[2], v[3], v[6], v[7], v[8], v[9], v[10], v[11]
            uflag_ = unit_flags.get(r["period_end"])
            note_ = r["note"] + ((" | FLAG (grade B): " + f"units {uflag_['value_in_original_file']} as printed in the base prospectus (p36); the supplement prints {uflag_['corrected_value']} for 22 Jan 2026 and the 22 issuance rows of 2 to 19 Jan 2026 add up to exactly that figure, so this value is probably a transposed digit. Left unchanged (2026 is parked). Source: data/raw/sprott/sput_unit_count_note_2026-01.csv") if uflag_ else "")
            fx_ = cfix.get(r["period_end"], {})
            cells_, pages_ = [], []
            for f_, _h, _fm in CORR:
                if f_ in fx_:
                    orig_ = r[f_].strip()
                    if orig_ != "" and num(orig_) != num(fx_[f_]["corrected_value"]):
                        raise SystemExit(f"SPUT cash correction for {r['period_end']} {f_} disagrees with a non-blank value in sput_statements.csv")
                    cells_.append(num(fx_[f_]["corrected_value"]))
                    pages_.append(f"{f_.split('_')[0] if f_ != 'uranium_oxide_fv_usd_thousand' else 'fair value'} {fx_[f_]['page']}")
                else:
                    cells_.append(None)
            if fx_:
                docs_ = sorted({x["source_document"] for x in fx_.values()})
                note_ += (" | CORRECTED VALUES (15 Nov 2026, grade A) are in the last columns, read from " + ", ".join(docs_) + " (03_data/source_documents/sput_reports): "
                          + "; ".join(f"{x['field']} {x['page']}" for x in fx_.values()) + ". The original cells stay as they were. Source: data/raw/sprott/sput_statements_cash_correction_2026-11-15.csv")
                if r["period_end"] == "2021-12-31":
                    note_ += ". Units: 163,471,400 in total = 75,210,456 issued on acquiring UPC + 88,260,944 sold for cash (only these earned the proceeds); the cash-flow statement covers 23 Apr to 31 Dec 2021"
            rows.append(v + [r["source"], note_, pr_ - co_ if pr_ is not None and co_ is not None else None, pr_ * 1000 / iss_ if pr_ is not None and iss_ else None,
                             lb_ / units_ if lb_ and units_ else None, fv_ * 1000 / lb_ if fv_ and lb_ else None, cash_ / nav_ * 100 if cash_ is not None and nav_ else None] + cells_)
        table("SPUT statements", "SPUT statements at reporting dates",
              "V019 (total NAV), V021 (units), V024 (proceeds), V026 (cash), V027 (purchases) and V029 (pounds held). Annual and interim reports and the prospectus; the five columns after Note are ratios of the others. The seven last columns show values that the kept annual reports print for 31 Dec 2021, 31 Dec 2022 and 31 Dec 2023 where the original cells are blank (15 Nov 2026); the original cells are not overwritten. Cash at 20 Jan and 31 Mar 2026 stays blank (2026 is parked).",
              heads + [h_ for _f, h_, _m in CORR], rows, [13, 14, 16, 16, 14, 16, 14, 18, 16, 16, 18, 18, 18, 36, 46, 18, 14, 14, 16, 12] + [16] * 7,
              "SPUT balance sheet and cash flow at each annual, interim and prospectus date.", "Reporting dates", f"{rows[0][0]:%Y-%m-%d} to {rows[-1][0]:%Y-%m-%d}",
              fmts=["yyyy-mm-dd", None, INT, INT, "0.0000", NUM, INT, INT, INT, INT, INT, INT, INT, None, None, INT, NUM, "0.0000", NUM, NUM] + [m_ for _f, _h, m_ in CORR], wrap=(14, 15))
    if "uxc_rows" in S:
        ur = sorted(S["uxc_rows"], key=lambda r: (r["period_start"], r["period_end"], r["announcement_date"]))
        rows = [[day(r["announcement_date"]), day(r["document_quarter_end"]), r["period_label"], r["period_type"], day(r["period_start"]), day(r["period_end"]), float(r["volume_mlb"]),
                 r["material"], r["provider_in_sentence"], r["source_footnote"], r["flag"]] for r in ur]
        table("UxC volumes", "UxC spot-market volumes quoted by Yellow Cake", "V101 (spot transaction volume). Every figure in every Yellow Cake update, with its period and the update that quoted it; later updates may revise earlier figures.",
              ["Announcement date", "Update covers quarter ending", "Period", "Period type", "Period start", "Period end", "Volume [million lb]", "Material", "UxC named in the sentence", "Footnote / source cited", "Flag"],
              rows, [14, 16, 26, 14, 12, 12, 12, 10, 14, 44, 60], "Every UxC spot-volume figure quoted in the Yellow Cake updates.", "Irregular",
              f"{rows[0][4]:%Y-%m-%d} to {max(r[5] for r in rows):%Y-%m-%d}", fmts=["yyyy-mm-dd", "yyyy-mm-dd", None, None, "yyyy-mm-dd", "yyyy-mm-dd", NUM, None, None, None, None], wrap=(10, 11))


    # ---- Yellow Cake reports and physical-market sources (added 8 Nov 2026)
    yc = RAW / "yellowcake" / "yellowcake_report_series_verified.csv"
    if yc.exists():
        r_ = read_csv(yc)
        heads_ = ["Announcement date", "Period end", "Shares in issue", "New shares issued in period", "Gross capital raised [USD million]", "Uranium purchased in period [lb]",
                  "of which from Kazatomprom [lb]", "of which other [lb]", "Cash [USD million]", "Cash and other net current assets [USD million]", "Shares bought back in period",
                  "Buyback cost [USD million]", "Source", "Note", "Read from the PDF"]
        keys_ = ["announcement_date", "period_end", "shares_outstanding", "new_shares_issued_in_period", "gross_capital_raised_usd_m", "uranium_purchased_in_period_lb_total",
                 "purchased_from_kazatomprom_lb", "purchased_other_lb", "cash_usd_m", "net_current_assets_usd_m", "shares_bought_back_in_period", "buyback_cost_usd_m", "source_url", "note", "verified_from_pdf"]
        def cell(k, v):
            if v in ("", None):
                return None
            if k in ("announcement_date", "period_end"):
                return day(v)
            if k in ("source_url", "note", "verified_from_pdf"):
                return v
            return num(v)
        rows_ = sorted([[cell(k, x[k]) for k in keys_] for x in r_], key=lambda z: (z[0] is None, z[0]), reverse=True)
        # 15 Nov 2026 corrections and flags. The raw files are not changed; the Note column says what was done.
        def row_of(ann, per):
            hit = [z for z in rows_ if z[0] == day(ann) and z[1] == day(per)]
            if len(hit) != 1:
                raise SystemExit(f"Yellow Cake reports: expected one row for {ann} / {per}, found {len(hit)}")
            return hit[0]
        pf = RAW / "yellowcake" / "yellowcake_purchase_correction_2024-06.csv"
        for fx in (read_csv(pf) if pf.exists() else []):
            z = row_of(fx["announcement_date"], fx["period_end"])
            old_, new_ = num(fx["value_in_original_file"]), num(fx["corrected_value"])
            if z[5] != old_ or z[6] != old_:
                raise SystemExit("Yellow Cake purchase correction: the file value no longer matches the raw report row")
            z[5] = z[6] = new_
            z[13] = (z[13] or "") + (f" | CORRECTED in this workbook (15 Nov 2026): the update prints only '1.53 million lb'; the exact {new_:,.0f} lb (delivered 3 Jun 2024, US$65.50/lb, US$100.0m) is printed in the interim report to 30 Sep 2024 (p12, p23), "
                                     "the annual results to 31 Mar 2025 (p22) and the annual results to 31 Mar 2024 (p22, p24, p29). Columns 'Uranium purchased in period' and 'of which from Kazatomprom' show the exact figure; the raw file still holds 1,530,000. Source: data/raw/yellowcake/yellowcake_purchase_correction_2024-06.csv")
        hf = RAW / "yellowcake" / "yellowcake_holdings_correction_2021-03-31.csv"
        for fx in (read_csv(hf) if hf.exists() else []):
            z = row_of(fx["announcement_date"], fx["quarter_end"])
            z[13] = (z[13] or "") + (" | FLAG (15 Nov 2026, grade A): the holdings figure of 12.96 million lb in this update (used on the Quarterly sheet) is pro forma. Actually held at 31 Mar 2021: 9,856,385 lb "
                                     "(update of 20 Apr 2021 p1 fn 2 and p5 fn 17). Not a column on this sheet; the Quarterly sheet shows both. Source: data/raw/yellowcake/yellowcake_holdings_correction_2021-03-31.csv")
        sf = RAW / "yellowcake" / "yellowcake_sale_uranium_royalty_2021-04.csv"
        for fx in (read_csv(sf) if sf.exists() else []):
            z = row_of("2021-08-02", "2021-06-30")
            z[13] = (z[13] or "") + (f" | NOTE (15 Nov 2026): the sale of {-num(fx['quantity_lb']):,.0f} lb to Uranium Royalty Corp. (title moved {fx['event_date']}, US${num(fx['price_usd_per_lb']):.2f}/lb, US$10.0m) is a sale, not a purchase, "
                                     "so it has no column on this sheet. It explains the 30 Jun 2021 holdings of 13,305,601 lb: 9,856,385 - 348,068 + 343,053 + 3,454,231. Source: data/raw/yellowcake/yellowcake_sale_uranium_royalty_2021-04.csv")
        # total NAV as stated in the same report (read from the PDFs 8 Nov 2026); inserted after "Cash and other net current assets"
        navf = RAW / "yellowcake" / "yellowcake_total_nav.csv"
        nav_ = {(day(x["announcement_date"]), day(x["period_end"])): x for x in read_csv(navf)} if navf.exists() else {}
        for z in rows_:
            x = nav_.get((z[0], z[1]))
            z[10:10] = [num(x["total_nav_usd_m"]) if x else None, num(x["nav_per_share_gbp"]) if x else None, (x["basis"] or None) if x else None]
        heads_[10:10] = ["Total NAV, as stated [USD million]", "NAV per share, as stated [GBP]", "NAV basis"]
        table("Yellow Cake reports", "Yellow Cake plc: facts from its reports", "V035-V044. One row per quarterly update, interim report and annual result, read from the company PDFs. Blank = not stated in that report. Period figures are never summed across lots. Total NAV is the figure the report states, never shares times NAV per share. Flags in the Note column (15 Nov 2026): the 31 Mar 2021 holdings of 12.96 million lb are pro forma (9,856,385 lb held); the 348,068 lb sale to Uranium Royalty Corp. (28 Apr 2021) is not a purchase and has no column; the Jun 2024 purchase shows the exact 1,526,717 lb.",
              heads_, rows_, [14, 12, 16, 16, 16, 16, 16, 14, 12, 18, 16, 14, 30, 14, 14, 40, 70, 10], "Shares, capital raised, purchases, cash, total NAV and buybacks from every Yellow Cake report.", "Quarterly / half-year / annual",
              f"{min(z[1] for z in rows_ if z[1]):%Y-%m-%d} to {max(z[1] for z in rows_ if z[1]):%Y-%m-%d}",
              fmts=["yyyy-mm-dd", "yyyy-mm-dd", INT, INT, NUM, INT, INT, INT, NUM, NUM, NUM, NUM, None, INT, NUM, None, None, None], wrap=(13, 17))

    up = RAW / "upc" / "upc_monthly_nav.csv"
    if up.exists():
        r_ = read_csv(up)
        sup = RAW / "upc" / "upc_monthly_nav_2018-05_supplement.csv"
        if sup.exists():  # May 2018 has no monthly release; the figures come from the quarterly filings
            for x in read_csv(sup):
                x["source_url"] = ("SEDAR+ (profile 000021970): UPC interim MD&A for the quarter ended 31 May 2018, filed 2018-06-27, and interim financial statements, filed 2018-06-28. "
                                   "Not a monthly press release; see data/source_manifest.csv")
                pos = next((i for i, y in enumerate(r_) if y["as_of_date"] > x["as_of_date"]), len(r_))
                r_.insert(pos, x)
            if [y["as_of_date"] for y in r_] != sorted(y["as_of_date"] for y in r_):
                raise SystemExit("UPC monthly NAV is not in date order after adding May 2018")
        # 19 Nov 2026: Jan 2010 (monthly release reproduced in the kept SEC filing 10015870.pdf, p109) and Feb 2010 (year-end annual report in the same PDF, pp 2 and 8).
        s10 = RAW / "upc" / "upc_monthly_nav_2010_supplement_2026-11-19.csv"
        if s10.exists():
            for x in read_csv(s10):
                if any(y["as_of_date"] == x["as_of_date"] for y in r_):
                    raise SystemExit(f"UPC 2010 supplement: {x['as_of_date']} already in upc_monthly_nav.csv")
                x["source_url"] = "Kept PDF in 03_data/source_documents/upc: 10015870.pdf (sec.gov Archives/edgar/vprr/1001/10015870.pdf), " + ("page 109" if x["as_of_date"] == "2010-01-31" else "pages 2 and 8 (annual report, not a monthly release)")
                pos = next((i for i, y in enumerate(r_) if y["as_of_date"] > x["as_of_date"]), len(r_))
                r_.insert(pos, x)
            if [y["as_of_date"] for y in r_] != sorted(y["as_of_date"] for y in r_):
                raise SystemExit("UPC monthly NAV is not in date order after adding Jan and Feb 2010")
        # 2009 months (SEC filings, kept PDFs): the raw file is read as it is; the dated correction file supplies the page of each kept PDF,
        # the Dec 2009 NAV per share and two note corrections. A row without a page reference is not added.
        s09 = RAW / "upc" / "upc_monthly_nav_2009_2010.csv"
        c09 = RAW / "upc" / "upc_monthly_nav_2009_correction_2026-11-15.csv"
        if s09.exists():
            fixes = {}
            for f_ in (read_csv(c09) if c09.exists() else []):
                fixes.setdefault(f_["as_of_date"], {})[f_["field"]] = f_
            new09 = []
            for x in read_csv(s09):
                fx = fixes.get(x["as_of_date"], {})
                if "source_page" not in fx:
                    raise SystemExit(f"UPC 2009 row {x['as_of_date']} has no kept-PDF page reference in the correction file: not added")
                if "nav_per_share_cad" in fx:
                    if x["nav_per_share_cad"] != "":
                        raise SystemExit("UPC 2009 correction for NAV per share expects a blank in the raw row")
                    x["nav_per_share_cad"] = fx["nav_per_share_cad"]["corrected_value"]
                    x["note"] = x["note"].replace(" NAV per share NOT VERIFIED: printed 'CDN$6.1 per share' (truncated by OCR); left blank.", "")
                note_ = "Checked against the kept PDF page scan by the Data Curator on 15 Nov 2026 (grade A). " + x["note"].split(" Read from the OCR text")[0]
                if "note" in fx:
                    note_ += " Correction of the note above: " + fx["note"]["corrected_value"]
                if "nav_per_share_cad" in fx:
                    note_ += " NAV per share 6.17 comes from the correction file (the OCR text had a truncated '6.1')."
                sp = fx["source_page"]
                x["note"] = note_
                x["source_url"] = f"Kept PDF in 03_data/source_documents/upc: {sp['corrected_value']}, page {sp['page']}. Originally {x['source_url']}"
                new09.append(x)
            r_[0:0] = new09
            if [y["as_of_date"] for y in r_] != sorted(y["as_of_date"] for y in r_):
                raise SystemExit("UPC monthly NAV is not in date order after adding the 2009 months")
        rows_ = [[day(x["release_date"]), day(x["as_of_date"]), num(x["nav_cad_thousand"]), num(x["nav_per_share_cad"]), num(x["shares_outstanding"]), num(x["u3o8_lb"]), x["uf6_kgu_or_lb_and_unit"],
                  num(x["uranium_price_used_usd"]), x["source_url"], x["note"]] for x in r_]
        table("UPC monthly NAV", "Uranium Participation Corporation: monthly net asset value", "D44. The fund that SPUT replaced (Sprott took it over in July 2021). Monthly releases, December 2010 to May 2021 (May 2018 has no release: it comes from the quarterly MD&A and financial statements, see its Source column), plus nine 2009 months from SEC filings (Jan to Apr and Aug to Dec 2009; Jan and Feb 2010 from the kept SEC filing 10015870.pdf; May to Jul 2009 and Mar to Nov 2010 are not in the sheet: releases not found); figures in CAD except the uranium price (USD per lb). Units as printed; shares outstanding are not printed in the releases.",
              ["Release date", "As of", "NAV [CAD thousand]", "NAV per share [CAD]", "Shares outstanding", "U3O8 held [lb]", "UF6 held", "Uranium price used [USD/lb]", "Source", "Note"], rows_, [13, 12, 16, 14, 14, 14, 18, 16, 50, 70],
              "Monthly NAV, NAV per share and uranium holdings of UPC.", "Monthly", f"{min(z[1] for z in rows_):%Y-%m-%d} to {max(z[1] for z in rows_):%Y-%m-%d}",
              fmts=["yyyy-mm-dd", "yyyy-mm-dd", INT, NUM, INT, INT, None, NUM, None, None], wrap=(10,))
    ph = sorted((RAW / "physical").glob("*.csv")) if (RAW / "physical").exists() else []
    if ph:
        rows_ = []
        for p in ph:
            for x in read_csv(p):
                v = x["value"]
                try:
                    v = float(v) if v != "" else None
                except ValueError:
                    pass
                rows_.append([p.stem, x["year"], x["country_or_scope"], v, x["unit"], x["note"], x["source_url"]])
        table("Physical sources", "Physical-market series, long table", "V092-V109. One row per series, year and country or scope, as printed in the source. Read through page text, not from downloaded files: check a value against the source before quoting it. Files are in data/raw/physical.",
              ["Series (file)", "Year or date", "Country or scope", "Value", "Unit", "Note", "Source"], rows_, [52, 14, 34, 14, 28, 60, 50],
              "US and world production, requirements, trade, inventories and IAEA reactor statistics from EIA, IAEA, NEA, WNA and Euratom.", "Annual (irregular coverage)",
              "1993 to 2024 (varies by series)", fmts=[None, None, None, "#,##0.###", None, None, None], wrap=(6,), plain=True)

    # ---- source files kept whole
    meta = [(t, fe.TICKERS[t], g, cur) for t, cur, g in EQUITIES]
    rows_by = {**S["equities_raw"], **S["markets_raw"]}
    meta += [(t, fm.MARKETS[t], "commodity / index", "USD") for t in S["markets_raw"]]
    rows = []
    for t, desc, g, cur in meta:
        for r in rows_by[t]:
            rows.append([day(r["date"]), t, desc, g, cur] + [num(r[k]) for k in ("open", "high", "low", "close", "adj_close", "volume")])
    ws_ = wb_layout.write_table_sheet(wb, "Prices OHLCV", "Daily prices and volume, all Yahoo Finance tickers (V011-V014, V033-V034, V059, V119)",
                                      "Long table: one row per ticker and day, with open, high, low, close, adjusted close and volume, exactly as Yahoo returns them (the seven futures series come from the date-repaired files of 15 Nov 2026, data/raw/markets_dated). Rows with volume 0 are days without trading (Yahoo repeats the previous price); the Daily sheet leaves those days blank and holds the adjusted close only.",
                                      ["Date", "Ticker", "Description", "Group", "Currency", "Open", "High", "Low", "Close", "Adj. close", "Volume"], rows,
                                      [12, 10, 40, 18, 10, 12, 12, 12, 12, 12, 14], fmts=["yyyy-mm-dd", None, None, None, None, NUM4, NUM4, NUM4, NUM4, NUM4, INT], tab="dump", plain=True)
    sheets.append(("Source files kept whole", "Prices OHLCV", "Open, high, low, close, adjusted close and volume for every ticker retrieved.", "Daily", f"{min(r[0] for r in rows):%Y-%m-%d} to {max(r[0] for r in rows):%Y-%m-%d}", len(rows), 11))
    wb_layout.write_data_sheet(wb, "World Bank monthly", "World Bank Pink Sheet, monthly (all columns as published)",
                               "Monthly period averages as published by the World Bank, 1960 onwards, kept whole. Columns that match a handoff variable are placed first.",
                               pink_months, pk, "Month", "yyyy-mm", mk, width=15, tab="dump", pink=True, extra_label="Other Pink Sheet series (not in the handoff table)")
    sheets.append(("Source files kept whole", "World Bank monthly", "The Pink Sheet in full: energy, metals, agriculture and the World Bank indices.", "Monthly (averages)", rng(pink_months), len(pink_months), len(pk)))

    # ---- guide sheets
    cnt = {}
    for st, _w, _n in by.values():
        cnt[st] = cnt.get(st, 0) + 1
    order_st = ["In workbook", "Partly in workbook", "Script written, not run", "Not retrieved", "Paid or unavailable", DERIVED]
    status_counts = [(k, cnt[k]) for k in order_st if k in cnt]
    wb_layout.load_spot_link(ROOT / "data" / "spot_link.csv")  # 19 Nov 2026: Spot link column on Variables (labels from the Market Expert screen)
    wb_layout.write_variables_sheet(wb, by, extras)
    sheets.insert(0, ("Guide", "Variables", "Start here: every variable of the handoff table with its role in the thesis, why it matters and where it is in this workbook.", "-", "V001 to V119", 119 + len(extras), 13))
    wb_layout.write_table_sheet(wb, "Gaps", "Gaps and limits of the data", "What is missing, weak or open, by area.", ["Area", "Gap or limitation", "Effect", "Status / action"],
                                [[clean(c) for c in g] for g in GAPS], [18, 90, 50, 60], wrap=(2, 3, 4), tab="guide")
    sheets.insert(1, ("Guide", "Gaps", "Missing, weak or open items, by area.", "-", "-", len(GAPS), 4))
    wb_layout.write_contents(wb, sheets, status_counts, "19 November 2026", by, extras)

    names = ["Contents", "Variables", "Gaps", "Monthly", "Quarterly", "Annual", "Weekly", "Daily", "SPUT purchases", "SPUT issuance", "SPUT statements", "UxC volumes", "Yellow Cake reports", "UPC monthly NAV", "Physical sources", "Prices OHLCV", "World Bank monthly"]
    wb._sheets = [wb[n] for n in names if n in wb.sheetnames]
    wb.active = 0
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(status_counts)
    print(f"Saved {OUT}")
    for n in wb.sheetnames:
        print(f"  {n}: {wb[n].max_row} rows x {wb[n].max_column} columns")


if __name__ == "__main__":
    main()
