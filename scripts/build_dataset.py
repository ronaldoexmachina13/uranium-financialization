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
OUT = Path(os.environ["THESIS_OUT"]) if os.environ.get("THESIS_OUT") else ROOT / "data" / "processed" / "thesis_dataset.xlsx"

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
EQ_NOTE = {"U-UN.TO": "SPUT only, from 26 Jul 2021 (the ticker's earlier history is Uranium Participation Corp, see the UPC column)", "U-U.TO": "USD units of the Sprott trust",
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


ISS_FIRST, ISS_LAST = (2025, 2), (2025, 12)  # months fully covered by the prospectus "Prior Sales" table (12 months to 22 Jan 2026, listed to 19 Jan 2026)


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
    print("SPUT reconciliations OK: unit roll-forward across all reports; H2 2025 issuance list equals the reported units and proceeds")

    mkey = lambda d: (d.year, d.month)
    me = lambda k: dt.date(k[0] + (k[1] == 12), k[1] % 12 + 1, 1) - dt.timedelta(days=1)
    full = list(months_between(ISS_FIRST, ISS_LAST))
    um, vm = defaultdict(int), defaultdict(float)
    for d, pr, n, _ in iss:
        if ISS_FIRST <= mkey(d) <= ISS_LAST:
            um[mkey(d)] += n
            vm[mkey(d)] += pr * n
    a_d, a_u = dt.date(2025, 6, 30), U("2025-06-30")
    rec = {k: a_u + sum(n for d, _, n, _ in iss if a_d < d <= me(k)) - sum(n for d, _, n, _ in iss if me(k) < d <= a_d) for k in full}
    need(rec[(2025, 12)] == U("2025-12-31"), "reconstructed Dec 2025 units = reported")
    stmt_src = "SPUT annual and interim reports (Sprott)"
    raw_s = "data/raw/sprott/sput_statements.csv"
    ends = [r for r in stm if r["period_type"] in ("annual", "interim")]

    def rep_col(field, hdr, unit, fmt):
        return col(hdr, {mkey(day(r["period_end"])): num(r[field]) for r in ends}, "native", stmt_src, field, "monthly (reporting dates only)", "reporting date", unit,
                   "blank in months without an annual or interim report; Statements_SPUT has the full table", "H3", raw_s, fmt)
    M.append(rep_col("units_outstanding", "SPUT units outstanding, reported", "units", INT))
    M.append(rep_col("nav_usd_thousand", "SPUT total NAV, reported", "USD thousand", INT))
    M.append(rep_col("cash_usd_thousand", "SPUT cash, reported", "USD thousand", INT))
    M.append(col("SPUT units outstanding, month-end, reconstructed", rec, "reconstructed", "SPUT interim report (30 Jun 2025) and base shelf prospectus (22 Jan 2026)", "units_outstanding", "monthly", "month-end",
                 "units", "units at 30 Jun 2025 plus or minus the dated issuances in the prospectus table; the result for Dec 2025 equals the reported figure. Feb-Dec 2025 only", "H3",
                 "data/raw/sprott/sput_issuances.csv", INT))
    M.append(col("SPUT units issued in month", {k: um.get(k, 0) for k in full}, "derived", "SPUT base shelf prospectus, 22 Jan 2026 (Prior Sales)", "units_issued", "monthly", "sum of issue dates in the month", "units",
                 "complete months only (the table covers 22 Jan 2025 to 19 Jan 2026); 0 = no issue listed", "H3", "data/raw/sprott/sput_issuances.csv", INT))
    M.append(col("SPUT gross issuance value in month [USD]", {k: vm.get(k, 0.0) for k in full}, "derived", "SPUT base shelf prospectus, 22 Jan 2026 (Prior Sales)", "price x units", "monthly", "sum of price x units in the month",
                 "USD", "an estimate of gross proceeds: for H2 2025 the sum equals the cash-flow proceeds in the annual report; net proceeds per issue are not published", "H3", "data/raw/sprott/sput_issuances.csv", INT))
    ts = [r for r in tsx if r["month"] < "2026-01"]
    for field, hdr, unit, fmt in [("uu_high_usd", "SPUT TSX high, U.U", "USD", NUM), ("uu_low_usd", "SPUT TSX low, U.U", "USD", NUM), ("uu_avg_daily_volume", "SPUT TSX average daily volume, U.U", "units", INT),
                                  ("uun_high_cad", "SPUT TSX high, U.UN", "CAD", NUM), ("uun_low_cad", "SPUT TSX low, U.UN", "CAD", NUM), ("uun_avg_daily_volume", "SPUT TSX average daily volume, U.UN", "units", INT)]:
        M.append(col(f"{hdr} [{unit}]", {(int(r["month"][:4]), int(r["month"][5:])): num(r[field]) for r in ts}, "native", "TMX via SPUT base shelf prospectus, 22 Jan 2026", field, "monthly", "calendar month", unit,
                     "Jan-Dec 2025 only; the table's January 2026 row covers 1-20 Jan and is kept in the raw file. Yahoo daily volumes agree within 0.03%", "H3", "data/raw/sprott/sput_tsx_monthly.csv", fmt))
    vol = defaultdict(float)
    for t in ("U-U.TO", "U-UN.TO"):
        for r in equities[t]:
            d = day(r["date"])
            if d >= dt.date(2021, 8, 1) and num(r["volume"]) is not None:
                vol[mkey(d)] += num(r["volume"])
    M.append(col("SPUT TSX volume, month total, U.U + U.UN [units]", dict(vol), "derived", "Yahoo Finance", "U-U.TO + U-UN.TO volume", "monthly", "sum of daily volume in the month", "units",
                 "both TSX lines; from Aug 2021 (first full month); matches the TSX monthly averages in the prospectus within 0.03% for every month of 2025", "H3", "data/raw/yahoo/U-U.TO.csv, U-UN.TO.csv", INT))
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
    # daily (published on weekdays, or on every calendar day for the GPR)
    for f, hdr, unit, sid, src, note, me in [
        ("fred_TEDRATE.csv", "TED spread (discontinued)", "percent", "TEDRATE", "FRED", "3-month LIBOR minus 3-month T-bill; discontinued by FRED, ends 2022-01-21", True),
        ("fred_DCPF3M.csv", "Commercial paper, 90-day AA financial", "percent", "DCPF3M", "FRED (Federal Reserve)", "candidate replacement for the TED spread when paired with the T-bill column; the spread is not computed here", True),
        ("fred_DCPN3M.csv", "Commercial paper, 90-day AA nonfinancial", "percent", "DCPN3M", "FRED (Federal Reserve)", "alternative to the financial commercial paper column", True),
        ("fred_DTB3.csv", "T-bill, 3-month, secondary market, discount basis", "percent", "DTB3", "FRED (Federal Reserve)", "discount basis; the 3-month Treasury yield used elsewhere (DGS3MO) is constant-maturity", True),
        ("fred_SOFR.csv", "SOFR, secured overnight financing rate", "percent", "SOFR", "FRED (NY Fed)", "starts 2018-04-03", True),
        ("gpr_GPRD_daily.csv", "Geopolitical risk index, daily overall (GPRD)", "index 1985-2019=100", "GPRD", "Caldara and Iacoviello", "calendar days including weekends; weekend values are not shown on this sheet (the date list is trading days only): full file in data/raw/macro", False),
        ("gpr_GPRD_THREAT_daily.csv", "Geopolitical risk index, daily threats (GPRD_THREAT)", "index 1985-2019=100", "GPRD_THREAT", "Caldara and Iacoviello", "calendar days; weekend values not shown on this sheet", False),
        ("gpr_GPRD_ACT_daily.csv", "Geopolitical risk index, daily acts (GPRD_ACT)", "index 1985-2019=100", "GPRD_ACT", "Caldara and Iacoviello", "calendar days; weekend values not shown on this sheet", False)]:
        v = series(f)
        D.append(col(f"{hdr} [{unit}]", v, "native", src, sid, "daily", "daily value", unit, note, "", raw(f), NUM))
        if me:
            M.append(col(f"{hdr}, month-end [{unit}]", month_end(v), "derived", src, sid, "monthly", "last observation of the month (from Daily)", unit, note, "", raw(f), NUM))

# ---------------------------------------------------------------- load
def load():
    S = {}
    S["daily"], S["weekly"], S["monthly"], S["quarterly"] = [], [], [], []
    cam = read_csv(RAW / "cameco" / "uranium_price_monthly.csv")
    spot = {ym(day(r["date"])): num(r["spot_usd_lb"]) for r in cam}
    lt = {ym(day(r["date"])): num(r["longterm_usd_lb"]) for r in cam}
    M = S["monthly"]
    cam_raw = "data/raw/cameco/uranium_price_monthly.csv"
    M.append(col("Uranium spot, month-end [USD/lb]", spot, "native", "Cameco (industry average of UxC and TradeTech month-end prices)", "spot_usd_lb", "monthly",
                 "month-end", "USD per lb", "MAIN PRICE. Dec 2017 differs from Cameco's own quarterly table (23.75 here, 22.32 there)", "H1-H4 (dependent variable)", cam_raw, NUM))
    M.append(col("Uranium long-term, month-end [USD/lb]", lt, "native", "Cameco (industry average of UxC and TradeTech month-end prices)", "longterm_usd_lb", "monthly",
                 "month-end", "USD per lb", "", "context", cam_raw, NUM))
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
    sd = read_csv(RAW / "sprott" / "sput_daily.csv")
    for k, hdr, unit, f in [("price_usd", "SPUT price", "USD", NUM), ("nav_usd", "SPUT NAV per unit", "USD", NUM), ("premium_discount_pct", "SPUT premium/discount", "percent", NUM)]:
        D.append(col(f"{hdr} [{unit}]", {day(r["date"]): num(r[k]) for r in sd}, "native", "Sprott fund page", k, "daily", "daily close", unit, "starts Jul 2021", "H3", "data/raw/sprott/sput_daily.csv", f))
    mdir = RAW / "markets"
    markets = {}
    for t in fm.MARKETS:
        p = mdir / f"{t}.csv"
        if p.exists():
            hdr, unit, role = MARKET_NAMES[t]
            markets[t] = read_csv(p)
            note = "continuous front-month; the contract changes at each roll, not back-adjusted" if t.endswith("=F") else ("spot quote from Yahoo" if t.endswith("=X") else "")
            D.append(col(f"{hdr} close [{unit}]", {day(r["date"]): num(r["close"]) for r in markets[t]}, "native", "Yahoo Finance", t, "daily", "daily close", unit, note, role, f"data/raw/markets/{t}.csv", NUM))
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
                         "no overlap with SPUT, so do not compute a return across the switch to the next column; days with zero volume are left blank", "H3 context", "data/raw/yahoo/U-UN.TO.csv", NUM4))
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
    Q.append(col("Yellow Cake uranium held [million lb U3O8]", {day(r["quarter_end"]): num(r["uranium_held_million_lb"]) for r in yc}, "native", ysrc, "uranium_held_million_lb", "quarterly", "quarter-end", "million lb U3O8", "see lb_basis", "H3", yraw, NUM))
    Q.append(col("Yellow Cake lb basis", {day(r["quarter_end"]): r["lb_basis"] for r in yc}, "native", ysrc, "lb_basis", "quarterly", "quarter-end", "text", "stated = in the announcement; derived = holdings value divided by quarter-end spot; carried = previous quarter", "H3", yraw, "@"))
    Q.append(col("Yellow Cake quarter-end spot used [USD/lb]", {day(r["quarter_end"]): num(r["spot_usd_lb"]) for r in yc}, "native", ysrc, "spot_usd_lb", "quarterly", "quarter-end", "USD per lb", "as given or derived for the announcement; compare with the Cameco quarter-end column", "H3", yraw, NUM))
    Q.append(col("Yellow Cake announcement date", {day(r["quarter_end"]): day(r["announcement_date"]) for r in yc if r["announcement_date"]}, "native", ysrc, "announcement_date", "quarterly", "date of announcement", "date", "", "H3", yraw, "yyyy-mm-dd"))
    conv = read_csv(RAW / "cameco" / "conversion_prices_quarterly.csv")
    craw = "data/raw/cameco/conversion_prices_quarterly.csv"
    csrc = "Cameco quarterly reports (average of TradeTech and UxC)"
    for hdr, k in [("Conversion spot, North America", "conv_spot_na_usd_kgu"), ("Conversion spot, Europe", "conv_spot_eu_usd_kgu"),
                   ("Conversion long-term, North America", "conv_lt_na_usd_kgu"), ("Conversion long-term, Europe", "conv_lt_eu_usd_kgu")]:
        Q.append(col(f"{hdr} [USD/kgU]", {day(r["quarter_end"]): num(r[k]) for r in conv}, "native", csrc, k, "quarterly", "quarter-end", "USD per kgU as UF6",
                     "complete from 2014-03; only scattered quarters 2010-2013; nothing before 2010; no SWU prices", "context", craw, NUM))
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
        d = {day(r["date"]): num(r[k]) for r in cam if day(r["date"]).month in (3, 6, 9, 12)}
        Q.append(col(f"{hdr}, quarter-end (Cameco) [USD/lb]", d, "derived", "Cameco", k, "quarterly", "month-end value of Mar, Jun, Sep, Dec", "USD per lb", "from the Monthly sheet source", "context", cam_raw, NUM))
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
    "V031": (DERIVED, "Inputs: Monthly, SPUT TSX volume (V014) and SPUT units outstanding (V021)", "Units outstanding are known only at reporting dates and for Feb-Dec 2025."),
    "V032": ("Partly in workbook", "Quarterly: SPUT purchases as % of UxC spot volume; UxC volumes", "Denominator is the UxC spot volume quoted in Yellow Cake updates (quarterly, from 2021 Q3); it is not continuous and may include SPUT itself."),
    "V101": ("Partly in workbook", "Monthly and Quarterly: UxC spot volume, first reported; UxC volumes (every figure and vintage)", "Only where Yellow Cake quotes UxC; revised between updates; no continuous history."),
    "V035": ("In workbook", "Yellow Cake reports: Shares in issue", "As printed in each report, including treasury shares."),
    "V036": ("Partly in workbook", "Yellow Cake reports: New shares issued; Gross capital raised", "Filled where a report states one placing; multiple placings are described in the note, not summed."),
    "V037": ("Partly in workbook", "Yellow Cake reports: Uranium purchased in period", "Only where the report states the quarter's purchases; elected but undelivered purchases left blank."),
    "V038": ("Partly in workbook", "Yellow Cake reports: of which from Kazatomprom", ""),
    "V039": ("Partly in workbook", "Yellow Cake reports: of which other", ""),
    "V041": ("Partly in workbook", "Yellow Cake reports: Cash; Cash and other net current assets", "Cash is stated separately only in some reports."),
    "V044": ("Partly in workbook", "Yellow Cake reports: Shares bought back; Buyback cost", "Only the periods with a buyback programme; early-2020 counts are cumulative and left blank."),
    "V093": ("Partly in workbook", "Physical sources: WNA 2015-2024; Red Book 2018 edition (2014-2017), 2022 edition (2018-2021) and 2024 edition (2020-2023) production by country", "Tonnes U. WNA values re-read against the WNA page on 8 Nov 2026 (70 values, no differences); Red Book country rows sum to their totals except the 2018 edition cumulative column (+165). Editions overlap (2020 and 2021) and revise each other: both are kept, the note names the edition. The 2021 WNA tU3O8 total (46,374) is printed so at source: use tonnes U. WNA and Red Book differ for China, India, Ukraine, South Africa and the USA."),
    "V097": ("Partly in workbook", "Physical sources: IAEA TECDOC-2030 inventories at end 2021 (two regions); Euratom EU utility stocks 2010 and 2012", "No continuous global series exists; as-of dates, owners and material forms differ."),
    "V076": ("Partly in workbook", "Daily: TED spread (to Jan 2022), commercial paper 90-day AA, T-bill 3-month, SOFR; Monthly: month-end versions", "Raw rates only: no spread is computed and no single replacement for TED is chosen."),
    "V082": ("Partly in workbook", "Monthly: industrial production OECD total and G7 (to 2024-03), OECD growth", "OECD level series end 2024-03 on FRED; US industrial production is V081."),
    "V083": ("In workbook", "Quarterly: real GDP, US, level and annualised growth", "US only; OECD and other national GDP not retrieved."),
    "V084": ("In workbook", "Monthly: Kilian global real economic activity index", "From the Dallas Fed file."),
    "V088": ("In workbook", "Monthly: categorical EPU, 11 categories", "No energy category in the categorical file."),
    "V089": ("In workbook", "Monthly and Daily: geopolitical risk index, overall", "Daily file covers calendar days; weekends are not shown on the Daily sheet."),
    "V090": ("In workbook", "Monthly and Daily: geopolitical risk, threats and acts", "Same note as V089."),
    "D21": ("In workbook", "Monthly: Google Trends 'uranium ETF'", "Relative interest, peak = 100; carries measurement error."),
    "D44": ("Partly in workbook", "UPC monthly NAV, Dec 2010 to May 2021; Daily and Monthly: UPC adj. close, 2006 to Jul 2021", "2009 to Nov 2010 and June 2021 not retrieved; shares outstanding not printed in the releases."),
    "V092": ("Partly in workbook", "Annual: US uranium concentrate production, 2011-2025", "US only. Global and country production (V093) not retrieved."),
    "V095": ("Partly in workbook", "Annual: US reactor owners inventory (total and by material), year-end 2021-2025; Physical sources: 2011-2015 (UMAR 2015), 2020 and 2024 preliminary (UMAR 2024), 2016-2019 as rounded totals (Table S3a)", "US only. 2016-2019 exist only as totals rounded to 0.1 million lb (EIA does not give the detailed tables for those years). The 2024 preliminary value (126.4 million lb) was revised to 115.1 in the 2025 edition."),
    "V096": ("Partly in workbook", "Annual: US supplier inventory (uranium concentrate; UF6 withheld), year-end 2021-2025; Physical sources: 2011-2015, 2020 and 2024 preliminary, 2016-2019 as rounded totals", "Natural and enriched UF6 held by suppliers are withheld by EIA."),
    "V098": ("Partly in workbook", "Annual: US reactor requirements, maximum anticipated, 2026-2035; Physical sources: NEA/IAEA Red Book 2018, 2022 and 2024 world requirements by region (actual years 2016, 2020, 2022; projections to 2050)", "US figures forward-looking (EIA). World figures are Red Book scenarios, not observations, and the editions revise each other (the 2035 low case is 53,010, 60,544 and 77,242 tU in the 2018, 2022 and 2024 editions)."),
    "V042": ("In workbook", "Quarterly: Yellow Cake NAV per share (with the quarter-end spot price used); Yellow Cake reports: total NAV and NAV per share as stated, per report", "Total NAV is the figure each of the 48 reports states (read from the PDFs), never shares times NAV per share. Three reports state a pro forma or later figure; the basis column says which."),
    "D17": ("Partly in workbook", "Monthly: S&P 500 (^GSPC, from 1988) and S&P 500 (FRED, last ten years)", "MSCI World is licensed and not retrieved."),
    "V099": ("Partly in workbook", "Annual: US unfilled requirements, as of end 2024 and end 2025", "Forward-looking; two vintages only."),
    "V100": ("Partly in workbook", "Annual: US contracted deliveries, minimum and maximum, in effect end 2025", "Forward-looking; one vintage; US/foreign split withheld by EIA."),
    "D47": ("Partly in workbook", "SPUT issuance (12 May 2025 to 19 Jan 2026); SPUT statements (period totals)", "Sprott publishes only annual and June reports, so no quarterly flows exist; earlier dated issuances would need earlier prospectuses."),
}


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
    freq("Monthly", "Monthly series", "One row per calendar month. Prices and indicators are month-end values; period averages appear only where no daily source exists (type colour: grey).",
         months, M, "Month", "yyyy-mm", mk, "Uranium prices and the month-end value of every daily series, plus monthly-published series.", "Monthly", width=17)
    freq("Quarterly", "Quarterly series", "One row per quarter-end. Series published once a quarter, plus the quarter-end value of monthly series.",
         quarters, Q, "Quarter end", "yyyy-mm-dd", lambda k: k, "Conversion prices, Yellow Cake NAV and holdings, UxC spot volume and SPUT purchases per quarter.", "Quarterly", width=18)
    freq("Weekly", "Weekly series", "One row per Friday.", wkeys, W, "Date (Friday)", "yyyy-mm-dd", lambda k: k, "SPUT holdings from the Sprott fund page.", "Weekly", width=18)
    freq("Daily", "Daily series", "One row per trading day. Blank cell = market closed or no observation; nothing is carried forward.",
         dates, D, "Date", "yyyy-mm-dd", lambda k: k, "Market, rate, risk and fund series at their daily frequency.", "Daily", width=16)

    ann = annual_eia()
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
        cum, rows = 0, []
        for d_, pr_, n_, t_ in S["issuances"]:
            cum += n_
            rows.append([d_, pr_, n_, pr_ * n_, t_, cum])
        table("SPUT issuance", "SPUT unit issuance", "V022 (new units), V023 (issue price) and V024 (gross proceeds). Dates from the Prior Sales table of the 22 Jan 2026 base shelf prospectus; gross value = price × units.",
              ["Issue date", "Price per unit [USD]", "Units issued", "Gross value, price × units [USD]", "Issue type", "Cumulative units in this list"], rows, [14, 18, 16, 26, 36, 22],
              "Dated unit issues of SPUT (private placement, bought deal, at-the-market).", "Irregular", f"{rows[0][0]:%Y-%m-%d} to {rows[-1][0]:%Y-%m-%d}",
              fmts=["yyyy-mm-dd", "0.0000", INT, INT, None, INT])
        heads = ["Period end", "Period type", "Units outstanding", "Total NAV [USD thousand]", "NAV per unit [USD]", "TSX closing price, U.U [USD]", "Cash [USD thousand]",
                 "Uranium oxide at fair value [USD thousand]", "Pounds U3O8 held", "Units issued in the period", "Proceeds from issuance [USD thousand]", "Issue costs, cash flow [USD thousand]",
                 "Purchases of uranium oxide [USD thousand]", "Source", "Note", "Net proceeds, proceeds minus costs [USD thousand]", "Average issue price [USD]", "Pounds per unit [lb]",
                 "Uranium oxide fair value per lb [USD]", "Cash as % of NAV"]
        fields = ["period_end", "period_type", "units_outstanding", "nav_usd_thousand", "nav_per_unit_usd", "tsx_close_usd", "cash_usd_thousand", "uranium_oxide_fv_usd_thousand",
                  "pounds_u3o8_held", "units_issued_in_period", "proceeds_from_issuance_usd_thousand", "issue_costs_cashflow_usd_thousand", "purchases_of_uranium_usd_thousand"]
        rows = []
        for r in S["statements"]:
            v = [num(r[f]) if f not in ("period_end", "period_type") else r[f] for f in fields]
            v[0] = day(v[0])
            units_, nav_, cash_, fv_, lb_, iss_, pr_, co_ = v[2], v[3], v[6], v[7], v[8], v[9], v[10], v[11]
            rows.append(v + [r["source"], r["note"], pr_ - co_ if pr_ is not None and co_ is not None else None, pr_ * 1000 / iss_ if pr_ is not None and iss_ else None,
                             lb_ / units_ if lb_ and units_ else None, fv_ * 1000 / lb_ if fv_ and lb_ else None, cash_ / nav_ * 100 if cash_ is not None and nav_ else None])
        table("SPUT statements", "SPUT statements at reporting dates",
              "V019 (total NAV), V021 (units), V024 (proceeds), V026 (cash), V027 (purchases) and V029 (pounds held). Annual and interim reports and the prospectus; the last five columns are ratios of the others.",
              heads, rows, [13, 14, 16, 16, 14, 16, 14, 18, 16, 16, 18, 18, 18, 36, 46, 18, 14, 14, 16, 12],
              "SPUT balance sheet and cash flow at each annual, interim and prospectus date.", "Reporting dates", f"{rows[0][0]:%Y-%m-%d} to {rows[-1][0]:%Y-%m-%d}",
              fmts=["yyyy-mm-dd", None, INT, INT, "0.0000", NUM, INT, INT, INT, INT, INT, INT, INT, None, None, INT, NUM, "0.0000", NUM, NUM], wrap=(14, 15))
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
        # total NAV as stated in the same report (read from the PDFs 8 Nov 2026); inserted after "Cash and other net current assets"
        navf = RAW / "yellowcake" / "yellowcake_total_nav.csv"
        nav_ = {(day(x["announcement_date"]), day(x["period_end"])): x for x in read_csv(navf)} if navf.exists() else {}
        for z in rows_:
            x = nav_.get((z[0], z[1]))
            z[10:10] = [num(x["total_nav_usd_m"]) if x else None, num(x["nav_per_share_gbp"]) if x else None, (x["basis"] or None) if x else None]
        heads_[10:10] = ["Total NAV, as stated [USD million]", "NAV per share, as stated [GBP]", "NAV basis"]
        table("Yellow Cake reports", "Yellow Cake plc: facts from its reports", "V035-V044. One row per quarterly update, interim report and annual result, read from the company PDFs. Blank = not stated in that report. Period figures are never summed across lots. Total NAV is the figure the report states, never shares times NAV per share.",
              heads_, rows_, [14, 12, 16, 16, 16, 16, 16, 14, 12, 18, 16, 14, 30, 14, 14, 40, 70, 10], "Shares, capital raised, purchases, cash, total NAV and buybacks from every Yellow Cake report.", "Quarterly / half-year / annual",
              f"{min(z[1] for z in rows_ if z[1]):%Y-%m-%d} to {max(z[1] for z in rows_ if z[1]):%Y-%m-%d}",
              fmts=["yyyy-mm-dd", "yyyy-mm-dd", INT, INT, NUM, INT, INT, INT, NUM, NUM, NUM, NUM, None, INT, NUM, None, None, None], wrap=(13, 17))

    up = RAW / "upc" / "upc_monthly_nav.csv"
    if up.exists():
        r_ = read_csv(up)
        rows_ = [[day(x["release_date"]), day(x["as_of_date"]), num(x["nav_cad_thousand"]), num(x["nav_per_share_cad"]), num(x["shares_outstanding"]), num(x["u3o8_lb"]), x["uf6_kgu_or_lb_and_unit"],
                  num(x["uranium_price_used_usd"]), x["source_url"], x["note"]] for x in r_]
        table("UPC monthly NAV", "Uranium Participation Corporation: monthly net asset value", "D44. The fund that SPUT replaced (Sprott took it over in July 2021). Monthly releases, December 2010 to May 2021; figures in CAD except the uranium price (USD per lb). Units as printed; shares outstanding are not printed in the releases.",
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
                                      "Long table: one row per ticker and day, with open, high, low, close, adjusted close and volume, exactly as Yahoo returns them. Rows with volume 0 are days without trading (Yahoo repeats the previous price); the Daily sheet leaves those days blank and holds the adjusted close only.",
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
    wb_layout.write_variables_sheet(wb, by, extras)
    sheets.insert(0, ("Guide", "Variables", "Start here: every variable of the handoff table with its role in the thesis, why it matters and where it is in this workbook.", "-", "V001 to V119", 119 + len(extras), 13))
    wb_layout.write_table_sheet(wb, "Gaps", "Gaps and limits of the data", "What is missing, weak or open, by area.", ["Area", "Gap or limitation", "Effect", "Status / action"],
                                [[clean(c) for c in g] for g in GAPS], [18, 90, 50, 60], wrap=(2, 3, 4), tab="guide")
    sheets.insert(1, ("Guide", "Gaps", "Missing, weak or open items, by area.", "-", "-", len(GAPS), 4))
    wb_layout.write_contents(wb, sheets, status_counts, "8 November 2026", by, extras)

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
