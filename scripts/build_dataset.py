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
import csv
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

WorksheetReader.bind_tables = lambda self: None  # the World Bank file has table links openpyxl cannot read

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
OUT = ROOT / "data" / "processed" / "thesis_dataset.xlsx"

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
EQ_NOTE = {"U-UN.TO": "data from 2006 look like the predecessor fund (not confirmed)", "U-U.TO": "USD units of the Sprott trust",
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
        D.append(col(f"{t} adj. close [{cur}]", {day(r["date"]): num(r["adj_close"]) for r in equities[t]}, "native", "Yahoo Finance", t, "daily", "daily close (adjusted for splits and dividends)", cur,
                     EQ_NOTE.get(t, ""), "H3 context" if grp in ("ETF", "physical fund") else "context", f"data/raw/yahoo/{t}.csv", NUM4))

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
    for t, cur, grp in EQUITIES:
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


# ---------------------------------------------------------------- writing
def style_header(ws, kinds):
    for i, k in enumerate(kinds, 1):
        c = ws.cell(1, i)
        c.font, c.fill = BOLD, PatternFill("solid", fgColor=FILL.get(k, "D9E2D5"))
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")


def write_table(wb, name, keys, cols, keyhdr, keyfmt, keyval, width=15, hdr_height=66):
    ws = wb.create_sheet(name)
    ws.append([keyhdr] + [c["header"] for c in cols])
    for k in keys:
        ws.append([keyval(k)] + [c["values"].get(k) for c in cols])
    for row in ws.iter_rows():
        for cell in row:
            cell.font = FONT
    style_header(ws, ["native"] + [c["kind"] for c in cols])
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 1).number_format = keyfmt
        for i, c in enumerate(cols, 2):
            ws.cell(r, i).number_format = c["fmt"]
    ws.column_dimensions["A"].width = 12
    for i in range(2, len(cols) + 2):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[1].height = hdr_height
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def write_text_sheet(wb, name, header, rows, widths, fills=None):
    ws = wb.create_sheet(name)
    ws.append(header)
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows():
        for cell in row:
            cell.font = FONT
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for cell in ws[1]:
        cell.font, cell.fill = BOLD, PatternFill("solid", fgColor="D9E2D5")
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def stats(c, freq):
    ks = sorted(c["values"])
    if not ks:
        return "", "", 0, ""
    f = lambda k: f"{k[0]}-{k[1]:02d}" if isinstance(k, tuple) else k.isoformat()
    miss = ""
    if freq == "monthly":
        miss = sum(1 for _ in months_between(ks[0], ks[-1])) - len(ks)
    elif freq == "quarterly":
        miss = len(quarters_between(ks[0], ks[-1])) - len(ks)
    return f(ks[0]), f(ks[-1]), len(ks), miss


def constructed(M, months):
    """Variables built from columns already in the Monthly sheet. Nothing here is a new observation."""
    import math
    import statistics
    byh = lambda start: next((c for c in M if c["header"].startswith(start)), None)
    idx = lambda k: k[0] * 12 + k[1]
    prev = lambda k, n=1: ((idx(k) - n - 1) // 12, (idx(k) - n - 1) % 12 + 1)
    spot, lt = byh("Uranium spot, month-end")["values"], byh("Uranium long-term, month-end")["values"]
    out = []

    def add(header, d, unit, note, role=""):
        out.append(col(header, d, "constructed", "Constructed in build_dataset.py", "", "monthly", "month-end", unit, note, role, "see note", NUM4))

    add("Spot minus long-term [USD/lb]", {k: spot[k] - lt[k] for k in spot if k in lt}, "USD per lb", "Uranium spot, month-end minus Uranium long-term, month-end", "context")
    ret = {k: math.log(spot[k] / spot[prev(k)]) for k in spot if prev(k) in spot}
    add("Uranium spot log return [log points]", ret, "log points", "ln(P_t / P_t-1) of the Cameco month-end spot", "H1-H4")
    ks, m12, s12 = sorted(ret), {}, {}
    for i, k in enumerate(ks):
        w = ks[max(0, i - 11):i + 1]
        if len(w) == 12 and idx(k) - idx(w[0]) == 11:
            r = [ret[x] for x in w]
            m12[k], s12[k] = statistics.mean(r), statistics.stdev(r)
    add("Uranium spot return, 12-month average [log points]", m12, "log points", "mean of the last 12 monthly log returns (momentum, as in Fan et al.)", "H1-H4 control")
    add("Uranium spot return, 12-month std. dev. [log points]", s12, "log points", "sample standard deviation of the last 12 monthly log returns", "H1-H4 control")
    stq = byh("UX futures close (Stooq)")["values"]
    add("UX futures basis, Stooq close minus Cameco spot [USD/lb]", {k: stq[k] - spot[k] for k in stq if k in spot}, "USD per lb",
        "spot-proxy basis; the Stooq contract roll is undocumented and spot is an assessment", "H2")
    add("UX futures basis [% of spot]", {k: (stq[k] / spot[k] - 1) * 100 for k in stq if k in spot}, "percent", "(Stooq close / Cameco spot - 1) x 100", "H2")
    prem = byh("SPUT premium/discount, month-end")["values"]
    add("SPUT in discount (1 = month-end premium below zero)", {k: 1 if v < 0 else 0 for k, v in prem.items()}, "0/1", "state variable from SPUT premium/discount, month-end; zero is the only cut-off used", "H3")
    cpi = byh("Consumer price index")
    if cpi:
        c = cpi["values"]
        add("CPI inflation, month on month [percent]", {k: (c[k] / c[prev(k)] - 1) * 100 for k in c if prev(k) in c}, "percent", "from the CPI-U column", "control")
        add("CPI inflation, year on year [percent]", {k: (c[k] / c[prev(k, 12)] - 1) * 100 for k in c if prev(k, 12) in c}, "percent", "from the CPI-U column", "control")
        last = max(k for k in c if k in spot)
        add("Uranium spot, real [USD/lb in last-month dollars]", {k: spot[k] * c[last] / c[k] for k in spot if k in c}, "USD per lb", f"spot deflated by CPI-U to {last[0]}-{last[1]:02d} dollars", "context")
    baa, aaa = byh("Baa corporate yield"), byh("Aaa corporate yield")
    if baa and aaa:
        add("Baa minus Aaa spread [percentage points]", {k: baa["values"][k] - aaa["values"][k] for k in baa["values"] if k in aaa["values"]}, "percentage points", "Baa yield minus Aaa yield, monthly", "H4")
    add("Dummy: SPUT launched (1 from Jul 2021)", {k: 1 if k >= (2021, 7) else 0 for k in months}, "0/1", "SPUT launch 19 Jul 2021 (month of launch = 1)", "H3 regime")
    add("Dummy: Kazakh unrest (1 in Jan 2022)", {k: 1 if k == (2022, 1) else 0 for k in months}, "0/1", "event month from the variables doc", "event")
    add("Dummy: Russia-Ukraine war (1 from Feb 2022)", {k: 1 if k >= (2022, 2) else 0 for k in months}, "0/1", "invasion 24 Feb 2022 (month of invasion = 1)", "event")
    return out


def coverage_rows():
    rows = []
    for (i, cat, var, variant, smst, src, st, where, note, req, step, purpose, issue, route) in COVERAGE:
        if req:
            st = ("Constructed" if st == "Constructed" else "In workbook") if (RAW / req).exists() else "Script written, not run"
        rows.append([i, step, cat, var, variant, purpose, smst, src, st, where, note, issue, route])
    return rows


def main():
    S = load()
    pk = pink()
    D, W, M, Q = S["daily"], S["weekly"], S["monthly"], S["quarterly"]

    # indexes
    spot = next(c for c in M if c["header"].startswith("Uranium spot, month-end"))["values"]
    months = list(months_between(min(spot), max(spot)))
    dates = set()
    for c in D:
        if c["sid"] != "DFF":
            dates |= set(c["values"])
    dropped_dff = sum(1 for c in D if c["sid"] == "DFF" for k in c["values"] if k not in dates)
    dates = sorted(dates)
    qkeys_all = [k for c in Q for k in c["values"] if c["freq"] == "quarterly"]
    quarters = quarters_between(min(qkeys_all), max(qkeys_all))
    wkeys = sorted(W[0]["values"])
    pink_months = list(months_between(min(min(c["values"]) for c in pk), max(max(c["values"]) for c in pk)))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    write_text_sheet(wb, "README", ["Thesis dataset"], [[""]], [110])  # filled at the end

    # counts for the analysis map
    def n_months(headers):
        cs = [next(c for c in M if c["header"].startswith(h)) for h in headers]
        return sum(1 for k in months if all(k in c["values"] for c in cs))
    has_mkt = all(t in S["markets_found"] for t in ("GC=F", "SI=F", "HG=F"))
    n_core = n_months(["Uranium spot, month-end", "WTI crude, spot, month-end", "Brent crude, spot, month-end"])
    n_h1 = n_months(["Uranium spot, month-end", "Gold futures (GC=F), month-end", "Silver futures (SI=F), month-end", "Copper futures (HG=F), month-end", "WTI crude, spot, month-end"]) if has_mkt else None
    n_h3 = n_months(["Uranium spot, month-end", "SPUT holdings, month-end"])
    n_h4 = n_months(["Uranium spot, month-end", "VIX, month-end", "Broad dollar index, month-end", "US 10y yield, month-end"])
    n_ux = sum(1 for k in months if k in next(c for c in M if c["header"].startswith("UX futures close (Stooq)"))["values"])
    n_yc = len(next(c for c in Q if c["header"].startswith("Yellow Cake NAV"))["values"])
    n_sput_q = sum(1 for k in quarters if k in next(c for c in Q if c["header"].startswith("Yellow Cake NAV"))["values"])

    CN = constructed(M, months)
    # Catalogue
    cat = []
    for sheet, cols_, freq in [("Monthly", M, "monthly"), ("Constructed", CN, "monthly"), ("Quarterly", Q, "quarterly"), ("Weekly", W, "weekly"), ("Daily", D, "daily"), ("Monthly_Pink", pk, "monthly")]:
        for c in cols_:
            first, last, n, miss = stats(c, freq)
            cat.append([sheet, c["header"], c["kind"], c["freq"], c["timing"], c["source"], c["sid"], c["unit"], first, last, n, miss, c["role"], c["note"], c["raw"]])
    cat.append(["Events_SPUT", "SPUT purchases: date, pounds acquired, cumulative pounds", "native", "irregular (purchase days)", "purchase date", "Sprott fund page", "sput_acquired", "lb U3O8",
                min(e[0] for e in S["events"]).isoformat(), max(e[0] for e in S["events"]).isoformat(), len(S["events"]), "", "H3", "list ends at the last purchase date", "data/raw/sprott/sput_acquired.csv"])
    for t, rows in list(S["equities_raw"].items()) + list(S["markets_raw"].items()):
        cat.append(["Daily_OHLCV", f"{t} open, high, low, close, adj. close, volume", "native", "daily", "daily", "Yahoo Finance", t, "listing currency", rows[0]["date"], rows[-1]["date"], len(rows), "",
                    "H3 context (volume)" if t in ("URA", "URNM", "NLR", "U-U.TO", "U-UN.TO", "YCA.L") else "context", EQ_NOTE.get(t, ""), f"data/raw/{'markets' if t in S['markets_raw'] else 'yahoo'}/{t}.csv"])

    amap = [
        ["H1 Integration", "Uranium vs gold, silver, copper, oil over time", "Monthly (month-end columns)", "Cameco spot; WTI and Brent spot (EIA); gold, silver, copper, platinum, gas futures (Yahoo, month-end) once fetch_markets.py has been run",
         f"{n_core} months with uranium + both oil series" + (f"; {n_h1} with gold, silver, copper too" if n_h1 else "; gold, silver, copper month-end not yet fetched (run fetch_markets.py)"),
         "Time-varying correlation or dynamic conditional correlation, rolling windows, structural-break and regime tests, cointegration", "Feasible. Monthly frequency only; the uranium price is assessed, not traded, so it is smoother than the comparators."],
        ["H2 Price discovery", "Do the other markets move before uranium?", "Monthly (month-end columns); Daily for the comparators", "Same as H1; UX futures close (Stooq)",
         f"{n_core} months; futures close on {n_ux} months", "Lead-lag regressions, Granger causality or VAR, information-share comparison of spot and futures",
         "Feasible at monthly frequency only (daily or weekly uranium spot is paid). Month-end values avoid the timing bias that period averages create. Futures are thin (volume 0) and 11 months come from other sources."],
        ["H3 Financial participation", "Does participation strengthen the links?", "Weekly, Monthly, Quarterly, Daily_OHLCV", "SPUT holdings, price, NAV, premium, monthly purchases (from Jul 2021); Yellow Cake NAV and holdings (quarterly); ETF volumes in Daily_OHLCV; UX volume and open interest (empty)",
         f"{n_h3} months with SPUT; {n_yc} quarters with Yellow Cake; none before 2021 for SPUT", "Sub-sample (2021 onwards) analysis, before/after split at SPUT's launch, interaction of participation proxies with the H1/H2 links",
         "Limited. No futures participation data, no participation measure before 2021, and no ETF assets or flows. Pre-2021 physical-fund holdings (Uranium Participation Corporation) not collected."],
        ["H4 Macro-financial exposure", "Is uranium sensitive to VIX, dollar and rates, and more so with participation?", "Monthly (month-end columns); Daily for the macro side", "VIX, broad dollar index, Fed funds, 10y yield, real yield, term and credit spreads, breakeven inflation, S&P 500",
         f"{n_h4} months with uranium + VIX + dollar + 10y yield", "Regressions of uranium returns on macro changes, rolling betas, interaction with a participation proxy",
         "Macro side feasible over the full period. The 'participation increases sensitivity' part depends on the same participation data as H3."],
        ["Fundamentals (context)", "Conversion price and fund NAV as uranium-specific controls", "Quarterly", "Conversion spot and long-term (NA, EU); Yellow Cake NAV and holdings",
         f"{n_yc} quarters of Yellow Cake; conversion complete from 2014", "Descriptive comparison and low-frequency checks only", "Too few observations for regression on its own. SWU (enrichment) prices not available."],
    ]
    gaps = [
        ["Futures", "Stooq UX has no bars for 11 months (2021-10 to 2022-03, 2023-04 to 2023-08)", "Futures series is incomplete in one column", "3 months filled from Investing (close only), 8 from the reconstructed settlement; each in its own column, none merged"],
        ["Futures", "Reconstructed settlements are benchmark-derived, not exchange bulletins; UxC month-end is the last Monday", "April 2023 would be 53.85 on a calendar month-end", "Open: check CME bulletins"],
        ["Futures", "Volume is 0 throughout; open interest exists only for Jan-Apr 2026 (310)", "No futures participation measure (H3, H4)", "Open: decision D-002 fallback"],
        ["Futures", "Stooq contract roll is undocumented; Stooq, Investing and settlements not confirmed to measure the same thing", "Small level differences between columns", "Open"],
        ["Spot price", "Cameco price is an industry average of assessed UxC and TradeTech prices, one point per month", "Smoother than a traded price; no intramonth data", "UxC weekly spot is paid; not collected"],
        ["Spot price", "Dec 2017 Cameco monthly 23.75 vs quarterly table 22.32", "One month", "Open"],
        ["Spot price", "IMF uranium runs about 20% below Cameco from 2021", "Cross-check disagrees", "Cause not found; IMF series kept as cross-check only"],
        ["Commodities", "Coal has no daily source; the Australian coal column is a monthly average", "Timing differs from the month-end columns", "Open"],
        ["Commodities", "Gold, silver, copper, platinum and gas month-end come from Yahoo continuous futures (and spot for gold and silver); rolls are not adjusted", "Small jumps at contract rolls", "Check against the Pink Sheet averages in Monthly_Pink"],
        ["Funds", "SPUT starts Jul 2021; pre-2021 Uranium Participation Corporation holdings not collected. Monthly UPC releases exist for 2009 to 2020, but UPC held U3O8 and UF6 and reported NAV in CAD", "No participation measure before 2021", "Open: retrieval is a long manual job (about 150 releases)"],
        ["Funds", "Yellow Cake NAV and holdings are quarterly only (from Sep 2018); some holdings are derived or carried (see lb basis)", "33 quarterly observations", "No monthly NAV is published"],
        ["Funds", "ETF assets under management and shares outstanding not collected", "ETF flows cannot be built; only price and volume", "Open"],
        ["Conversion / SWU", "Conversion complete from 2014-03 only; scattered quarters 2010-2013; nothing before 2010", "Short quarterly series", "UxC month-end prices are subscriber-only"],
        ["Conversion / SWU", "No SWU (enrichment) prices", "Missing fuel-cycle price", "Only annual EIA figures exist free"],
        ["Equities", "Currencies assigned by ticker; U-UN.TO data from 2006 look like the predecessor fund", "Check before comparing across tickers", "Open"],
        ["Process", "Cameco, Yellow Cake and conversion files were first collected by hand in the browser; the scripts are written but not yet run", "Not yet reproducible from scripts", "Run fetch_cameco.py, fetch_yellowcake.py, fetch_conversion.py, fetch_markets.py"],
        ["Process", "No native annual series has been retrieved", "No Annual sheet", "Candidates: annual production, demand and enrichment (EIA, WNA)"],
    ]

    # sheets in order
    write_table(wb, "Monthly", months, M, "Month", "yyyy-mm", lambda k: dt.date(k[0], k[1], 1), width=15, hdr_height=78)
    write_table(wb, "Constructed", months, CN, "Month", "yyyy-mm", lambda k: dt.date(k[0], k[1], 1), width=15, hdr_height=78)
    write_table(wb, "Quarterly", quarters, Q, "Quarter end", "yyyy-mm-dd", lambda k: k, width=16, hdr_height=66)
    write_table(wb, "Weekly", wkeys, W, "Date (Friday)", "yyyy-mm-dd", lambda k: k, width=16, hdr_height=40)
    write_table(wb, "Daily", dates, D, "Date", "yyyy-mm-dd", lambda k: k, width=14, hdr_height=66)
    ws = wb.create_sheet("Events_SPUT")
    ws.append(["Purchase date", "Pounds acquired [lb U3O8]", "Cumulative pounds acquired [lb U3O8]"])
    for r in sorted(S["events"]):
        ws.append(list(r))
    for row in ws.iter_rows():
        for cell in row:
            cell.font = FONT
    style_header(ws, ["native"] * 3)
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 1).number_format = "yyyy-mm-dd"
        ws.cell(r, 2).number_format = ws.cell(r, 3).number_format = INT
    for L, w in zip("ABC", (14, 20, 24)):
        ws.column_dimensions[L].width = w
    ws.row_dimensions[1].height = 40
    ws.freeze_panes = "A2"
    ws = wb.create_sheet("Daily_OHLCV")
    ws.append(["Date", "Ticker", "Description", "Group", "Currency", "Open", "High", "Low", "Close", "Adj. close", "Volume"])
    meta = [(t, fe.TICKERS[t], g, cur) for t, cur, g in EQUITIES]
    rows_by = {**S["equities_raw"], **S["markets_raw"]}
    meta += [(t, fm.MARKETS[t], "commodity / index", "USD") for t in S["markets_raw"]]
    for t, desc, g, cur in meta:
        for r in rows_by[t]:
            ws.append([day(r["date"]), t, desc, g, cur] + [num(r[k]) for k in ("open", "high", "low", "close", "adj_close", "volume")])
    for row in ws.iter_rows():
        for cell in row:
            cell.font = FONT
    style_header(ws, ["native"] * 11)
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 1).number_format = "yyyy-mm-dd"
        for c in range(6, 11):
            ws.cell(r, c).number_format = NUM4
        ws.cell(r, 11).number_format = INT
    for L, w in zip("ABCDEFGHIJK", (12, 10, 40, 18, 10, 12, 12, 12, 12, 12, 14)):
        ws.column_dimensions[L].width = w
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    write_table(wb, "Monthly_Pink", pink_months, pk, "Month", "yyyy-mm", lambda k: dt.date(k[0], k[1], 1), width=13, hdr_height=66)

    ev = [[dt.date(2005, 1, 1), "Start of the post-2005 speculation period (Fan et al. dummy); before the sample, so a constant 1 here", "Variables doc item 13"],
          [dt.date(2005, 3, 15), "Uranium Participation Corporation incorporated (the fund SPUT replaced)", "Variables doc item 44"],
          [dt.date(2021, 7, 19), "SPUT launch: first purchases recorded on the Sprott page; Sprott took over UPC in July 2021", "Variables doc items 13 and 44; Sprott page"],
          [dt.date(2022, 1, 1), "Kazakh unrest (January 2022; month only)", "Variables doc item 27 (UxC 2022)"],
          [dt.date(2022, 2, 24), "Russian invasion of Ukraine", "Variables doc items 13 and 27 (UxC 2022; Todorova 2025a)"]]
    wse = write_text_sheet(wb, "Events", ["Date", "Event", "Source"], ev, [14, 100, 60])
    for r in range(2, wse.max_row + 1):
        wse.cell(r, 1).number_format = "yyyy-mm-dd"
    cov = coverage_rows()
    write_text_sheet(wb, "Variable_Coverage", ["ID", "Logic step (ChatGPT handoff)", "Category", "Variable", "Variant", "Purpose (source map)", "Source-map status", "Source", "Workbook status",
                                               "Where in the workbook", "Note", "Remaining issue (source map)", "Retrieval route (source map)"], cov,
                     [8, 34, 24, 24, 34, 54, 14, 28, 24, 50, 60, 50, 40])
    write_text_sheet(wb, "Catalogue", ["Sheet", "Column", "Type", "Native frequency", "Timing", "Source", "Series ID", "Unit", "First", "Last", "Observations", "Missing periods inside range",
                                       "Proposed role (not agreed)", "Note / caveat", "Raw file"], cat, [14, 52, 13, 14, 26, 36, 22, 16, 11, 11, 12, 12, 20, 70, 44])
    write_text_sheet(wb, "Analysis_Map", ["Hypothesis", "Question", "Sheets", "Variables available", "Observations", "Candidate methods", "Feasibility and limits"], amap, [24, 36, 30, 60, 40, 50, 70])
    write_text_sheet(wb, "Gaps", ["Area", "Gap or limitation", "Effect", "Status / action"], gaps, [18, 80, 50, 60])

    readme = [
        ["Thesis dataset: uranium financialization. Built by scripts/build_dataset.py from data/raw/ (see DATA_LOG.md for every download)."],
        [""],
        ["HOW IT IS ORGANISED"],
        ["One sheet per native frequency. A series is in the sheet of the frequency at which it is published; nothing is averaged or dropped to fit another frequency."],
        ["Daily: market and macro series, SPUT daily, month rows are not aggregated. Weekly: SPUT holdings. Monthly: uranium prices, futures and month-end versions of daily series."],
        ["Quarterly: Yellow Cake NAV and holdings, conversion prices. Events_SPUT: SPUT purchase days. Daily_OHLCV: full open/high/low/close/adjusted close/volume for every Yahoo ticker (long table)."],
        ["Monthly_Pink: the World Bank Pink Sheet as published (monthly AVERAGES, 1960 onwards), kept as a cross-check. There is no Annual sheet: no annual series has been retrieved yet."],
        ["Constructed: variables built from Monthly columns (spot-minus-long-term, log return, 12-month momentum and volatility, futures basis, SPUT discount state, CPI inflation, real spot, Baa-Aaa, dummies). Events: dated events behind the dummies."],
        ["Variable_Coverage lists the 119 records of the Uranium Data Source Map (the ChatGPT handoff) and 11 extra items from the variables doc, with where each is in this workbook or why it is not."],
        ["Catalogue describes every column (source, unit, timing, coverage, caveat). Analysis_Map says which analyses each hypothesis can support. Gaps lists what is missing or weak."],
        [""],
        ["CONVENTIONS"],
        ["Month-end = the last available observation of the month. The main uranium price is Cameco's month-end spot, so the Monthly sheet uses month-end values for every daily series. Averages appear only where no daily source exists (coal, IMF cross-checks, Monthly_Pink)."],
        ["Header colours: green = native, blue = derived from a higher-frequency series (month-end or quarter-end), grey = period average, yellow = gap-fill from another source, orange = reconstructed. The Catalogue repeats this in the Type column."],
        ["Blank cell = no observation (market closed, series not started, source has no value). Nothing is interpolated or forward-filled. One source per column; sources are never mixed inside a column."],
        ["Dates are real Excel dates. Monthly rows show the first day of the month and mean that whole month; Quarterly rows use the quarter-end date."],
        ["Equity prices are adjusted close in each listing's own currency (see the header). Currencies are assigned by ticker; check before comparing across tickers."],
        [""],
        ["WHAT IS LEFT OUT OF THIS SHEET (all still in data/raw)"],
        [f"Daily Fed funds rate on weekends and market holidays when no other series trades ({dropped_dff} rows); the rate is unchanged on those days. The Nov 2026 Stooq futures bar (partial month) is outside the Monthly sheet, which ends at the last Cameco month ({months[-1][0]}-{months[-1][1]:02d})."],
        ["The other Pink Sheet sheets (indices, weights) and the Stooq/Investing files are not altered. Raw files are never edited."],
        [""],
        ["REBUILD"],
        ["python scripts/fetch_*.py (each downloader), then python scripts/build_dataset.py. Needs openpyxl."],
    ]
    ws = wb["README"]
    ws.delete_rows(1, 5)
    for r in readme:
        ws.append(r)
    for row in ws.iter_rows():
        for cell in row:
            cell.font = BOLD if (cell.value and str(cell.value).isupper()) or cell.row == 1 else FONT
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 150
    order = ["README", "Variable_Coverage", "Catalogue", "Analysis_Map", "Gaps", "Monthly", "Constructed", "Quarterly", "Weekly", "Daily", "Events_SPUT", "Events", "Daily_OHLCV", "Monthly_Pink"]
    wb._sheets = [wb[n] for n in order]
    from collections import Counter
    print(dict(Counter(r[8] for r in cov if r[0].startswith("V"))))
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"Saved {OUT}")
    for n in wb.sheetnames:
        print(f"  {n}: {wb[n].max_row - 1} rows x {wb[n].max_column} columns")


if __name__ == "__main__":
    main()
