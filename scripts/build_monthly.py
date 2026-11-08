"""Build data/processed/monthly_panel.xlsx from the raw files in data/raw/.

Run from the repo root after the fetch scripts:  python scripts/build_monthly.py
Needs openpyxl.

Rules (also recorded in DATA_LOG.md):
- One source per series; sources are never mixed inside a column.
- Main uranium price: Cameco month-end spot (industry average of UxC and TradeTech month-end prices). The IMF series is kept as a cross-check only (it drifts
  about 20% below the market from 2021).
- Timing: the main price is a month-end price, so daily series use the LAST available value of each month.
  The Pink Sheet commodities are only published as monthly averages and are used as published.
- The panel covers Jan 2007 to the last month with a Cameco spot price.
"""
import csv
import datetime as dt
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet._reader import WorksheetReader

WorksheetReader.bind_tables = lambda self: None  # the World Bank file has table links openpyxl cannot read

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
OUT = ROOT / "data" / "processed" / "monthly_panel.xlsx"
START = (2007, 1)

FRED_DAILY = [  # (column, FRED id, unit)
    ("VIX", "VIXCLS", "index points"),
    ("Dollar index (broad)", "DTWEXBGS", "index"),
    ("Fed funds rate", "DFF", "percent"),
    ("US 10y yield", "DGS10", "percent"),
    ("US 10y real yield", "DFII10", "percent"),
    ("Term spread 10y-3m", "T10Y3M", "percentage points"),
    ("Baa-10y spread", "BAA10Y", "percentage points"),
    ("Breakeven inflation 10y", "T10YIE", "percent"),
    ("USD/CAD", "DEXCAUS", "CAD per USD"),
    ("S&P 500", "SP500", "index"),
]
PINK = [  # (column, Pink Sheet name, unit)
    ("Gold", "Gold", "USD per troy oz"),
    ("Silver", "Silver", "USD per troy oz"),
    ("Platinum", "Platinum", "USD per troy oz"),
    ("Copper", "Copper", "USD per tonne"),
    ("Coal, Australian", "Coal, Australian", "USD per tonne"),
    ("Natural gas, US", "Natural gas, US", "USD per MMBtu"),
    ("Brent crude", "Crude oil, Brent", "USD per barrel"),
    ("WTI crude", "Crude oil, WTI", "USD per barrel"),
]
EQUITIES = [  # (ticker, listing currency assigned by ticker)
    ("URA", "USD"), ("URNM", "USD"), ("NLR", "USD"), ("CCJ", "USD"), ("KAP.IL", "USD"), ("U-U.TO", "USD"),
    ("U-UN.TO", "CAD"), ("YCA.L", "GBp"), ("UEC", "USD"), ("DNN", "USD"), ("NXE", "USD"), ("PDN.AX", "AUD"),
    ("BOE.AX", "AUD"), ("DYL.AX", "AUD"), ("SLX.AX", "AUD"), ("BWXT", "USD"), ("1164.HK", "HKD"),
]


def ym(s):
    return int(s[:4]), int(s[5:7])


def read_csv(path):
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def month_end(rows, datecol, valcol):
    """Last non-empty value of each month from dated rows."""
    out = {}
    for r in sorted(rows, key=lambda r: r[datecol]):
        v = r[valcol]
        if v not in ("", "."):
            out[ym(r[datecol])] = float(v)
    return out


def read_pink():
    ws = openpyxl.load_workbook(RAW / "worldbank" / "CMO-Historical-Data-Monthly.xlsx", data_only=True)["Monthly Prices"]
    names = {ws.cell(5, c).value.strip(): c for c in range(2, ws.max_column + 1) if ws.cell(5, c).value}
    series = {}
    for name, col in names.items():
        d = {}
        for r in range(7, ws.max_row + 1):
            p, v = ws.cell(r, 1).value, ws.cell(r, col).value
            if p and isinstance(v, (int, float)):
                d[(int(p[:4]), int(p[5:7]))] = float(v)
        series[name] = d
    return series


def main():
    data, meta = {}, {}  # column -> {(y, m): value}; column -> (series, source, unit, monthly value)

    cam = read_csv(RAW / "cameco" / "uranium_price_monthly.csv")
    data["Uranium spot (Cameco)"] = {ym(r["date"]): float(r["spot_usd_lb"]) for r in cam}
    meta["Uranium spot (Cameco)"] = ("month-end spot", "Cameco", "USD per lb", "month-end price as published - MAIN PRICE")
    data["Uranium long-term (Cameco)"] = {ym(r["date"]): float(r["longterm_usd_lb"]) for r in cam}
    meta["Uranium long-term (Cameco)"] = ("month-end long-term", "Cameco", "USD per lb", "month-end price as published")

    imf = read_csv(RAW / "fred" / "PURANUSDM.csv")
    data["Uranium (IMF, cross-check)"] = {ym(r["observation_date"]): float(r["PURANUSDM"]) for r in imf if r["PURANUSDM"] not in ("", ".")}
    meta["Uranium (IMF, cross-check)"] = ("PURANUSDM", "FRED (IMF)", "USD per lb", "monthly average as published; about 20% below Cameco from 2021")

    stq = read_csv(RAW / "stooq" / "ux_futures_monthly.csv")
    data["Uranium futures close (Stooq)"] = {ym(r["date"]): float(r["close"]) for r in stq}
    meta["Uranium futures close (Stooq)"] = ("UX.F", "Stooq", "USD per lb", "monthly bar close; 11 months missing at source")
    inv = read_csv(RAW / "investing" / "ux_futures_monthly.csv")
    data["Uranium futures close (Investing, gap fill)"] = {ym(r["date"]): float(r["close"]) for r in inv}
    meta["Uranium futures close (Investing, gap fill)"] = ("UXXc2 page", "Investing.com", "USD per lb", "monthly close; not merged with Stooq; Jan 2022 to Feb 2023 only")

    rec = read_csv(RAW / "uxc" / "ux_final_settlement_reconstructed.csv")
    data["Uranium futures final settlement (reconstructed from UxC benchmark)"] = {ym(r["month"] + "-01"): float(r["price_usd_lb"]) for r in rec}
    meta["Uranium futures final settlement (reconstructed from UxC benchmark)"] = ("UX final settlement", "UxC benchmark via Cameco and Yellow Cake documents", "USD per lb", "reconstructed, 8 months Stooq lacks; see data/raw/uxc; not verified against CME bulletins")

    yc = read_csv(RAW / "yellowcake" / "nav_quarterly.csv")
    data["Yellow Cake NAV per share"] = {ym(r["quarter_end"]): float(r["nav_per_share_gbp"]) for r in yc}
    meta["Yellow Cake NAV per share"] = ("nav_quarterly", "Yellow Cake plc quarterly operating updates", "GBP", "QUARTERLY: filled at quarter-end months only, blank otherwise (no monthly NAV is published)")
    data["Yellow Cake uranium held"] = {ym(r["quarter_end"]): float(r["uranium_held_million_lb"]) for r in yc if r["uranium_held_million_lb"]}
    meta["Yellow Cake uranium held"] = ("nav_quarterly", "Yellow Cake plc quarterly operating updates", "million lb U3O8", "QUARTERLY, quarter-end months only; see lb_basis in the raw file (stated, derived or carried)")
    conv = read_csv(RAW / "cameco" / "conversion_prices_quarterly.csv")
    for col, key in [("Conversion spot, North America", "conv_spot_na_usd_kgu"), ("Conversion spot, Europe", "conv_spot_eu_usd_kgu"),
                     ("Conversion long-term, North America", "conv_lt_na_usd_kgu"), ("Conversion long-term, Europe", "conv_lt_eu_usd_kgu")]:
        data[col] = {ym(r["quarter_end"]): float(r[key]) for r in conv if r[key]}
        meta[col] = (key, "Cameco quarterly reports (TradeTech and UxC average)", "USD per kgU as UF6", "QUARTERLY: quarter-end months only; complete from 2014-03, only scattered quarters from 2010 to 2013; monthly values are not free")

    sd = read_csv(RAW / "sprott" / "sput_daily.csv")
    for col, key, unit in [("SPUT price", "price_usd", "USD"), ("SPUT NAV per unit", "nav_usd", "USD"), ("SPUT premium/discount", "premium_discount_pct", "percent")]:
        data[col] = month_end(sd, "date", key)
        meta[col] = ("sput_daily", "Sprott", unit, "last trading day of the month")
    data["SPUT holdings"] = month_end(read_csv(RAW / "sprott" / "sput_holdings.csv"), "date", "pounds_u3o8_millions")
    meta["SPUT holdings"] = ("sput_holdings", "Sprott", "million lb U3O8", "last weekly value of the month")
    acq = defaultdict(float)
    for r in read_csv(RAW / "sprott" / "sput_acquired.csv"):
        acq[ym(r["date"])] += float(r["pounds_acquired"])
    data["SPUT pounds bought in month"] = {k: acq.get(k, 0.0) for k in data["SPUT price"]}
    meta["SPUT pounds bought in month"] = ("sput_acquired", "Sprott", "lb U3O8", "sum of purchases in the month (0 = none)")

    for sid_col, sid, unit in FRED_DAILY:
        data[sid_col] = month_end(read_csv(RAW / "fred" / f"{sid}.csv"), "observation_date", sid)
        meta[sid_col] = (sid, "FRED", unit, "last available daily value of the month")
    pink = read_pink()
    for col, pname, unit in PINK:
        data[col] = pink[pname]
        meta[col] = (pname, "World Bank Pink Sheet", unit, "monthly average as published")
    for tick, cur in EQUITIES:
        col = f"{tick} (adj. close)"
        data[col] = month_end(read_csv(RAW / "yahoo" / f"{tick}.csv"), "date", "adj_close")
        meta[col] = (tick, "Yahoo Finance", cur, "adjusted close, last trading day of the month; listing currency by ticker")

    last = max(data["Uranium spot (Cameco)"])
    months = [(y, m) for y in range(START[0], last[0] + 1) for m in range(1, 13) if START <= (y, m) <= last]
    cols = list(data)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Monthly"
    font, bold = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)
    head_fill = PatternFill("solid", fgColor="D9E2D5")
    ws.append(["Month"] + cols)
    for y, m in months:
        ws.append([dt.date(y, m, 1)] + [data[c].get((y, m)) for c in cols])
    for row in ws.iter_rows():
        for cell in row:
            cell.font = font
    for cell in ws[1]:
        cell.font, cell.fill = bold, head_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 1).number_format = "yyyy-mm"
        for c in range(2, len(cols) + 2):
            ws.cell(r, c).number_format = "#,##0.00"
    ws.column_dimensions["A"].width = 10
    for c in range(2, len(cols) + 2):
        ws.column_dimensions[get_column_letter(c)].width = 14
    ws.row_dimensions[1].height = 54
    ws.freeze_panes = "B2"

    src = wb.create_sheet("Sources")
    src.append(["Column", "Source series", "Source", "Unit", "Monthly value", "First month", "Last month", "Months with a value"])
    for c in cols:
        have = sorted(k for k in months if k in data[c])
        s, so, u, t = meta[c]
        src.append([c, s, so, u, t, f"{have[0][0]}-{have[0][1]:02d}" if have else "", f"{have[-1][0]}-{have[-1][1]:02d}" if have else "", len(have)])
    src.append([])
    notes = [
        "One source per series; sources are not mixed inside a column (see DATA_LOG.md).",
        "Main uranium price is the Cameco month-end spot (industry average of UxC and TradeTech month-end prices). Daily series use the last value of each month to match it.",
        "Pink Sheet commodities are monthly averages (no month-end values are published) and are used as published.",
        "The IMF uranium series is kept only as a cross-check; it runs about 20% below Cameco from 2021 (cause not confirmed).",
        "Stooq futures are missing Oct 2021 to Mar 2022 and Apr to Aug 2023; Investing fills Jan to Mar 2022 and a reconstructed UxC-benchmark column fills Oct to Dec 2021 and Apr to Aug 2023, each in its own column.",
        "Yellow Cake NAV and holdings and the conversion prices are quarterly: they appear only in quarter-end months (Mar, Jun, Sep, Dec) and are blank in between. They are not interpolated.",
        "Blank cells mean the source has no value for that month. SPUT starts in Jul 2021; equities start at their listing dates.",
        "Equity listing currencies are assigned by ticker; check before comparing across tickers.",
        f"The panel ends at {last[0]}-{last[1]:02d}, the last month with a Cameco spot price.",
        "Built by scripts/build_monthly.py from the files in data/raw/.",
    ]
    for n in notes:
        src.append([n])
    for row in src.iter_rows():
        for cell in row:
            cell.font = font
    for cell in src[1]:
        cell.font, cell.fill = bold, head_fill
    for col, w in zip("ABCDEFGH", (40, 20, 24, 18, 58, 12, 12, 16)):
        src.column_dimensions[col].width = w
    src.freeze_panes = "A2"

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"Saved {OUT} with {len(months)} months and {len(cols)} series")


if __name__ == "__main__":
    main()
