"""Build data/processed/monthly_panel.xlsx from the raw FRED and World Bank files.

Run from the repo root after fetch_fred.py and fetch_worldbank.py:  python scripts/build_monthly.py
Needs openpyxl.

Rules (also recorded in DATA_LOG.md):
- One source per series. Uranium and the macro series come from FRED; gold, silver, platinum, copper,
  coal, natural gas and oil come from the World Bank Pink Sheet.
- Timing: the uranium price is a monthly period average, so daily FRED series are averaged over the
  days of each month. The panel covers Jan 1992 to the last month with a uranium price.
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
FRED = ROOT / "data" / "raw" / "fred"
PINK = ROOT / "data" / "raw" / "worldbank" / "CMO-Historical-Data-Monthly.xlsx"
OUT = ROOT / "data" / "processed" / "monthly_panel.xlsx"

# (column name, source series, source, unit, how monthly value is made)
FRED_COLS = [
    ("Uranium", "PURANUSDM", "FRED (IMF)", "USD per lb", "monthly period average (as published)"),
    ("VIX", "VIXCLS", "FRED", "index points", "average of daily closes"),
    ("Dollar index (broad)", "DTWEXBGS", "FRED", "index", "average of daily values"),
    ("Fed funds rate", "DFF", "FRED", "percent", "average of daily values"),
    ("US 10y yield", "DGS10", "FRED", "percent", "average of daily values"),
    ("US 10y real yield", "DFII10", "FRED", "percent", "average of daily values"),
    ("Term spread 10y-3m", "T10Y3M", "FRED", "percentage points", "average of daily values"),
    ("Baa-10y spread", "BAA10Y", "FRED", "percentage points", "average of daily values"),
    ("Breakeven inflation 10y", "T10YIE", "FRED", "percent", "average of daily values"),
    ("USD/CAD", "DEXCAUS", "FRED", "CAD per USD", "average of daily values"),
    ("S&P 500", "SP500", "FRED", "index", "average of daily closes (FRED has only about 10 years)"),
]
PINK_COLS = [
    ("Gold", "Gold", "World Bank Pink Sheet", "USD per troy oz"),
    ("Silver", "Silver", "World Bank Pink Sheet", "USD per troy oz"),
    ("Platinum", "Platinum", "World Bank Pink Sheet", "USD per troy oz"),
    ("Copper", "Copper", "World Bank Pink Sheet", "USD per tonne"),
    ("Coal, Australian", "Coal, Australian", "World Bank Pink Sheet", "USD per tonne"),
    ("Natural gas, US", "Natural gas, US", "World Bank Pink Sheet", "USD per MMBtu"),
    ("Brent crude", "Crude oil, Brent", "World Bank Pink Sheet", "USD per barrel"),
    ("WTI crude", "Crude oil, WTI", "World Bank Pink Sheet", "USD per barrel"),
]


def read_fred(sid):
    """Return {(year, month): [values]} from a raw FRED csv."""
    out = defaultdict(list)
    with open(FRED / f"{sid}.csv", newline="", encoding="utf-8") as f:
        for row in list(csv.reader(f))[1:]:
            if len(row) > 1 and row[1] not in ("", "."):
                out[(int(row[0][:4]), int(row[0][5:7]))].append(float(row[1]))
    return out


def read_pink():
    ws = openpyxl.load_workbook(PINK, data_only=True)["Monthly Prices"]
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
    data = {}  # column name -> {(y, m): value}
    for name, sid, *_ in FRED_COLS:
        data[name] = {k: sum(v) / len(v) for k, v in read_fred(sid).items()}
    pink = read_pink()
    for name, pname, *_ in PINK_COLS:
        data[name] = pink[pname]

    last = max(data["Uranium"])
    months = [(y, m) for y in range(1992, last[0] + 1) for m in range(1, 13) if (y, m) <= last]
    cols = [c[0] for c in FRED_COLS[:1]] + [c[0] for c in PINK_COLS] + [c[0] for c in FRED_COLS[1:]]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Monthly"
    font, bold = Font(name="Arial", size=10), Font(name="Arial", size=10, bold=True)
    head_fill = PatternFill("solid", fgColor="D9E2D5")
    ws.append(["Month"] + [f"{c}" for c in cols])
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
    ws.row_dimensions[1].height = 42
    ws.freeze_panes = "B2"

    src = wb.create_sheet("Sources")
    src.append(["Column", "Source series", "Source", "Unit", "Monthly value", "First month", "Last month", "Months filled"])
    meta = {c[0]: (c[1], c[2], c[3], c[4]) for c in FRED_COLS}
    meta.update({c[0]: (c[1], c[2], c[3], "monthly period average (as published)") for c in PINK_COLS})
    for c in cols:
        have = sorted(k for k in months if k in data[c])
        s, so, u, t = meta[c]
        src.append([c, s, so, u, t, f"{have[0][0]}-{have[0][1]:02d}", f"{have[-1][0]}-{have[-1][1]:02d}", len(have)])
    src.append([])
    notes = [
        "One source per series; sources are not mixed (see DATA_LOG.md).",
        "Uranium is the IMF monthly period average. Daily series are averaged over each month's days to match it.",
        "The Pink Sheet has no uranium series. Blank cells mean the source has no value for that month.",
        f"The panel ends at {last[0]}-{last[1]:02d}, the last month with a uranium price.",
        "Built by scripts/build_monthly.py from the files in data/raw/.",
    ]
    for n in notes:
        src.append([n])
    for row in src.iter_rows():
        for cell in row:
            cell.font = font
    for cell in src[1]:
        cell.font, cell.fill = bold, head_fill
    for col, w in zip("ABCDEFGH", (26, 20, 24, 18, 44, 12, 12, 14)):
        src.column_dimensions[col].width = w
    src.freeze_panes = "A2"

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"Saved {OUT} with {len(months)} months and {len(cols)} series")


if __name__ == "__main__":
    main()
