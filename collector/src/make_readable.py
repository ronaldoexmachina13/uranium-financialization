"""Build data/DATA_OVERVIEW.xlsx: a human-readable copy of every downloaded series.

One sheet per series with plain-language column names and units, plus a first
'README' sheet describing each sheet (source, unit, frequency, coverage, raw file).

This is a DISPLAY copy only:
* raw files in data/raw/ are read, never modified;
* values are shown as downloaded (prices rounded to 4 decimals to hide float noise,
  e.g. 273.8999938964844 -> 273.9); nothing is cleaned, filled or removed;
* the econometric dataset is built separately (Phases 7-10).
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from common import PROCESSED, RAW, ROOT
from fetchers import read_raw

OUT = ROOT / "data" / "DATA_OVERVIEW.xlsx"


def newest(pattern: str) -> Path:
    files = sorted(RAW.glob(pattern), key=lambda p: p.stat().st_mtime)
    if not files:
        raise FileNotFoundError(pattern)
    return files[-1]


def yahoo(pattern: str, name: str, unit: str, vol_unit: str | None) -> tuple[pd.DataFrame, Path]:
    path = newest(pattern)
    d = pd.read_csv(path)
    out = pd.DataFrame({"Date": pd.to_datetime(d["Date"].str[:10])})
    for c in ["Open", "High", "Low", "Close"]:
        out[f"{name} {c} ({unit})"] = d[c].round(4)
    if vol_unit:
        out[f"Volume ({vol_unit})"] = d["Volume"]
    return out, path


def fred(pattern: str, label: str) -> tuple[pd.DataFrame, Path]:
    path = newest(pattern)
    d = read_raw({"source": "fred"}, path)
    return pd.DataFrame({"Date": d["date"], label: d["value"]}), path


def build_sheets() -> list[dict]:
    S = []

    def add(sheet, what, source, unit, freq, notes, frame_path):
        df, path = frame_path
        S.append(dict(sheet=sheet, what=what, source=source, unit=unit, freq=freq, notes=notes,
                      df=df, raw=str(path.relative_to(ROOT)) if isinstance(path, Path) else path))

    # ---------------- uranium
    stooq = pd.read_csv(PROCESSED / "daily/uranium_futures_UX_price_daily__stooq.csv", parse_dates=["date"])
    add("Uranium_daily_Stooq", "URANIUM - CME/COMEX UxC U3O8 futures (UX), front-month daily settlement",
        "Stooq (UX.F), validated vs CME and Cameco", "USD per lb U3O8", "Daily",
        "Gaps: Oct-2021..Mar-2022 and Apr..Sep-2023 (no rows). Volume/Open interest are 0 in source = not available.",
        (pd.DataFrame({"Date": stooq["date"],
                       "Uranium Close / settlement (USD/lb)": stooq["close"],
                       "Uranium Open (USD/lb)": stooq["open"],
                       "Uranium High (USD/lb)": stooq["high"],
                       "Uranium Low (USD/lb)": stooq["low"],
                       "Change (USD/lb)": stooq["change"],
                       "Change (%)": stooq["change_pct"]}),
         "data/raw/uranium/stooq_ux_pages/ (116 pages)"))

    cam_path = newest("uranium/uranium_spot_monthly_Cameco__*.html")
    cam = read_raw({"source": "html", "id": "cameco-uranium-price"}, cam_path)
    add("Uranium_monthly_Cameco", "URANIUM - spot price, month-end (average of UxC and TradeTech)",
        "Cameco", "USD per lb U3O8", "Monthly",
        "June 2019 appears twice on the source page (identical). Do not redistribute (Cameco terms).",
        (pd.DataFrame({"Month": cam["month"].astype(str), "Date as published": cam["date_raw"],
                       "Uranium spot, month-end (USD/lb)": cam["value"],
                       "Uranium long-term price (USD/lb)": cam["long_term"]}), cam_path))

    add("Uranium_monthly_IMF", "URANIUM - IMF global price (monthly AVERAGE)", "IMF via FRED (PURANUSDM)",
        "USD per lb U3O8", "Monthly", "Monthly average, not month-end. Cross-check only.",
        fred("uranium/uranium_price_monthly_IMF__*.csv", "Uranium, IMF monthly average (USD/lb)"))

    # ---------------- commodities
    add("Gold", "GOLD - COMEX front-month futures", "Yahoo Finance (GC=F)", "USD per troy ounce", "Daily",
        "Continuous front-month; latest row may be an unfinished session.",
        yahoo("commodities/gold_futures__*.csv", "Gold", "USD/oz", "contracts"))
    add("Silver", "SILVER - COMEX front-month futures", "Yahoo Finance (SI=F)", "USD per troy ounce", "Daily",
        "Continuous front-month; latest row may be an unfinished session.",
        yahoo("commodities/silver_futures__*.csv", "Silver", "USD/oz", "contracts"))
    add("Copper", "COPPER - COMEX front-month futures", "Yahoo Finance (HG=F)", "USD per pound", "Daily",
        "Continuous front-month; latest row may be an unfinished session.",
        yahoo("commodities/copper_futures__*.csv", "Copper", "USD/lb", "contracts"))
    add("Oil_WTI", "CRUDE OIL (WTI) - NYMEX front-month futures (PRIMARY oil series)", "Yahoo Finance (CL=F)",
        "USD per barrel", "Daily", "Matches EIA contract 1 on 96% of days 2000-2024. Negative price on 2020-04-20.",
        yahoo("commodities/oil_WTI_futures__*.csv", "Oil WTI", "USD/bbl", "contracts"))
    p = newest("commodities/oil_WTI_futures_EIA_check__*.xls")
    d = read_raw({"source": "eia_xls"}, p)
    add("Oil_WTI_EIA_check", "CRUDE OIL (WTI) - NYMEX contract 1, official EIA (validation only)", "EIA (RCLC1)",
        "USD per barrel", "Daily", "Ends 2024-04-05 (EIA stopped publishing).",
        (pd.DataFrame({"Date": d["date"], "Oil WTI contract 1 (USD/bbl)": d["value"]}), p))
    p = newest("commodities/oil_WTI_spot_EIA_check__*.xls")
    d = read_raw({"source": "eia_xls"}, p)
    add("Oil_WTI_spot_check", "CRUDE OIL (WTI) - Cushing SPOT price (cross-check only, not futures)", "EIA (RWTC)",
        "USD per barrel", "Daily", "Spot, a different definition from futures.",
        (pd.DataFrame({"Date": d["date"], "Oil WTI spot Cushing (USD/bbl)": d["value"]}), p))

    # ---------------- macro
    p = newest("macro/VIX_index__*.csv")
    d = read_raw({"source": "cboe"}, p)
    add("VIX", "VIX - Cboe Volatility Index (official)", "Cboe", "index points", "Daily", "Close = level used.",
        (pd.DataFrame({"Date": d["date"], "VIX Close": d["value"], "VIX Open": d["open"],
                       "VIX High": d["high"], "VIX Low": d["low"]}), p))
    add("VIX_FRED_check", "VIX - copy from FRED (cross-check only)", "FRED (VIXCLS)", "index points", "Daily", "",
        fred("macro/VIX_index_FRED_check__*.csv", "VIX Close"))
    add("USD_Broad", "US DOLLAR - Broad trade-weighted index (PRIMARY)", "Federal Reserve via FRED (DTWEXBGS)",
        "index, Jan-2006 = 100", "Daily", "Blank cells = US holidays.",
        fred("macro/USD_broad_index__*.csv", "USD broad index"))
    df, p = yahoo("macro/USD_DXY_index__*.csv", "DXY", "index", None)
    add("USD_DXY", "US DOLLAR - ICE Dollar Index DXY (robustness)", "Yahoo Finance (DX-Y.NYB)",
        "index, Mar-1973 = 100", "Daily", "6-currency basket; not the same as the broad index.", (df, p))
    add("US10Y", "INTEREST RATE - 10-Year Treasury yield", "FRED (DGS10)", "percent", "Daily",
        "Blank cells = bond-market holidays.", fred("macro/US_10Y_treasury_yield__*.csv", "US 10-year yield (%)"))
    add("FedFunds", "INTEREST RATE - Effective Federal Funds Rate", "FRED (DFF)", "percent",
        "Daily (7 days/week)", "Weekends carry Friday's rate.", fred("macro/US_fed_funds_rate__*.csv", "Fed funds rate (%)"))

    # ---------------- physical-uranium funds
    df, p = yahoo("sput/uranium_fund_UPC_SPUT_Toronto__*.csv", "Price", "CAD", "units traded")
    df.insert(1, "Vehicle", pd.cut(df["Date"], [pd.Timestamp.min, pd.Timestamp("2021-07-16"),
                                                pd.Timestamp("2021-07-23"), pd.Timestamp.max],
                                   labels=["UPC", "placeholder (fake rows)", "SPUT"]).astype(str))
    add("UPC_SPUT_Toronto", "URANIUM FUND - Uranium Participation Corp (to 16-Jul-2021) then Sprott Physical Uranium Trust",
        "Yahoo Finance (U-UN.TO, Toronto)", "CAD per unit", "Daily",
        "19-23 Jul 2021 are fake rows (frozen price, zero volume) - to be dropped in cleaning.", (df, p))
    add("SPUT_US_OTC", "URANIUM FUND - Sprott Physical Uranium Trust, US over-the-counter (robustness)",
        "Yahoo Finance (SRUUF)", "USD per unit", "Daily", "",
        yahoo("sput/uranium_fund_SPUT_US_OTC__*.csv", "Price", "USD", "units traded"))

    # ---------------- CME bulletin
    c = pd.read_csv(PROCESSED / "daily/uranium_futures_UX_volume_OI_daily__CME.csv", parse_dates=["trade_date"])
    add("UX_CME_bulletin", "URANIUM FUTURES ACTIVITY - UX volume and open interest (official CME)",
        "CME Daily Bulletin (forward collection)", "contracts (1 contract = 250 lb)", "Daily",
        "Only from 25-Sep-2026 onward; grows daily via the scheduled task.",
        (pd.DataFrame({"Date": c["trade_date"], "Front contract": c["front_contract"],
                       "Front settlement (USD/lb)": c["front_settle"],
                       "Volume, all contracts (contracts)": c["ux_total_volume"],
                       "Open interest, all contracts (contracts)": c["ux_open_interest"]}),
         "data/raw/cme/daily_bulletin/"))
    return S


U = "Uranium_daily_Stooq"
BY_HYPOTHESIS = [
    # hypothesis, role, variable, sheet, status
    ("H1", "Dependent", "Uranium price (UX futures settlement)", U, "Have (gaps Oct-21..Mar-22, Apr..Sep-23)"),
    ("H1", "Explanatory", "Gold price", "Gold", "Have"),
    ("H1", "Explanatory", "Oil price (WTI)", "Oil_WTI", "Have"),
    ("H1", "Explanatory", "Copper price", "Copper", "Have"),
    ("H1", "Explanatory", "Silver price", "Silver", "Have"),
    ("H1", "Robustness", "Uranium spot, monthly (no gaps)", "Uranium_monthly_Cameco", "Have"),
    ("H2", "Dependent", "Uranium price (UX futures settlement)", U, "Have (same as H1; lags built in R)"),
    ("H2", "Explanatory", "Gold, Oil, Copper, Silver (lagged)", "Gold / Oil_WTI / Copper / Silver", "Have"),
    ("H2", "Robustness", "Uranium spot, monthly", "Uranium_monthly_Cameco", "Have"),
    ("H3", "Dependent", "Uranium price (UX futures settlement)", U, "Have"),
    ("H3", "Explanatory", "Gold, Oil, Copper, Silver", "Gold / Oil_WTI / Copper / Silver", "Have"),
    ("H3", "Participation I_t (planned)", "UX futures volume + open interest", "UX_CME_bulletin", "MISSING 2007-2026 (only from 25-Sep-2026) - CME/FEP emails"),
    ("H3", "Participation I_t (fallback)", "UPC -> SPUT traded value (price x units)", "UPC_SPUT_Toronto", "Have (CAD; CAD/USD rate still to add)"),
    ("H3", "Robustness", "Post-SPUT dummy (from 19-Jul-2021)", "(built from dates)", "Built in Phase 9"),
    ("H4", "Dependent", "Uranium price (UX futures settlement)", U, "Have"),
    ("H4", "Explanatory", "VIX", "VIX", "Have"),
    ("H4", "Explanatory", "US dollar, broad index", "USD_Broad", "Have"),
    ("H4", "Explanatory", "Interest rate: 10-year yield", "US10Y", "Have"),
    ("H4", "Explanatory", "Interest rate: Fed funds", "FedFunds", "Have (alternative to 10Y)"),
    ("H4", "Participation I_t", "Same as H3", "UX_CME_bulletin / UPC_SPUT_Toronto", "As H3"),
    ("H4", "Robustness", "US dollar, DXY", "USD_DXY", "Have"),
    ("Checks only", "Validation (not in models)", "VIX from FRED; EIA oil contract 1 and spot; IMF uranium; SPUT US OTC",
     "VIX_FRED_check / Oil_WTI_EIA_check / Oil_WTI_spot_check / Uranium_monthly_IMF / SPUT_US_OTC", "Have"),
]


def write(sheets: list[dict]) -> None:
    rows = []
    for s in sheets:
        dcol = "Date" if "Date" in s["df"] else "Month"
        rows.append({"Sheet": s["sheet"], "What it is": s["what"], "Unit": s["unit"],
                     "Frequency": s["freq"], "Source": s["source"],
                     "From": str(s["df"][dcol].min())[:10], "To": str(s["df"][dcol].max())[:10],
                     "Rows": len(s["df"]), "Notes": s["notes"], "Raw file": s["raw"]})
    readme = pd.DataFrame(rows)
    with pd.ExcelWriter(OUT, engine="openpyxl", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as xw:
        pd.DataFrame({"": ["Readable copy of all downloaded series - one sheet per series.",
                           "Values as downloaded (prices rounded to 4 decimals); NOT cleaned: nothing filled, removed or adjusted.",
                           "Raw files in data/raw/ are the authoritative copies. Rebuild: python src/make_readable.py",
                           ""]}).to_excel(xw, sheet_name="README", index=False, header=False)
        readme.to_excel(xw, sheet_name="README", index=False, startrow=4)
        raw_of = {s["sheet"]: s["raw"] for s in sheets}
        pd.DataFrame([{"Hypothesis": h, "Role": r, "Variable": v, "Sheet in this file": sh,
                       "Raw file / folder": " | ".join(raw_of[x.strip()] for x in sh.split("/") if x.strip() in raw_of),
                       "Status": st} for h, r, v, sh, st in BY_HYPOTHESIS]
                     ).to_excel(xw, sheet_name="By_Hypothesis", index=False)
        for s in sheets:
            s["df"].to_excel(xw, sheet_name=s["sheet"][:31], index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A6" if ws.title == "README" else "A2"
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(12, width + 2), 70)
    print(f"wrote {OUT.relative_to(ROOT)} with {len(sheets)} series sheets")
    print(readme[["Sheet", "From", "To", "Rows"]].to_string(index=False))


if __name__ == "__main__":
    write(build_sheets())
