# Data log

One line per download: date, series, source, frequency and price timing (period average or month-end), file.

| Date | Series | Source | Frequency / timing | File |
| --- | --- | --- | --- | --- |
| 2026-11-08 | PURANUSDM: Global price of uranium, USD per lb (IMF) | FRED | monthly, period average, 417 rows | data/raw/fred/PURANUSDM.csv |
| 2026-11-08 | PCOALAUUSDM: Coal, Australia, USD per tonne (IMF) | FRED | monthly, period average, 417 rows | data/raw/fred/PCOALAUUSDM.csv |
| 2026-11-08 | PCOPPUSDM: Copper, USD per tonne (IMF) | FRED | monthly, period average, 417 rows | data/raw/fred/PCOPPUSDM.csv |
| 2026-11-08 | PNGASUSUSDM: Natural gas, Henry Hub, USD per MMBtu (IMF) | FRED | monthly, period average, 417 rows | data/raw/fred/PNGASUSUSDM.csv |
| 2026-11-08 | DCOILWTICO: WTI crude oil, USD per barrel (EIA) | FRED | daily, 9611 rows | data/raw/fred/DCOILWTICO.csv |
| 2026-11-08 | DCOILBRENTEU: Brent crude oil, USD per barrel (EIA) | FRED | daily, 9611 rows | data/raw/fred/DCOILBRENTEU.csv |
| 2026-11-08 | VIXCLS: CBOE VIX | FRED | daily close, 9613 rows | data/raw/fred/VIXCLS.csv |
| 2026-11-08 | DTWEXBGS: Broad US dollar index, nominal | FRED | daily, 5435 rows | data/raw/fred/DTWEXBGS.csv |
| 2026-11-08 | DFF: Effective federal funds rate | FRED | daily, 13458 rows | data/raw/fred/DFF.csv |
| 2026-11-08 | DGS10: 10-year Treasury yield | FRED | daily, 9613 rows | data/raw/fred/DGS10.csv |
| 2026-11-08 | DFII10: 10-year real Treasury yield (TIPS) | FRED | daily, 6221 rows | data/raw/fred/DFII10.csv |
| 2026-11-08 | T10Y3M: 10-year minus 3-month Treasury spread | FRED | daily, 9614 rows | data/raw/fred/T10Y3M.csv |
| 2026-11-08 | BAA10Y: Moody's Baa corporate yield minus 10-year Treasury | FRED | daily, 9613 rows | data/raw/fred/BAA10Y.csv |
| 2026-11-08 | T10YIE: 10-year breakeven inflation | FRED | daily, 6222 rows | data/raw/fred/T10YIE.csv |
| 2026-11-08 | DEXCAUS: CAD per USD exchange rate | FRED | daily, 9609 rows | data/raw/fred/DEXCAUS.csv |
| 2026-11-08 | SP500: S&P 500 index (FRED keeps about 10 years only) | FRED | daily close, 2609 rows | data/raw/fred/SP500.csv |
| 2026-11-08 | World Bank Pink Sheet, monthly prices (gold, silver, coal, copper, oil, gas, etc.) | World Bank | monthly, period average, 573 KB | data/raw/worldbank/CMO-Historical-Data-Monthly.xlsx |

## Source rule (8 Nov 2026)

One source per series, never mixed. Uranium and the macro series come from FRED. Gold, silver, platinum, copper, Australian coal, US natural gas, Brent and WTI come from the World Bank Pink Sheet.

Why: Pink Sheet and FRED differ for Australian coal (211 of 417 months differ by more than 5%) and US natural gas (44 months, largest in cold winters). Cause not checked. The Pink Sheet is used because it goes back to 1970 and the Source Map points to it. Provisional.

Timing: uranium is an IMF monthly period average, so daily series are averaged over each month.

Processed: data/processed/monthly_panel.xlsx, built by scripts/build_monthly.py (417 months, Jan 1992 to Sep 2026, 19 series).
