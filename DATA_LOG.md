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
| 2026-11-08 | SPUT price, NAV and premium/discount (USD, U.U) | Sprott fund page | daily, 1332 rows, 2021-07-19 to 2026-11-06 | data/raw/sprott/sput_daily.csv |
| 2026-11-08 | SPUT pounds of U3O8 held | Sprott fund page | weekly, 271 rows, 2021-07-16 to 2026-11-06 | data/raw/sprott/sput_holdings.csv |
| 2026-11-08 | SPUT pounds acquired and cumulative | Sprott fund page | purchase days, 259 rows, 2021-07-19 to 2026-07-30 | data/raw/sprott/sput_acquired.csv |
| 2026-11-08 | URA: Global X Uranium ETF | Yahoo Finance | daily, USD, 4025 rows, 2010-11-05 to 2026-11-06 | data/raw/yahoo/URA.csv |
| 2026-11-08 | URNM: Sprott Uranium Miners ETF | Yahoo Finance | daily, USD, 1741 rows, 2019-12-04 to 2026-11-06 | data/raw/yahoo/URNM.csv |
| 2026-11-08 | NLR: VanEck Uranium and Nuclear ETF | Yahoo Finance | daily, USD, 4839 rows, 2007-08-15 to 2026-11-06 | data/raw/yahoo/NLR.csv |
| 2026-11-08 | CCJ: Cameco (NYSE) | Yahoo Finance | daily, USD, 7713 rows, 1996-03-14 to 2026-11-06 | data/raw/yahoo/CCJ.csv |
| 2026-11-08 | KAP.IL: Kazatomprom GDR (London) | Yahoo Finance | daily, USD, 2017 rows, 2018-11-13 to 2026-11-06 | data/raw/yahoo/KAP.IL.csv |
| 2026-11-08 | U-U.TO: Sprott Physical Uranium Trust, USD units (TSX) | Yahoo Finance | daily, USD, 1327 rows, 2021-07-26 to 2026-11-06 | data/raw/yahoo/U-U.TO.csv |
| 2026-11-08 | U-UN.TO: Sprott Physical Uranium Trust, CAD units (TSX) | Yahoo Finance | daily, CAD, 5145 rows, 2006-05-09 to 2026-11-06 | data/raw/yahoo/U-UN.TO.csv |
| 2026-11-08 | YCA.L: Yellow Cake plc (London, pence) | Yahoo Finance | daily, GBp, 2109 rows, 2018-07-05 to 2026-11-06 | data/raw/yahoo/YCA.L.csv |
| 2026-11-08 | UEC: Uranium Energy Corp | Yahoo Finance | daily, USD, 4930 rows, 2007-04-05 to 2026-11-06 | data/raw/yahoo/UEC.csv |
| 2026-11-08 | DNN: Denison Mines (NYSE American) | Yahoo Finance | daily, USD, 5471 rows, 2005-02-09 to 2026-11-06 | data/raw/yahoo/DNN.csv |
| 2026-11-08 | NXE: NexGen Energy (NYSE) | Yahoo Finance | daily, USD, 3329 rows, 2013-08-14 to 2026-11-06 | data/raw/yahoo/NXE.csv |
| 2026-11-08 | PDN.AX: Paladin Energy (ASX) | Yahoo Finance | daily, AUD, 8333 rows, 1994-03-31 to 2026-11-06 | data/raw/yahoo/PDN.AX.csv |
| 2026-11-08 | BOE.AX: Boss Energy (ASX) | Yahoo Finance | daily, AUD, 4882 rows, 2007-07-25 to 2026-11-06 | data/raw/yahoo/BOE.AX.csv |
| 2026-11-08 | DYL.AX: Deep Yellow (ASX) | Yahoo Finance | daily, AUD, 9797 rows, 1988-01-29 to 2026-11-06 | data/raw/yahoo/DYL.AX.csv |
| 2026-11-08 | SLX.AX: Silex Systems (ASX) | Yahoo Finance | daily, AUD, 7247 rows, 1998-05-29 to 2026-11-06 | data/raw/yahoo/SLX.AX.csv |
| 2026-11-08 | BWXT: BWX Technologies | Yahoo Finance | daily, USD, 4093 rows, 2010-08-02 to 2026-11-06 | data/raw/yahoo/BWXT.csv |
| 2026-11-08 | 1164.HK: CGN Mining (Hong Kong) | Yahoo Finance | daily, HKD, 6144 rows, 2002-02-07 to 2026-11-06 | data/raw/yahoo/1164.HK.csv |
| 2026-11-08 | UX.F: uranium futures (Stooq) | Stooq history page (monthly), read from the page in the user's Chrome (downloads are blocked) | monthly, 224 rows, 2007-05 to 2026-11; 11 months missing at source (2021-10 to 2022-03, 2023-04 to 2023-08); volume 0 throughout; open interest 310 for Jan-Apr 2026, 0 otherwise; Nov 2026 is a partial month (to 6 Nov); dates are month-end labels; Stooq change columns checked against the closes, no mismatch | data/raw/stooq/ux_futures_monthly.csv |
| 2026-11-08 | Uranium futures, gap fill (Investing.com, front-month Dec 26 contract page) | Investing.com historical data, monthly, read from the page in the user's Chrome | monthly CLOSE only, 14 rows, 2022-01 to 2023-02; fills Jan-Mar 2022 that Stooq lacks; on the 11 months both have, closes agree in 6 and differ by at most 0.15 in 5; Investing's open/high/low are not usable (flat or inconsistent) and were not kept; still missing in both sources: Oct-Dec 2021 and Apr-Aug 2023 (8 months) | data/raw/investing/ux_futures_monthly.csv |
| 2026-11-08 | Uranium month-end spot and long-term price (Cameco, UxC) | cameco.com/invest/markets/uranium-price, table read from the page in the user's Chrome | monthly, 238 rows, 2007-01 to 2026-10; complete, no gaps; one exact duplicate row (2019-06) in the source removed; cross-check: futures close minus this spot averages 0.34 over 223 shared months | data/raw/cameco/uranium_price_monthly.csv |
