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

One source per series, never mixed. Main uranium price: Cameco month-end spot (Cameco's industry average of the UxC and TradeTech month-end prices, not UxC alone). The IMF series (FRED) is a cross-check only: it tracks Cameco closely through 2020 (monthly averages versus month-end values) but runs about 20% below from 2021 (cause not confirmed). Macro series come from FRED. Gold, silver, platinum, copper, Australian coal, US natural gas, Brent and WTI come from the World Bank Pink Sheet.

Why: Pink Sheet and FRED differ for Australian coal (211 of 417 months differ by more than 5%) and US natural gas (44 months, largest in cold winters). Cause not checked. The Pink Sheet is used because it goes back to 1970 and the Source Map points to it. Provisional.

Timing: the main price is a month-end price, so daily series use the last available value of each month. The Pink Sheet commodities are monthly averages and are kept only as a cross-check (Monthly_Pink); month-end gold, silver, copper, platinum, oil and gas come from daily sources (FRED/EIA oil; Yahoo via scripts/fetch_markets.py). Futures: Stooq and Investing closes stay in separate columns.

Processed: data/processed/thesis_dataset.xlsx, built by scripts/build_dataset.py. One sheet per native frequency (Daily, Weekly, Monthly, Quarterly, Events_SPUT, Daily_OHLCV, Monthly_Pink) plus Constructed, Events, Variable_Coverage (the 119 Source Map records and 11 extra items, with their status), README, Catalogue, Analysis_Map and Gaps. Monthly uses month-end values for every series that has a daily source; period averages are used only where no daily source exists (coal, IMF cross-checks) and in Monthly_Pink. Replaces monthly_panel.xlsx.
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
| 2026-11-08 | UX final settlement, reconstructed from UxC benchmark (8 months Stooq and Investing lack) | Handoff note from the user's research (00_admin/uranium_futures_price_handoff.md) citing Ux Weekly, Cameco SEC filings and Yellow Cake documents; evidence links in the file | 8 monthly values: 2021-10 to 2021-12 and 2023-04 to 2023-08. Benchmark-derived, using UxC's reporting month-end (the last Monday), not exchange bulletins; CME bulletin check and equivalence with Stooq/Investing closes still open; April 2023 calendar month-end would be 53.85 | data/raw/uxc/ux_final_settlement_reconstructed.csv |
| 2026-11-08 | Uranium month-end spot and long-term price (Cameco industry average of UxC and TradeTech month-end prices) | cameco.com/invest/markets/uranium-price, table read from the page in the user's Chrome | monthly, 238 rows, 2007-01 to 2026-10; complete, no gaps; one exact duplicate row (2019-06) in the source removed; cross-check: Stooq futures close minus this spot averages 0.34 over 223 shared months | data/raw/cameco/uranium_price_monthly.csv |
| 2026-11-08 | Yellow Cake plc: NAV per share (GBP), uranium held, quarter-end spot | Quarterly Operating Update announcements (Yellow Cake regulatory news), read from the pages in the user's Chrome | QUARTERLY, 33 rows, 2018-09 to 2026-09, complete. lb_basis: stated = in the announcement; derived = holdings value divided by quarter-end spot (agrees with stated holdings and known deliveries). No monthly NAV is published: title searches for NAV, net asset value, monthly and valuation in the announcements return nothing | data/raw/yellowcake/nav_quarterly.csv |
| 2026-11-08 | UF6 conversion prices, spot and long-term, North America and Europe (Cameco industry average of TradeTech and UxC, USD per kgU as UF6) | Cameco quarterly management reports (SEC 6-K exhibits and cameco.com), table 'Industry prices at quarter end', read with the web fetch tool | QUARTERLY, 61 rows, no gaps from 2014-03 to 2026-09; earlier only scattered quarters (2010-03, 2010-06, 2011-03 to 2011-09, 2012-03 to 2012-09, 2013-03, 2013-06); nothing before 2010. Uranium prices in the same tables agree with data/raw/cameco/uranium_price_monthly.csv except Dec 2017 (23.75 here, 22.32 there). Enrichment (SWU) prices are not in these tables; UxC sells month-end conversion and SWU prices to subscribers only | data/raw/cameco/conversion_prices_quarterly.csv |
| 2026-11-08 | GC=F: Gold futures, COMEX front month, USD per troy oz | Yahoo Finance | daily, USD, 6572 rows, 2000-08-29 to 2026-11-06 | data/raw/markets/GC=F.csv |
| 2026-11-08 | SI=F: Silver futures, COMEX front month, USD per troy oz | Yahoo Finance | daily, USD, 6573 rows, 2000-08-29 to 2026-11-06 | data/raw/markets/SI=F.csv |
| 2026-11-08 | HG=F: Copper futures, COMEX front month, USD per lb | Yahoo Finance | daily, USD, 6576 rows, 2000-08-29 to 2026-11-06 | data/raw/markets/HG=F.csv |
| 2026-11-08 | PL=F: Platinum futures, NYMEX front month, USD per troy oz | Yahoo Finance | daily, USD, 6599 rows, 1997-10-29 to 2026-11-06 | data/raw/markets/PL=F.csv |
| 2026-11-08 | CL=F: WTI crude futures, NYMEX front month, USD per barrel | Yahoo Finance | daily, USD, 6580 rows, 2000-08-22 to 2026-11-06 | data/raw/markets/CL=F.csv |
| 2026-11-08 | BZ=F: Brent crude futures, ICE front month, USD per barrel | Yahoo Finance | daily, USD, 4798 rows, 2007-07-29 to 2026-11-06 | data/raw/markets/BZ=F.csv |
| 2026-11-08 | NG=F: Natural gas futures (Henry Hub), NYMEX front month, USD per MMBtu | Yahoo Finance | daily, USD, 6577 rows, 2000-08-29 to 2026-11-06 | data/raw/markets/NG=F.csv |
| 2026-11-08 | ^GSPC: S&P 500 index (full history) | Yahoo Finance | daily, USD, 14335 rows, 1970-01-02 to 2026-11-06 | data/raw/markets/^GSPC.csv |
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
| 2026-11-08 | DGS3MO: 3-month Treasury yield | FRED | daily, 9613 rows | data/raw/fred/DGS3MO.csv |
| 2026-11-08 | FEDFUNDS: Effective federal funds rate, monthly | FRED | monthly, 442 rows | data/raw/fred/FEDFUNDS.csv |
| 2026-11-08 | CPIAUCSL: Consumer price index, all urban consumers, seasonally adjusted | FRED | monthly, 441 rows | data/raw/fred/CPIAUCSL.csv |
| 2026-11-08 | BAA: Moody's Baa corporate bond yield | FRED | monthly, 442 rows | data/raw/fred/BAA.csv |
| 2026-11-08 | AAA: Moody's Aaa corporate bond yield | FRED | monthly, 442 rows | data/raw/fred/AAA.csv |
| 2026-11-08 | BAMLC0A0CM: ICE BofA US corporate option-adjusted spread (FRED now serves only about 3 years of this series) | FRED | daily, 794 rows | data/raw/fred/BAMLC0A0CM.csv |
| 2026-11-08 | RTWEXBGS: Broad US dollar index, real | FRED | monthly, 250 rows | data/raw/fred/RTWEXBGS.csv |
| 2026-11-08 | DEXUSUK: USD per GBP exchange rate | FRED | daily, 9609 rows | data/raw/fred/DEXUSUK.csv |
| 2026-11-08 | INDPRO: US industrial production index | FRED | monthly, 441 rows | data/raw/fred/INDPRO.csv |
| 2026-11-08 | USEPUINDXD: Economic Policy Uncertainty index, US, daily | FRED | daily, 13458 rows | data/raw/fred/USEPUINDXD.csv |
| 2026-11-08 | USEPUINDXM: Economic Policy Uncertainty index, US, monthly | FRED | monthly, 442 rows | data/raw/fred/USEPUINDXM.csv |
| 2026-11-08 | GEPUCURRENT: Global Economic Policy Uncertainty index, current-price GDP weights | FRED | monthly, 355 rows | data/raw/fred/GEPUCURRENT.csv |
| 2026-11-08 | DHHNGSP: Henry Hub natural gas spot price, USD per MMBtu (EIA) | FRED | daily, 7781 rows | data/raw/fred/DHHNGSP.csv |

## SPUT fundraising extracts (typed by hand from Sprott documents, 8 Nov 2026)

These three files are not downloaded by a script: figures were read from the documents named below and typed into CSV. `build_dataset.py` stops if the unit roll-forward or the H2 2025 proceeds reconciliation fails.

| Date | File | Source | Content | Check |
|---|---|---|---|---|
| 2026-11-08 | data/raw/sprott/sput_issuances.csv | SPUT short form base shelf prospectus, 22 Jan 2026, "Prior Sales" (library: SPUT2026 - Short Form Base Shelf Prospectus 22 January 2026) | 99 dated issuances, 12 May 2025 to 19 Jan 2026: price per unit (USD), units; one private placement (12 May 2025), one bought deal (20 Jun 2025), the rest at-the-market | H2 2025 units equal the change in reported units (27,695,562); H2 2025 price x units equals the reported proceeds (US$529.862 million) exactly. Units on 20 Jan 2026 (318,488,764) exceed 311,070,452 + the table's January units by 40,000, so about 40,000 units were issued on 20 Jan or are not listed |
| 2026-11-08 | data/raw/sprott/sput_statements.csv | SPUT Annual Report 2024 (PDF in library); Q2 2025, Q4 2025 and Q2 2026 reports on sprott.com (read through WebFetch, which summarises); prospectus capitalisation (20 Jan 2026) | Units outstanding, NAV, NAV per unit, TSX close, cash, uranium oxide fair value, pounds, units issued, proceeds, issue costs (cash flow), purchases, for Dec 2021 to Jun 2026 (9 dates; flows from 2023) | Units roll forward exactly across all six reports; Dec 2024 figures in the Q2 2025 report equal the Annual Report 2024; H2 2025 proceeds equal the prospectus table. The money lines of the 2025-26 reports were not re-read against the PDFs (the shell cannot reach sprott.com) |
| 2026-11-08 | data/raw/sprott/sput_tsx_monthly.csv | Same prospectus, "Market Price of Trust Units" (TMX data) | Monthly TSX high, low and average daily volume for U.U (USD) and U.UN (CAD), Jan 2025 to 20 Jan 2026 | Yahoo daily volumes for U-U.TO and U-UN.TO reproduce the monthly averages within 0.03% for every month of 2025 |

Not extracted: 2021-2022 cash and flows, earlier prospectuses' prior-sales tables, Sprott monthly reports.

## Yellow Cake UxC spot volumes (typed by hand from Yellow Cake updates, 8 Nov 2026)
- `data/raw/yellowcake/uxc_spot_volume_reported.csv`: 100 rows from 31 of 32 quarterly updates (Dec 2023 and Dec 2018 quote none). Every figure with period, update, UxC wording and footnote; revised figures are kept as separate rows. Q2 2025 (27.6 million lb) conflicts with later updates and is flagged INCONSISTENT; the quarterly series excludes it.

## EIA annual series (read from eia.gov tables, 8 Nov 2026)
- `data/raw/eia/`: Table 3 (US uranium production 2011-2025), Tables 22 and 23 (US inventories at year-end 2021-2025), Table 11 (unfilled requirements), Table 10 (contracted deliveries), Table 12 (maximum requirements). Withheld cells (W) are left as W in the raw files and blank in the workbook. IAEA PRIS, NEA/IAEA and WNA series are not retrieved.
