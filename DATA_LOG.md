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

## Yellow Cake report series, physical-market series, macro series, UPC NAV (8 Nov 2026)
- `data/raw/yellowcake/yellowcake_report_series_verified.csv`: 48 reports (32 quarterly updates, 8 interim, 8 annual results), every number re-read from the company PDFs (pdf.js text). Seven corrections to the earlier web-text pass (buyback cost in GBP, a pro-forma net-assets figure, one missing annual-report row, rounded cash values). The first-pass file `yellowcake_report_series_FIRSTPASS_unverified.csv` is superseded; delete it.
- `data/raw/physical/` (15 files): WNA and NEA/IAEA Red Book production by country, Red Book requirements (2022 actual, projections after), EIA production 1993-2010, EIA UMAR 2015 tables (1994-2015 purchases, enrichment, trade, inventories), IAEA TECDOC-2030 inventories end 2021 (two regions read), Euratom EU stocks (2010, 2012), IAEA reactor data booklet selected years, PRIS home-page snapshot. Read through page text, NOT downloaded files: verify before quoting. PRIS annual series not retrieved.
- `data/raw/macro/` (38 files, see macro_index.csv): US GDP and growth, OECD/G7 industrial production, categorical EPU, GPR (monthly, daily, historical), TED spread and commercial-paper/T-bill/SOFR rates, Kilian index, Google Trends 'uranium ETF', energy uncertainty index (source dates corrupted; not used). Baltic Dry Index not retrieved.
- `data/raw/upc/upc_monthly_nav.csv`: 125 monthly releases, Dec 2010 to May 2021 (May 2018 missing); CAD; shares outstanding not printed.
- Not found: SPUT dated issuances before 12 May 2025; SPUT monthly units.
| 2026-11-08 | Uranium month-end spot and long-term price (Cameco industry average of UxC and TradeTech month-end prices), script run | https://www.cameco.com/invest/markets/uranium-price | monthly, 466 rows, 1988-01 to 2026-10 | data/raw/cameco/uranium_price_monthly.csv |
| 2026-11-08 | Yellow Cake NAV per share and uranium held (quarterly operating updates), script run | https://otp.tools.investis.com/clients/uk/yellow_cake/rns/regulatory-news.aspx | quarterly, 33 rows, 0 new, 0 announcements unread | data/raw/yellowcake/nav_quarterly.csv |
| 2026-11-08 | UF6 conversion prices (Cameco quarter-end tables), script run | Cameco quarterly reports | quarterly, 61 rows, 0 new, 17 reports skipped | data/raw/cameco/conversion_prices_quarterly.csv |
| 2026-11-08 | GC=F: Gold futures, COMEX front month, USD per troy oz | Yahoo Finance | daily, USD, 6572 rows, 2000-08-29 to 2026-11-06 | data/raw/markets/GC=F.csv |
| 2026-11-08 | SI=F: Silver futures, COMEX front month, USD per troy oz | Yahoo Finance | daily, USD, 6573 rows, 2000-08-29 to 2026-11-06 | data/raw/markets/SI=F.csv |
| 2026-11-08 | HG=F: Copper futures, COMEX front month, USD per lb | Yahoo Finance | daily, USD, 6576 rows, 2000-08-29 to 2026-11-06 | data/raw/markets/HG=F.csv |
| 2026-11-08 | PL=F: Platinum futures, NYMEX front month, USD per troy oz | Yahoo Finance | daily, USD, 6599 rows, 1997-10-29 to 2026-11-06 | data/raw/markets/PL=F.csv |
| 2026-11-08 | CL=F: WTI crude futures, NYMEX front month, USD per barrel | Yahoo Finance | daily, USD, 6580 rows, 2000-08-22 to 2026-11-06 | data/raw/markets/CL=F.csv |
| 2026-11-08 | BZ=F: Brent crude futures, ICE front month, USD per barrel | Yahoo Finance | daily, USD, 4798 rows, 2007-07-29 to 2026-11-06 | data/raw/markets/BZ=F.csv |
| 2026-11-08 | NG=F: Natural gas futures (Henry Hub), NYMEX front month, USD per MMBtu | Yahoo Finance | daily, USD, 6577 rows, 2000-08-29 to 2026-11-06 | data/raw/markets/NG=F.csv |
| 2026-11-08 | ^GSPC: S&P 500 index (full history) | Yahoo Finance | daily, USD, 14335 rows, 1970-01-02 to 2026-11-06 | data/raw/markets/^GSPC.csv |

## 2026-11-08, SPUT statements checked against the PDFs
- `data/raw/sprott/sput_statements.csv`: the 2024-06-30, 2025-06-30, 2025-12-31 and 2026-06-30 rows were re-read against the Sprott PDFs (Q2 2024, Q2 2025, Q4 2025 annual and Q2 2026 reports, sprott.com). All 40 fields compared; the 37 already filled all match. Three blank 2024-06-30 cells (NAV per unit 21.07, uranium oxide fair value 5,590,484, pounds held 65,511,826) were filled from the PDF. Field-by-field result: `data/raw/sprott/sput_statements_check.csv`. Flow figures are six-month year-to-date for interim rows and twelve-month for 2025-12-31.
- Variables sheet notes updated (V042, V098, D17; "needs fetch script" wording replaced for series already retrieved).
- Common-sample counts (Monthly sheet): uranium spot + URA, CCJ, WTI, gold, S&P 500 = 192 months (2010-11 to 2026-10); spot + SPUT price, volume, NAV, premium, pounds bought and holdings = 63 (2021-08 to 2026-10); spot + 9 macro controls = 285 (2003-01 to 2026-09); SPUT block + macro = 62 (2021-08 to 2026-09, limited by industrial production and the Kilian index, which end in September).

## 2026-11-08, equity zero-volume days and WNA production check
- Daily sheet: adjusted closes on days with volume 0 are now blank (Yahoo repeats the previous price on non-trading days). Raw rows stay in `data/raw/yahoo/` and the Prices OHLCV sheet. Effect: DYL.AX starts Jul 1999 (was Jan 1988, price about 1,150), PDN.AX Jan 1999, BOE.AX unchanged start with 977 fewer days.
- `data/raw/physical/wna_uranium_production_by_country.csv`: 70 values (world tU, world tU3O8, USA, China, India, Ukraine, Russia, 2015-2024) re-read on the WNA page; no differences. The 2021 tU3O8 total (46,374) is printed so at source. Red Book country rows sum to its totals in all four years; world totals within 1% of WNA.
- Known issue not yet fixed: U-UN.TO joins Uranium Participation Corp (to 23 Jul 2021) and SPUT (from 26 Jul 2021).

## 2026-11-08, partial-variable completion round
- `data/raw/sprott/sput_statements.csv`: added 2022-06-30 and 2023-06-30 interim rows (read from the Sprott interim PDFs; units x NAV per unit within 0.002%, units issued equals the change in year-end units) and 2026-03-31 (total NAV only, from Sprott's one-page quarterly portfolio disclosure). Sprott publishes only an annual and a June report; there are no Q1/Q3 reports and no downloadable history of units outstanding on the fund page.
- `data/raw/yellowcake/yellowcake_total_nav.csv`: total NAV as stated in each of the 48 reports (USD million), NAV per share in GBP, basis (period end or pro forma). Never computed from per-share figures. Odd cases noted in the file.
- `data/raw/physical/`: `nea_iaea_redbook2018_2022_uranium_production_history.csv` (2014-2021), `nea_iaea_redbook2018_2022_reactor_requirements.csv`, `eia_umar2024_inventories_2020_and_2024p.csv`, `eia_s3a_inventories_2016_2019_rounded.csv` (million lb, rounded: EIA's detailed 2016-2019 editions could not be opened). Red Book 2018 'Total to 2017' column does not sum (+165, India and Romania rows). WNA world requirements pages were read but left out: their year labels and one value that equals a Red Book projection are unresolved. 2013 US reactor-owner inventory corrected to 113,077 (Table 22 and the total agree).
- Not found: Red Book 2020 and earlier, WNA production 2004-2014.

## 2026-11-08 · U-UN.TO split into two columns
- The Yahoo U-UN.TO series was Uranium Participation Corp (UPC) up to 16 Jul 2021 and SPUT from 26 Jul 2021 (price jumps from about CAD 5 to CAD 11 with no return).
- Daily and Monthly now hold two columns: "UPC adj. close" (2006-05-09 to 2021-07-16, 3,813 observations, mapped to D44) and "U-UN.TO adj. close" (SPUT only, from 2021-07-26, V011). UPC level checked against UPC's own NAV per share (CAD 4.87 at 31 May 2021).
- Prices OHLCV keeps the raw Yahoo series unchanged. Gaps row marked Done.

## 2026-11-08 · SEDAR+ attempt for dated SPUT issuances
- sprott.com hosts only the 22 Jan 2026 shelf prospectus and the 26 Jan 2026 supplement (both already in the workbook, 12 May 2025 to 19 Jan 2026). The 2022 annual information form gives only a total (73,894,300 units sold through the ATM in 2022).
- sedarplus.ca returned 403 Forbidden to the browser pane; not worked around. Earlier shelf prospectuses (e.g. 7 Sep 2023) need a manual download.

## 2026-11-08 · IAEA PRIS world annual series (V103, V104)
- Source: PRIS Analytics Data Browser (pris-stats.iaea.org), Operational > Electricity production, table view, all countries, years 1970 to 2025. The page has no export, so the 56 rows were read from the on-screen table (each row seen in at least two overlapping views; one misread, 1989 capacity, was caught and corrected on a clean re-read: 314,043 MW(e)).
- Checks: generation equals the RDS-2 (2024 edition) world totals for 1995, 2000, 2005, 2010, 2015 and 2020 to two decimals (TWh); implied load factor 38% in 1970 rising to 79% in 2025, no outliers after 1970. Capacity is PRIS "operated capacity", not RDS-2 year-end net capacity (2015: 364,367 against 360,504 MW(e)). 2022 and 2023 generation differ from RDS-2 (revisions).
- File: data/raw/iaea_pris/iaea_pris_world_annual_1970_2025.csv; workbook: Annual sheet, two columns (V103, V104). V105 (capacity factor) and V106 (starts and closures) not retrieved: they sit in the "Performance indicators" and "Reactor status at year-end" browsers.

## 2026-11-08 · SEDAR+ reached through Chrome: SPUT prospectus shortlist (nothing downloaded yet)
- sedarplus.ca returns 403 to the built-in browser pane but works in the user's Chrome. SPUT profile: "Sprott Physical Uranium Trust" (000052841), 159 public documents (Documents search, profile selected, no filters).
- Documents that should hold "Prior Sales" tables for Nov 2020 to Jan 2026 (the 22 Jan 2026 prospectus and 26 Jan 2026 supplement, 12 May 2025 to 19 Jan 2026, are already in the workbook): ATM supplement 22 Nov 2021 (687 KB); ATM supplement 15 Feb 2023 (1,877 KB); final short form prospectus 7 Sep 2023 (422 KB; ATM supplement 8 Sep 2023, 1,138 KB); final short form prospectus 3 Jan 2024 (427 KB; ATM supplement 3 Jan 2024, 868 KB; ATM supplement 20 Nov 2023, 703 KB); amendment to ATM supplement 6 Dec 2024 (739 KB); prospectus supplement 18 Jun 2025 (1,246 KB).
- The results page re-renders after every tick and the automated clicks landed on wrong rows, and a click on a document link downloaded nothing, so no file was taken. The user's Downloads folder is readable from the device shell, so manually downloaded PDFs can be read directly.

## 8 Nov 2026 - SPUT dated issuances back to 16 Feb 2022 (SEDAR+ prospectus supplements)
- Source: four prospectus supplements the user downloaded by hand from SEDAR+ (SEDAR+ blocks automated downloads): 15 Feb 2023, 3 Jan 2024, amended and restated Dec 2024, 18 Jun 2025; plus the 22 Jan 2026 prospectus already in use. The "Prior Sales" and "Market price of trust units" tables were extracted with pdftotext and parsed by script.
- data/raw/sprott/sput_issuances.csv: 343 dated rows, 16 Feb 2022 to 19 Jan 2026 (column "source" gives the document). The 12-month windows of the documents overlap without a gap; overlaps (Jan-Feb 2023, 12 Dec 2023) agree row for row.
- Issue type: "at-the-market" where the document says all issues were under the sales agreement (Feb 2023, Dec 2024, Jun 2025 supplements); 39 rows from Mar to Dec 2023 are marked "not stated" because the Jan 2024 supplement does not say.
- Reconciliation to Sprott reports (build stops if it fails): dated units and price x units equal reported units and proceeds for H1 2023 (11,070,217 / USD 140.874 m), H1 2024 (10,743,500 / 243.129 m), FY2024 (14,097,400 / 310.621 m) and H1 2025 (13,506,362 / 232.774 m). Known differences: H2 2023 table is 33,603 units short of the reports (7,301,608 vs 7,335,211); 1 Jan to 15 Feb 2022 (18,642,400 units, implied by FY2022 73,894,300 minus table 55,251,900, and the same figure from H1 2022) is before the first table.
- data/raw/sprott/sput_tsx_monthly.csv: Feb 2022 to Dec 2025 (47 -> 48 months incl. Jan 2026 partial in raw). Average volumes checked against Yahoo U-U.TO and U-UN.TO daily volumes: all 47 months x 2 lines within 0.18%.
- Workbook: SPUT issuance sheet (343 rows + source column); Monthly "SPUT units issued in month" and "gross issuance value" for Mar 2022 to Dec 2025; reconstructed month-end units now Jun 2022 to Dec 2025 (anchored on the latest reported date); TSX high/low/volume columns Feb 2022 to Dec 2025. Gaps and coverage_map updated.
- Not retrieved: Nov 2021 supplement (would cover Aug 2021 to early 2022). The 7 Sep 2023 shelf has no filings in the SEDAR+ search and is not needed.

## 2026-11-09 - IAEA PRIS world capacity factors, 1970-2025 (V105)
- Source: PRIS Analytics Data Browser > Performance indicators, table view, all countries (pris-stats.iaea.org/browser/534a68a0-610d-4152-a3bf-11e80df5cd29). The page has no export; the year filter was set to four windows (1970-1980, 1981-1995, 1996-2010, 2011-2025) so every row fitted on screen, and the 56 rows were read from the screenshots.
- File: data/raw/iaea_pris/iaea_pris_world_performance_1970_2025.csv (EAF, EUF, XUF, FLR, LF, OF, PUF, UCF, UCL, UUF in %). Workbook, Annual sheet: EAF, load factor, unit capability factor, forced loss rate (four columns, V105).
- Checks: EAF + EUF = 100, UCF = 100 - PUF - UCL and EUF = PUF + UCL + XUF hold within rounding for all 56 years (this caught one misread, 1998 PUF 13.32 not 13.52). Load factor never exceeds operating factor except 1970 (31.48% operating factor as published; one-year, a handful of reactors). The table also shows operated capacity and generation; compared with the 8 Nov file, six values differed by one digit (1973, 1984, 1985, 1986, 1997, 2020); the 8 Nov values were kept (the 2020 one, 2,553,244,244 MWh, matches RDS-2) and these two columns were not taken from the 9 Nov read.
- V106 (starts and closures): PRIS Data Browser only gives year-end stocks (reactors and capacity by status); its event lists are per reactor. The RDS-2 report (2025 edition, Table 7 pp. 19-21, per a web-fetch summary) lists grid connections and reactors in operation for every year 1954-2024; Table 16 lists shut-down reactors with dates. Not retrieved yet: needs the 2026 RDS-2 PDF downloaded by hand (not on the sandbox allowlist).

## 2026-11-09 - World reactor events 1954-2025 from IAEA RDS-2 No. 46 (V106)
- Source: IAEA Reference Data Series No. 2, 46th edition (2026), downloaded by the user (RDS-2_46web.pdf); text extracted with pdftotext and parsed by script. Table 7 (construction starts, grid connections, reactors in operation, units and MW(e), 1954-2025); Table 16 (220 permanently shut down reactors with shutdown month).
- Files: data/raw/iaea_pris/iaea_rds2_2026_world_reactor_events_1954_2025.csv (72 years) and iaea_rds2_2026_permanently_shut_down_reactors.csv (220 rows, with the printed PDF row). Workbook Annual sheet: six columns (V106).
- Checks: Table 16 yields 220 unique reactor codes and net capacity 110,812 MW(e), exactly the totals the report prints; Table 7 operating units and MW(e) for 2015 and 2025 equal the figures already used from PRIS (360,504 MW(e); 413 reactors, 377,076 MW(e)). Sum of grid connections 1954-2025 is 657 units against 413 operating + 23 suspended + 220 shut down = 656 (one unit unexplained, possibly a reactor connected twice).
- Not an identity: operating units(t) = operating(t-1) + connections - shutdowns fails in 18 of 71 years (e.g. 2011: 436 + 6 - 13 = 429 against 414 reported) because PRIS moves reactors in and out of "suspended operation" (Japan after Fukushima). The operating count is kept as published; no reconciling column was invented.
- Blank cells in Table 7 (no event that year, e.g. no grid connection in 1955 and 2008, no construction start in 1995) are entered as 0.

## 2026-11-09 - Baltic Dry Index (D18): looked, not retrieved
- Stooq has no BDI symbol (^BDI and a name search both return "does not exist"). Investing.com (indices/baltic-dry-historical-data) has a daily/weekly/monthly table without sign-in for the default range (last value 2,805 on 6 Nov 2026), but widening the date range to 1985 triggers a sign-up pop-up; no account was created. D18 stays "Not retrieved". Options: sign in yourself and read the monthly table, or use another free source (e.g. a broker or shipping-index site) if one is found.

## 2026-11-09 - Baltic Dry Index (D18): retrieved (supersedes the entry above)
- Investing.com did not apply the 1985 date range (its download held only the last 21 daily rows). The full daily series (10,482 observations, 7 Jan 1985 to 9 Nov 2026) is embedded in the chart page of stock-marketdata.com (time-series/charts/commodity-bdi-baltic-dry-index.html); month-end values were read from that chart in the browser.
- `data/raw/baltic/baltic_dry_index_monthly.csv`: 503 months, Jan 1985 to Nov 2026 (Nov 2026 partial, left out of the workbook). Check: Jun to Oct 2026 equal the Investing.com monthly figures (2,501; 2,732; 3,186; 3,113; 2,714). The Monthly sheet column starts where the uranium price starts.
- Dates inside each month are not stored (last trading day assumed); the daily series is not stored. Use for thesis only: Baltic Exchange restricts redistribution.

## 2026-11-09 - Three open decisions applied (V076, V082, D17)
- V076: funding spread = commercial paper 90-day AA financial (DCPF3M) minus T-bill 3-month (DTB3), computed on days where both exist (6,754 daily values, 1997-01-02 to 2026-11-05; 357 month-end values). Added to Daily and Monthly as "derived". Not spliced to the TED spread, which stays as raw data to Jan 2022. Caveat: commercial paper vs LIBOR, discount-basis bill.
- V082: OECD industrial production growth (to 2026-07) is the main series. Check: growth computed from the level series equals the published growth series with a mean absolute difference of 0.28 points (max 1.19) over 579 overlapping months (different data vintages). Levels kept (end 2024-03).
- D17: S&P 500 is the global-equity proxy (US index, longest history); MSCI World not retrieved (licensed).
- Workbook now: 60 in workbook, 34 partly, 18 derived not computed, 11 not retrieved, 7 paid or unavailable.
