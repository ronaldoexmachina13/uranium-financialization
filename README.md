# Financialization and the Uranium Price

**Physical Investment Vehicles, Premium to Net Asset Value and Cross-Market Links**

Data and code for the Master's thesis of Ronaldo Silva, Mestrado em Contabilidade e Finanças, ISCAP (Politécnico do Porto). Supervisor: Prof. António Cunha. Co-supervisor: Prof.ª Patrícia Ramos. The title is the working title of the thesis work plan and may still change.

The thesis studies the uranium spot price and the funds that hold physical uranium. Sprott Physical Uranium Trust (SPUT, from July 2021) is at the centre. Uranium Participation Corporation (UPC, 2005 to 2021, the fund SPUT took over) and Yellow Cake plc (from July 2018) are used for comparison. Other commodities, financial markets, macro data, the physical uranium market and a dated list of market events complete the data set.

This repository is private and is shared with the supervisors for supervision only.

## Start here

Open **`data/processed/thesis_dataset.xlsx`** and read its first sheet, *Read me*. In short:

- The data set runs to a cut-off of **31 October 2026**. First-published values are cut by the date they were published, and the EIA's forward-looking figures for 2026 to 2035 are kept because they were published before the cut-off. It is version 6, built on 20 November 2026.
- There is one sheet per frequency (Monthly, Quarterly, Annual, Weekly, Daily), with one column per series. Record sheets hold one row per document or event (SPUT purchases and issuance, UPC and Yellow Cake reports, events and others). *Prices OHLCV* and *World Bank monthly* are source files kept whole.
- On the Monthly sheet the date shown is the 1st of the month, but the values are month-end values or values for the whole month.
- The *Variables* sheet lists every variable (V001 to V119, plus extra items D13 to D48) and, at the bottom, every column with its labels: how it may be used in a model, its first and last date, a timing code (month-end, average, flow and so on) and its link to the spot price (A: contains the spot by construction; B: partly linked; C: no link).
- The *Gaps* sheet lists what is missing and why.
- Light orange cells are values from an unofficial source, the public "SPUT Tracker" Google Sheet, which copies the Sprott fund page daily. They were checked against Sprott's documents wherever both exist and are graded C.
- The workbook holds inputs only. Returns, ratios and statistics are calculated separately, as part of the analysis.

## Folder guide

```
uranium_thesis/
├── README.md                this page
├── requirements.txt         Python packages needed to rebuild the data set
├── DATA_LOG.md              log of every download and every change to the data set
├── .gitattributes           tells Git to store every file byte for byte
│
├── data/
│   ├── processed/
│   │   ├── thesis_dataset.xlsx      the working data set
│   │   └── archive/                 earlier versions (v1 to v5) and the lists of changes between them
│   │
│   ├── raw/                         original downloads, never edited
│   │   ├── sprott/ upc/ yellowcake/           the physical uranium funds
│   │   ├── cameco/ uxc/ stooq/ cme/ investing/   uranium prices and futures
│   │   ├── markets_dated/                     commodity futures (Yahoo), correct trading dates
│   │   ├── markets/                           the same futures as first downloaded (kept as a record), and the S&P 500
│   │   ├── yahoo/ worldbank/ baltic/          equities and funds, commodity prices, freight
│   │   ├── fred/ alfred/ macro/               macro and financial series, first-published values
│   │   ├── physical/ eia/ iaea_pris/          production, reactors, contracting, inventories
│   │   └── events/ labels/                    dated market events, column labels
│   │
│   ├── raw_manifest.csv     every raw file with its SHA-256 fingerprint and source
│   ├── source_manifest.csv  the source document behind each data item
│   ├── release_dates.csv    the date each figure became public
│   └── spot_link.csv        which series contain the uranium spot price by construction
│
├── scripts/                 Python that downloads the data and builds the data set
├── collector/               CME uranium futures bulletin collector (stopped in November 2026, kept as a record)
└── reference/               code removed from the build, kept to document what was removed and why
```

Two notes on the raw folders. The futures files in `markets/` carry dates one day early in US summer-time months, so the build uses the corrected copies in `markets_dated/`; the originals are kept because raw files are never changed. The S&P 500 file in `markets/` is still used. The `collector/` folder holds the CME Daily Bulletin pages for uranium futures from March to 12 November 2026 (the earlier ones recovered from the Internet Archive), monthly CME volume and open-interest reports, 116 saved Stooq pages of UX futures prices, the collector's own September 2026 downloads of commodity, macro and SPUT series, and three screenshots of the bulletin and monthly-report columns used by the parser (`collector/data/raw/cme/*.png`).

## How the data are kept

- Raw files are never overwritten or edited. A correction is a new, dated file beside the original.
- No value is filled in, interpolated or invented. Outliers are flagged, not deleted.
- `data/raw_manifest.csv` lists every raw file with its size and SHA-256 fingerprint, so anyone can check that a file is unchanged.
- `data/source_manifest.csv` names the document (report, filing or saved web page) that supports each data item. Most of these documents are kept in the thesis folder on Google Drive; the CME bulletins and reports and the Stooq pages are in `collector/`.
- `data/release_dates.csv` gives the date each figure became public, so that a forecast never uses a figure before it was known.
- `DATA_LOG.md` records every download and every change to the data set, with its date.

## Versions

There is one working file, `data/processed/thesis_dataset.xlsx`. The version number and build date are on its *Read me* sheet. When a new version is built, the current file moves to `data/processed/archive/` under its version number (for example `thesis_dataset_v6.xlsx`), together with the list of changes between the two versions.

## Rebuilding the data set

From the repository root, with Python 3 installed:

```
python -m pip install -r requirements.txt
python scripts/build_dataset.py
```

The build reads the files in `data/raw/`, plus `data/release_dates.csv` and `data/spot_link.csv` for the labels, and writes `data/processed/thesis_dataset.xlsx` (about a minute and a half). To compare two versions cell by cell:

```
python scripts/compare_workbooks.py data/processed/archive/thesis_dataset_v5.xlsx data/processed/thesis_dataset.xlsx changes.csv
```

The download scripts (`scripts/fetch_*.py`) are only needed to collect new data. `scripts/fetch_alfred_first_release.py` needs a free FRED API key in a file named `.env` in the repository root (`FRED_API_KEY=...`); that file is excluded from the repository.

## Sources and licence

All data come from free public sources: Sprott, Cameco, Yellow Cake plc and UPC reports and filings (SEDAR+, SEC EDGAR); FRED and ALFRED (Federal Reserve Bank of St. Louis), including the IMF uranium price; the Federal Reserve Banks of Dallas and San Francisco; the World Bank; the IAEA, the NEA/IAEA Red Book, the World Nuclear Association, the Euratom Supply Agency and the US EIA; CME Group, Stooq, Yahoo Finance and Investing.com; the Baltic Dry Index (via stock-marketdata.com); Google Trends; published economic-policy-uncertainty and geopolitical-risk indices; and the public "SPUT Tracker" Google Sheet (grade C). Prices from UxC and TradeTech appear only as quoted in company reports.

Some of these sources restrict redistribution, among them CME Group (personal licence for the Daily Bulletin), Sprott, Investing.com and the Baltic Exchange. The repository is therefore private and shared for supervision only. Please do not redistribute its contents.

## Not in this repository

- The thesis documents, the literature and most source documents behind the data: they are in the thesis folder on Google Drive.
- Paid data: UxC and TradeTech price series, CME DataMine history, Bloomberg and LSEG. These are not available for the thesis.
- The FRED API key (`.env`), kept on the author's computer only.
- The analysis code. Notebooks will be added in `notebooks/` when the analysis starts.

Contact: Ronaldo Silva.
