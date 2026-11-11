# Uranium thesis — data acquisition pipeline

*When Nuclear Fuel Becomes a Financial Asset: Financialization and Price Discovery in the Uranium Market* (Master's thesis, ISCAP).

This repository covers **data engineering only**: acquiring, preserving, cleaning and documenting data. The econometric analysis happens separately in R.

## Setup

The virtual environment lives outside OneDrive so it doesn't sync thousands of package files:

```
python -m venv C:\Users\ronal\.venvs\uranium_thesis
C:\Users\ronal\.venvs\uranium_thesis\Scripts\python.exe -m pip install -r requirements.txt
```

- `requirements.txt` lists the minimum versions.
- `requirements-lock.txt` pins the exact versions used, for reproducibility.
- No API keys are required (FRED is read through its keyless CSV endpoint).

## Run order

Run these from the project root:

| Step | Command | Output |
|---|---|---|
| Macro series | `python src/download_macro_data.py` | `data/raw/macro/` |
| Commodities | `python src/download_market_data.py` | `data/raw/commodities/` |
| CME UX bulletin | `python src/download_cme_uranium.py` (`--parse` re-parses stored PDFs only) | `data/raw/cme/daily_bulletin/`, `data/processed/daily/cme_ux_*.csv` |

## Rules

- **Raw files are never overwritten.** Each download is saved as `<what>__<source>-<code>__<YYYYMMDD>.<ext>`, e.g. `gold_futures__yahoo-GC-F__20260929.csv`. An identical re-download is discarded. A changed file gets a time suffix (`..._HHMMSS`).
- Raw files downloaded before this naming scheme were renamed on 2026-09-29 with their contents unchanged. `logs/raw_file_renames.csv` maps each old name to its new one, with its SHA-256 hash. `logs/retrieval_log.csv` keeps the names that were current at download time.
- `data/DATA_OVERVIEW.xlsx` is a readable copy of every series, with one sheet per series and a README sheet. Rebuild it with `python src/make_readable.py`.
- Every retrieval attempt is logged in `logs/retrieval_log.csv`, with its URL, SHA-256 hash, byte size and status.
- No values are filled in, interpolated or invented. Outliers are flagged, never deleted.

## CME UX forward collection

CME publishes only the *current* Daily Bulletin for free; historical bulletins are sold through CME DataMine. A Windows scheduled task, `UraniumThesis_CME_UX_Bulletin`, runs the collector daily at 14:00 and 23:30, and catches up if the PC was off at those times. Any day on which the PC stays off throughout is lost permanently.

The collector uses two bulletin sections:

- **Section 62:** per-contract settlement, Globex volume, PNT/ClearPort volume and open interest.
- **Section 02B:** exchange summary, used as a cross-check.

The parser reads each word's position from the PDF text layer with `pdfplumber` (no OCR). It checks the page layout against the column headers and cross-checks the sum of contract open interest against both the TOTAL line and Section 02B.

CME's terms of use grant a personal, non-transferable licence. **Do not redistribute the raw PDFs.**

Manage the task with:

```
Get-ScheduledTaskInfo -TaskName UraniumThesis_CME_UX_Bulletin
Unregister-ScheduledTask -TaskName UraniumThesis_CME_UX_Bulletin
```

## Sources and status

See `data/source_audit.csv` for every source, its verified coverage, licensing notes and retrieval status.
