"""Writes data/raw_manifest.csv: one row per raw file with hash, source, retrieval date and the kind of support behind it."""
import csv, hashlib, datetime as dt
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]; RAW = ROOT / "data" / "raw"
idx = {r["file"]: r for r in csv.DictReader(open(RAW / "macro" / "macro_index.csv", encoding="utf-8"))}
ORIG = "original download: the file is the support"
TYPED = "read from documents or pages and typed: source documents not yet kept"
KEPT = "read from documents and typed: source documents kept in 03_data/source_documents (partly, see source_manifest.csv)"
CHECKFILE = "my own check output (not a source): compares the typed SPUT statement values with the source reports; no separate document"
INDEXFILE = "index of the macro files (metadata, not data)"
LIVE = "read from a live page with no downloadable document: the page text of the 9 Nov 2026 reading is kept in 03_data/source_documents/iaea"
NOPAGE = "read from a page and typed: the page (Investing.com UX futures, monthly) could not be saved again (404); cross-checked against Stooq, see DATA_LOG.md"
NODOC = "read from the PRIS Data Browser on screen (no export, no document exists; the RDS-2 report does not print these series): not supportable by a saved document"
PARTIAL = "read from documents and typed: source documents kept for part of the values (Uranium 2022 Red Book edition is behind an NEA login, not kept), see source_manifest.csv"
def classify(p):
    d, f = p.parent.name, p.name
    if d == "fred": return ORIG, f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={p.stem}&cosd=1990-01-01"
    if d == "macro":
        if f == "macro_index.csv": return INDEXFILE, "list of the macro files and their sources"
        r = idx.get(f); return ORIG if r else TYPED, (r["source_url"] if r else "see macro_index.csv / DATA_LOG.md")
    if d in ("markets", "yahoo"): return ORIG, f"https://query1.finance.yahoo.com/v8/finance/chart/{p.stem}"
    if d == "worldbank": return ORIG, "https://www.worldbank.org/en/research/commodity-markets"
    if d == "cameco": return KEPT, ("https://www.cameco.com/invest/markets/uranium-price" if "uranium_price" in f else "Cameco quarterly MD&A and SEC filings, listed in scripts/fetch_conversion.py")
    if d == "yellowcake": return KEPT, "Yellow Cake regulatory news (otp.tools.investis.com/clients/uk/yellow_cake/rns) and company PDFs, kept in 03_data/source_documents/yellowcake (48 PDFs)"
    if d == "sprott":
        return (CHECKFILE if f == "sput_statements_check.csv" else KEPT if f in ("sput_issuances.csv", "sput_tsx_monthly.csv", "sput_prior_sales_2021-11-22_supplement.csv", "sput_statements.csv", "sput_daily.csv", "sput_holdings.csv", "sput_acquired.csv") else TYPED), "SPUT prospectus supplements (SEDAR+ 000052841) and sprott.com reports/fund page"
    if d == "iaea_pris": return (KEPT if "rds2" in f else NODOC), "IAEA RDS-2 (2026) PDF; PRIS Data Browser (Power BI, read from screen)"
    if d == "eia": return KEPT, "eia.gov uranium marketing annual report tables (UMAR 2025 and DUPR PDFs kept in 03_data/source_documents/eia)"
    if d == "physical": return (LIVE if "pris_world_snapshot" in f else PARTIAL if "redbook2018_2022" in f else KEPT), "WNA, NEA/IAEA Red Book, EIA UMAR 2015, IAEA TECDOC-2030, Euratom (see DATA_LOG.md)"
    if d == "baltic": return KEPT, "stock-marketdata.com/time-series/charts/commodity-bdi-baltic-dry-index.html (chart data, month-end), checked against investing.com/indices/baltic-dry-historical-data"
    if d in ("investing", "stooq"): return (NOPAGE if d == "investing" else KEPT), "investing.com / stooq.com UX futures pages (see DATA_LOG.md)"
    if d == "upc": return KEPT, "Uranium Participation Corporation monthly releases (125)"
    if d == "uxc": return KEPT, "reconstructed from UxC/Numerco benchmark (see DATA_LOG.md)"
    return TYPED, ""
rows = []
for p in sorted(RAW.rglob("*")):
    if p.is_file():
        sup, url = classify(p); b = p.read_bytes()
        rows.append([str(p.relative_to(ROOT)).replace("\\", "/"), len(b), hashlib.sha256(b).hexdigest(), sup, url,
                     dt.datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d")])
with open(ROOT / "data" / "raw_manifest.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f); w.writerow(["file", "bytes", "sha256", "support", "source", "file_date"]); w.writerows(rows)
import collections; print(len(rows), collections.Counter(r[3][:20] for r in rows))
