"""Build data/release_dates.csv: the date on which each delayed figure became public, so a predictive model never uses a figure before it was known.

19 Nov 2026: new script (labels approved by Ronaldo, 19 Nov 2026). It only reads what the repo already holds: the record sheets of the built workbook
(UPC monthly NAV, Yellow Cake reports, UxC volumes, SPUT issuance, SPUT statements, the Quarterly announcement-date column) and data/source_manifest.csv.
No date is estimated: where no date is stored, release_date is blank and the note says "not verified". No data value is read into or written to the workbook.

Usage: python3 scripts/release_dates.py [workbook]   (default data/processed/thesis_dataset_v5.xlsx)
"""
import csv
import datetime as dt
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "release_dates.csv"
FIELDS = ["series", "variable_id", "reference_period_end", "release_date", "release_date_source", "grade", "note"]


def d(x):
    if x is None or x == "":
        return None
    if isinstance(x, dt.datetime):
        return x.date()
    if isinstance(x, dt.date):
        return x
    return dt.date.fromisoformat(str(x)[:10])


def table(wb, name, first):
    """Rows of a record sheet as dicts, found by the text of its first header cell."""
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    h = next(i for i, r in enumerate(rows) if r and r[0] == first)
    heads = rows[h]
    return [dict(zip(heads, r)) for r in rows[h + 1:] if r and r[0] is not None]


def data_sheet(wb, name, keyhdr_start):
    """Series header (row 'Series') -> {key: value} of a data sheet; works before and after the Timing/Release rows were added."""
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    srow = next(i for i, r in enumerate(rows) if r and r[0] == "Series")
    krow = next(i for i, r in enumerate(rows) if i > srow and r and isinstance(r[0], str) and (r[0] == keyhdr_start or r[0].startswith(keyhdr_start + " (")))
    out = {}
    for j, s in enumerate(rows[srow]):
        if j == 0 or not s:
            continue
        out[s] = {d(r[0]): r[j] for r in rows[krow + 1:] if r[0] is not None and r[j] is not None}
    return out


def month_end(k):
    nxt = dt.date(k.year + (k.month == 12), k.month % 12 + 1, 1)
    return nxt - dt.timedelta(days=1)


def manifest_dates():
    return {r["file"]: r["document_date"] for r in csv.DictReader(open(ROOT / "data" / "source_manifest.csv", encoding="utf-8"))}


# SPUT issuance documents: label in the 'Source document' column -> (date, grade, where the date comes from).
# 19 Nov 2026: the date is that of the kept document that lists the row, read from source_manifest.csv where the document is listed there.
ISS_DOCS = {
    "Prospectus supplement (ATM) 22 Nov 2021": ("sedar_nov2021_mar2022/SPUT_ATM_prospectus_supplement_2021-11-22_SEDAR_03257748.pdf", "A"),
    "Prospectus supplement 15 Feb 2023": ("sput_prospectus_supplements/SPUT_prospectus_supplement_2023-02-15.pdf", "A"),
    "Prospectus supplement 3 Jan 2024": ("sput_prospectus_supplements/SPUT_prospectus_supplement_2024-01-03.pdf", "A"),
    "Amended and restated prospectus supplement Dec 2024": ("sput_prospectus_supplements/SPUT_amended_restated_prospectus_supplement_2024-12.pdf", "B"),
    "Prospectus supplement 18 Jun 2025": ("sput_prospectus_supplements/SPUT_prospectus_supplement_2025-06_underwritten.pdf", "A"),
}
BASE_SHELF_2026 = dt.date(2026, 1, 22)  # named in the document label; the base shelf itself is not a row of source_manifest.csv (the 26 Jan 2026 supplement is)

# Series with a typical publication lag but no stored release dates (Data Curator, report of 19 Nov 2026, section 2 and coverage_map_monthly.csv
# field availability_lag; general knowledge or Market Expert claims, never checked against a kept file).
CUR = "Data Curator report 2026-11-19, section 2 / coverage_map_monthly.csv availability_lag"
LAGS = [
    ("Uranium spot, month-end", "V001", "about 3 business days after month-end (Market Expert claim); the kept Cameco page copy is dated 9 Nov 2026 only. Route: Wayback captures of the Cameco price page", "C"),
    ("Uranium long-term, month-end", "V002", "about 3 business days after month-end (Market Expert claim); the kept Cameco page copy is dated 9 Nov 2026 only", "C"),
    ("Consumer price index (CPI-U, seasonally adjusted)", "V068", "about two to three weeks after the month (BLS); FRED keeps the latest vintage only. Route: ALFRED", "B"),
    ("US industrial production", "V081", "about two to three weeks after the month (Fed G.17); latest vintage only. Route: ALFRED", "B"),
    ("Industrial production, OECD total", "V082", "about six to ten weeks after the month; latest vintage only. Route: ALFRED", "B"),
    ("Industrial production, G7", "V082", "about six to ten weeks after the month; latest vintage only. Route: ALFRED", "B"),
    ("OECD industrial production growth, year on year", "V082", "about six to ten weeks after the month; latest vintage only. Route: ALFRED", "B"),
    ("Fed funds rate, monthly", "V062", "first business day of the next month", "B"),
    ("Baa corporate yield", "V070", "early in the next month", "B"),
    ("Aaa corporate yield", "V071", "early in the next month", "B"),
    ("Broad dollar index, real", "V078", "with the Fed H.10 monthly release", "C"),
    ("Kilian global real economic activity index", "V084", "about one to two months", "C"),
    ("Coal, Australia (IMF), monthly avg", "V045", "early in the following month (the IMF columns stop at Sep 2026 while the Pink Sheet has Oct 2026)", "B"),
    ("Natural gas, Henry Hub (IMF), monthly avg", "V051", "early in the following month", "B"),
    ("Copper (IMF), monthly avg", "V054", "early in the following month", "B"),
    ("Uranium (IMF), monthly avg", "", "early in the following month", "B"),
    ("World Bank monthly (all columns)", "V045-V055", "Pink Sheet early in the next month", "B"),
    ("Google Trends search interest, 'uranium ETF'", "D21", "near real time, but the whole series is rescaled at each fetch", "B"),
    ("Monetary-policy surprise, raw, sum over FOMC events in month", "V075", "research dataset; vintage ends Dec 2023; publication dates of its updates not stored", "B"),
    ("Monetary-policy surprise, orthogonalized, sum over FOMC events in month", "V075", "research dataset; vintage ends Dec 2023; publication dates of its updates not stored", "B"),
    ("Real GDP, US", "V083", "no typical lag stated in the Curator report; latest vintage only. Route: ALFRED", "C"),
    ("Real GDP growth, US, annualised", "V083", "no typical lag stated in the Curator report; latest vintage only. Route: ALFRED", "C"),
    ("Conversion spot, North America", "V007", "Cameco quarterly report: filing month only in source_manifest.csv (Q1 in May, Q2 in Aug, Q3 in Nov for 2007 to 2013); exact filing dates not stored. Route: EDGAR index pages", "B"),
    ("Conversion spot, Europe", "V007", "as for Conversion spot, North America", "B"),
    ("Conversion long-term, North America", "V008", "as for Conversion spot, North America", "B"),
    ("Conversion long-term, Europe", "V008", "as for Conversion spot, North America", "B"),
    ("World spot market volume, quarter, as reported in Cameco's MD&A", "V101", "Cameco MD&A: filing month only in source_manifest.csv; exact filing dates not stored. Route: EDGAR index pages", "B"),
]
EPU = ["Economic policy uncertainty, US", "Economic policy uncertainty, global"] + [f"Categorical EPU, {x}" for x in (
    "monetary policy", "fiscal policy (taxes or spending)", "taxes", "government spending", "health care", "national security", "entitlement programs",
    "regulation", "financial regulation", "trade policy", "sovereign debt and currency crises")]
GPR = [("Geopolitical risk index, overall (GPR)", "V089"), ("Geopolitical risk index, threats (GPRT)", "V090"), ("Geopolitical risk index, acts (GPRA)", "V090"),
       ("Geopolitical risk index, historical overall (GPRH)", "V089")]
LAGS += [(s, "V086" if "Categorical" not in s else "V088", "within the first days of the next month", "B") for s in EPU]
LAGS += [(s, v, "within the first days of the next month", "B") for s, v in GPR]
LAGS += [(s, "", "no typical lag stated in the Curator report", "C") for s in ("Energy uncertainty index, global, GDP-weighted (EUI)", "Energy uncertainty index, global, equally weighted (EUI)")]


def main():
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data" / "processed" / "thesis_dataset_v5.xlsx"
    wb = openpyxl.load_workbook(src, read_only=True)
    man = manifest_dates()
    out = []

    def add(series, vid, ref, rel, source, grade, note=""):
        out.append({"series": series, "variable_id": vid, "reference_period_end": ref.isoformat() if ref else "", "release_date": rel.isoformat() if rel else "",
                    "release_date_source": source, "grade": grade if rel else "", "note": note})

    # 1. UPC monthly NAV releases
    for r in table(wb, "UPC monthly NAV", "Release date"):
        s_ = str(r["Source"] or "")
        if s_.startswith("Kept PDF"):
            g, n = "B", "Date as in the sheet, from a kept SEC filing; it may be the SEC filing date rather than the press-release date (not checked). " + s_
        elif s_.startswith("SEDAR+"):
            g, n = "A", "Not a monthly release: " + s_
        else:
            g, n = "A", ""
        add("UPC monthly NAV (all columns of the UPC monthly NAV sheet)", "D44", d(r["As of"]), d(r["Release date"]),
            "UPC monthly NAV sheet, col A 'Release date' (data/raw/upc/upc_monthly_nav*.csv release_date; document per row in col I)", g, n)

    # 2. Yellow Cake quarter-end NAV and holdings: announcement date on the Quarterly sheet
    qs = data_sheet(wb, "Quarterly", "Quarter end")
    ann = qs["Yellow Cake announcement date"]
    for hdr, vid in [("Yellow Cake NAV per share", "V042"), ("Yellow Cake uranium held", "V040"), ("Yellow Cake quarter-end spot used", "V042")]:
        for q in sorted(qs[hdr]):
            add(hdr, vid, q, d(ann.get(q)), "Quarterly, col 'Yellow Cake announcement date' (data/raw/yellowcake/nav_quarterly.csv announcement_date)", "A",
                "same date applies to the other Yellow Cake columns of the quarter (lb basis, corrected and exact lb: their documents are in the cell comments)" if hdr == "Yellow Cake uranium held" else
                ("" if ann.get(q) else "announcement date not stored; not verified"))

    # 3. Yellow Cake reports sheet (shares, capital raised, purchases, cash, total NAV, buybacks): one row per report
    seen = set()
    for r in sorted(table(wb, "Yellow Cake reports", "Announcement date"), key=lambda r: (d(r["Period end"]), d(r["Announcement date"]))):
        pe, an = d(r["Period end"]), d(r["Announcement date"])
        first = pe not in seen
        seen.add(pe)
        add("Yellow Cake reports (all columns of the Yellow Cake reports sheet)", "V035-V042", pe, an,
            "Yellow Cake reports sheet, col A 'Announcement date' (data/raw/yellowcake/yellowcake_report_series_verified.csv and yellowcake_total_nav.csv announcement_date)",
            "A" if str(r.get("Read from the PDF") or "").lower() == "yes" else "B",
            "first report for this period end" if first else "later report for a period end already reported (figures may be restated)")

    # 4. UxC spot volume: first quote of each period in the Yellow Cake updates (same rule as the build: rows flagged INCONSISTENT are skipped)
    ux = table(wb, "UxC volumes", "Announcement date")
    grp = {}
    for r in ux:
        if str(r["Flag"] or "").startswith("INCONSISTENT"):
            continue
        grp.setdefault((r["Period type"], d(r["Period start"]), d(r["Period end"])), []).append(r)
    for (pt, ps, pe), rs in sorted(grp.items(), key=lambda t: (t[0][2], t[0][0])):
        rs = sorted(rs, key=lambda r: d(r["Announcement date"]))
        series = {"month": "UxC spot volume, month, first reported", "quarter": "UxC spot volume, quarter, first reported"}.get(pt, f"UxC volumes sheet, {pt}")
        add(series, "V101", pe, d(rs[0]["Announcement date"]), "UxC volumes sheet, col A 'Announcement date', earliest quote of the period (data/raw/yellowcake/uxc_spot_volume_reported.csv)", "A",
            f"first quote ({rs[0]['Period']}); {len(rs) - 1} later quote(s) of the same period in the sheet" + (f"; period starts {ps}" if pt not in ("month", "quarter") else ""))

    # 5. SPUT dated issuances: date of the kept document that lists each issue date
    def doc_of(label):
        best = None
        for part in [p.strip() for p in str(label).split(";")]:
            if part.startswith("Base shelf prospectus 22 Jan 2026"):
                c = (BASE_SHELF_2026, "B", "base shelf prospectus 22 Jan 2026 (date from the document name in data/raw/sprott/sput_issuances.csv; the base shelf is not a row of source_manifest.csv, the 26 Jan 2026 supplement is)")
            else:
                key = next((k for k in ISS_DOCS if part.lower().startswith(k.lower()) or k.lower().startswith(part.lower())), None)
                if key is None and part.lower().startswith("amended and restated"):
                    key = "Amended and restated prospectus supplement Dec 2024"
                if key is None:
                    continue
                f, g = ISS_DOCS[key]
                c = (d(man[f]), g, f"03_data/source_documents/{f} (source_manifest.csv document_date)" + ("; the label gives the month only, the day is the manifest date" if g == "B" else ""))
            if best is None or c[0] < best[0]:
                best = c
        return best
    iss = table(wb, "SPUT issuance", "Issue date")
    by_day = {}
    for r in iss:
        k = d(r["Issue date"])
        c = doc_of(r["Source document"])
        if c is None:
            raise SystemExit(f"unknown issuance document: {r['Source document']}")
        if k not in by_day or c[0] < by_day[k][0]:
            by_day[k] = c
    CAVEAT = ("earliest kept document that lists the issue; an earlier filing not in the repo may exist, so the date is an upper bound. Units outstanding were also shown daily "
              "on the Sprott fund page, so the fact of an issue may have been visible sooner")
    for k in sorted(by_day):
        rel, g, s_ = by_day[k]
        add("SPUT issuance (dated rows of the SPUT issuance sheet)", "V022-V024", k, rel, "SPUT issuance sheet, col G 'Source document' -> " + s_, g, CAVEAT)
    ms = data_sheet(wb, "Monthly", "Month")
    for hdr, vid in [("SPUT units issued in month", "V022"), ("SPUT gross issuance value in month", "V024")]:
        for m in sorted(ms.get(hdr, {})):
            days = [k for k in by_day if (k.year, k.month) == (m.year, m.month)]
            if not days:
                # 19 Nov 2026: a month with no issue shows 0; the zero is known once a document covering the month is out. The Prior Sales windows have no gap
                # between them (build check), so the document listing the next issue after the month covers it: its date is an upper bound, never earlier than the truth.
                nxt = [k for k in by_day if k > month_end(m)]
                if ms[hdr][m] != 0 or not nxt:
                    add(hdr, vid, month_end(m), None, "", "", "no dated issuance row in this month; not verified")
                    continue
                rel, g, s_ = by_day[min(nxt)]
                add(hdr, vid, month_end(m), rel, "SPUT issuance sheet, document of the next dated issue after the month -> " + s_, "B",
                    f"no issue in this month (value 0); the zero is taken as known from the document that lists the next issue ({min(nxt)}), whose Prior Sales window covers the month "
                    "because the windows have no gap (build check); upper bound")
                continue
            last = max(days, key=lambda k: by_day[k][0])
            rel, g, s_ = by_day[last]
            add(hdr, vid, month_end(m), rel, "SPUT issuance sheet, latest first-listing document among the month's issue dates -> " + s_, g,
                "the month total is complete only when its last-listed issue date is published; " + CAVEAT + ". The reconstructed units outstanding column depends on these dates too")

    # 6. SPUT statements (annual, interim and other report dates): release dates stored only for FY2021
    fy21 = "sedar_nov2021_mar2022/SPUT_annual_MRFP_FY2021_2022-03-18_SEDAR_03352693.pdf"
    st = table(wb, "SPUT statements", "Period end")
    for r in st:
        pe = d(r["Period end"])
        for hdr, vid, field in [("SPUT total NAV, reported", "V019", "Total NAV [USD thousand]"), ("SPUT units outstanding, reported", "V021", "Units outstanding"),
                                ("SPUT cash, reported", "V026", "Cash [USD thousand]")]:
            if r.get(field) is None:
                continue
            if pe == dt.date(2021, 12, 31):
                add(hdr, vid, pe, d(man[fy21]), f"03_data/source_documents/{fy21} (SEDAR+ filing 03352693; source_manifest.csv document_date)", "B",
                    "the workbook cell is read from a later report (" + str(r["Source"]) + "); the FY2021 annual MRFP filed on this date carries the FY2021 figures (source_manifest.csv), so this is the first public date")
            elif pe == dt.date(2026, 1, 20):
                add(hdr, vid, pe, BASE_SHELF_2026, "SPUT statements sheet, Source: " + str(r["Source"]), "B", "date from the document name; the base shelf is not a row of source_manifest.csv")
            else:
                add(hdr, vid, pe, None, "", "", "report kept (" + str(r["Source"]) + "), but its filing or publication date is not stored (source_manifest.csv gives the period end); not verified. Route: SEDAR+ filing dates or PDF metadata")

    # 7. series with a typical lag only
    for s, v, lag, g in LAGS:
        add(s, v, None, None, CUR + f" (grade {g})", "", f"typical lag {lag}; not verified")

    with open(OUT, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(out)
    n_ok = sum(1 for r in out if r["release_date"])
    print(f"{OUT}: {len(out)} rows, {n_ok} dated, {len(out) - n_ok} not verified")


if __name__ == "__main__":
    main()
