"""Compare two versions of the thesis workbook and list every difference (equality checks only, no calculation on the data).

20 Nov 2026: new script, written for the v5 -> v6 check. Usage, from the repo root:
    python3 scripts/compare_workbooks.py data/processed/archive/thesis_dataset_v5.xlsx data/processed/thesis_dataset.xlsx data/processed/archive/v5_to_v6_changes.csv

How it compares
  - Data sheets (one column per series): series are matched by their name (row 'Series'), rows by their date, so a column that
    moved does not count as a change. Every cell present in both files is compared. Values dropped by the cut-off (after 31 Oct 2026)
    are listed once per series, with how many and from when to when.
  - Record sheets (one row per record): columns are matched by header, rows by key columns (dates and identifiers).
  - Guide sheets (Read me, Contents, Variables, Gaps): cell by cell (Gaps and Read me only summarised when rewritten or new).
Big files: add a fourth argument (only these sheets) or a fifth (skip these sheets), comma-separated, e.g. '' "Prices OHLCV".
Output columns: sheet, cell_old, cell_new, series, row_key, what, old, new (the same layout as v4_to_v5_changes.csv).
"""
import csv
import datetime as dt
import sys
from collections import Counter, defaultdict

import openpyxl
from openpyxl.utils import get_column_letter

CUTOFF = dt.date(2026, 10, 31)
DATA = ("Monthly", "Quarterly", "Annual", "Weekly", "Daily", "World Bank monthly")
HEADER_ROWS = {4: "Role in the thesis", 5: "Variable", 6: "Why it matters", 8: "Source", 9: "Type · coverage", 10: "Timing", 11: "Release", 12: "Unit"}
TABLE_KEYS = {"SPUT purchases": (0,), "SPUT issuance": (0, 1, 2), "SPUT statements": (0, 1), "UxC volumes": (0, 1, 2, 3, 4, 5, 6), "Yellow Cake reports": (0, 1),
              "UPC monthly NAV": (1,), "Physical sources": (0, 1, 2), "Prices OHLCV": (0, 1)}


def norm(v):
    if isinstance(v, dt.datetime):
        return v.date() if v.time() == dt.time(0) else v
    return v


def key_date(k):
    if isinstance(k, dt.datetime):
        k = k.date()
    if isinstance(k, dt.date):
        if k.day == 1:  # Monthly and World Bank rows show the 1st: the month ends later
            nxt = dt.date(k.year + (k.month == 12), k.month % 12 + 1, 1)
            return nxt - dt.timedelta(days=1)
        return k
    return None


def fmt(v):
    v = norm(v)
    return v.isoformat() if isinstance(v, (dt.date, dt.datetime)) else ("" if v is None else str(v))


def rows_of(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


def data_sheet(rows):
    hdr = rows[6]
    series = {}
    dup = {x for x in hdr[1:] if x and list(hdr[1:]).count(x) > 1}  # same name twice (Annual IAEA units and MW(e)): told apart by the unit row
    for j, s in enumerate(hdr):
        if j == 0 or not s:
            continue
        if s in dup:
            s = f"{s} [{rows[11][j]}]"
        vals = {}
        for i in range(12, len(rows)):
            k = norm(rows[i][0])
            if k is None:
                continue
            v = rows[i][j] if j < len(rows[i]) else None
            if v not in (None, ""):
                vals[k] = (v, i + 1)
        heads = {lab: (rows[r - 1][j] if j < len(rows[r - 1]) else None) for r, lab in HEADER_ROWS.items()}
        series[s] = dict(col=j + 1, vals=vals, heads=heads)
    keys = [norm(rows[i][0]) for i in range(12, len(rows)) if rows[i][0] is not None]
    return series, keys


def compare_data(name, ro, rn, out, stats):
    so, ko = data_sheet(ro)
    sn, kn = data_sheet(rn)
    gone_rows = [k for k in ko if k not in set(kn)]
    if gone_rows:
        cut = all(key_date(k) and key_date(k) > CUTOFF for k in gone_rows) if not isinstance(gone_rows[0], int) else False
        out.append([name, "", "", "", f"{fmt(gone_rows[0])} to {fmt(gone_rows[-1])} ({len(gone_rows)} rows)", "rows removed" + (": after the 31 Oct 2026 cut-off" if cut else ""), "", ""])
    new_rows = [k for k in kn if k not in set(ko)]
    if new_rows:
        out.append([name, "", "", "", f"{fmt(new_rows[0])} to {fmt(new_rows[-1])} ({len(new_rows)} rows)", "rows added", "", ""])
    for s, o in so.items():
        if s not in sn:
            out.append([name, get_column_letter(o["col"]), "", s, "", f"column removed ({len(o['vals'])} values)", "", ""])
            stats["removed_columns"].append(f"{name}: {s}")
            continue
        n = sn[s]
        for lab, ov in o["heads"].items():
            if ov != n["heads"][lab]:
                out.append([name, get_column_letter(o["col"]), get_column_letter(n["col"]), s, lab, "header changed", fmt(ov), fmt(n["heads"][lab])])
        cut = []
        for k, (ov, r) in o["vals"].items():
            if k in n["vals"]:
                stats["compared"] += 1
                nv, rn_ = n["vals"][k]
                if ov != nv:
                    stats["changed"] += 1
                    out.append([name, f"{get_column_letter(o['col'])}{r}", f"{get_column_letter(n['col'])}{rn_}", s, fmt(k), "value changed", fmt(ov), fmt(nv)])
            else:
                kd = key_date(k) if not isinstance(k, int) else None
                if kd and kd > CUTOFF:
                    cut.append(k)
                else:
                    stats["dropped_other"] += 1
                    out.append([name, f"{get_column_letter(o['col'])}{r}", "", s, fmt(k), "value removed", fmt(ov), ""])
        if cut:
            stats["cut"] += len(cut)
            out.append([name, get_column_letter(o["col"]), get_column_letter(n["col"]), s, f"{fmt(min(cut))} to {fmt(max(cut))} ({len(cut)} values)", "values removed: after the 31 Oct 2026 cut-off", "", ""])
        added = [(k, v) for k, v in n["vals"].items() if k not in o["vals"]]
        for k, (nv, r) in added:
            stats["added_in_old_columns"] += 1
            out.append([name, "", f"{get_column_letter(n['col'])}{r}", s, fmt(k), "value added", "", fmt(nv)])
    for s, n in sn.items():
        if s not in so:
            ks = sorted(n["vals"], key=lambda x: (str(type(x)), x))
            stats["new_columns"].append(f"{name}: {s} ({len(ks)} values)")
            out.append([name, "", get_column_letter(n["col"]), s, f"{fmt(ks[0])} to {fmt(ks[-1])}" if ks else "", f"new column ({len(ks)} values)", "", ""])


def table(rows, hr=5):
    h = rows[hr - 1]
    return [x for x in h], [r for r in rows[hr:] if any(c not in (None, "") for c in r)]


def compare_table(name, ro, rn, out, stats):
    ho, bo = table(ro)
    hn, bn = table(rn)
    for a, lab in ((0, "A1"), (1, "A2")):
        if ro[a][0] != rn[a][0]:
            out.append([name, lab, lab, "", "", "title changed", fmt(ro[a][0]), fmt(rn[a][0])])
    keys = TABLE_KEYS.get(name, (0,))

    def keyed(body):
        d, seen = {}, Counter()
        for i, r in enumerate(body):
            k = tuple(norm(r[j]) for j in keys)
            seen[k] += 1
            d[k + (seen[k],)] = (r, i)
        return d
    do, dn = keyed(bo), keyed(bn)
    io_, in_ = {h: j for j, h in enumerate(ho) if h}, {h: j for j, h in enumerate(hn) if h}
    for h in io_:
        if h not in in_:
            nv = sum(1 for r in bo if r[io_[h]] not in (None, ""))
            stats["removed_columns"].append(f"{name}: {h}")
            out.append([name, get_column_letter(io_[h] + 1), "", h, "", f"column removed ({nv} values)", "", ""])
    for h in in_:
        if h not in io_:
            stats["new_columns"].append(f"{name}: {h}")
            out.append([name, "", get_column_letter(in_[h] + 1), h, "", "new column", "", ""])
    cut = []
    for k, (r, i) in do.items():
        if k not in dn:
            kd = next((key_date(norm(r[j])) for j in keys if isinstance(norm(r[j]), dt.date)), None)
            if name in ("UxC volumes", "Yellow Cake reports", "SPUT statements") and isinstance(norm(r[1]), dt.date):
                kd = norm(r[1]) if name != "UxC volumes" else norm(r[5])
            if kd and kd > CUTOFF:
                cut.append(k)
            else:
                out.append([name, f"A{i + 6}", "", "", " | ".join(fmt(x) for x in k[:-1]), "row removed", "", ""])
            continue
        rn_, i2 = dn[k]
        for h, j in io_.items():
            if h in in_:
                ov, nv = norm(r[j]) if j < len(r) else None, norm(rn_[in_[h]]) if in_[h] < len(rn_) else None
                if ov in ("",):
                    ov = None
                if nv in ("",):
                    nv = None
                stats["table_compared"] += 1
                if ov != nv:
                    stats["table_changed"] += 1
                    out.append([name, f"{get_column_letter(j + 1)}{i + 6}", f"{get_column_letter(in_[h] + 1)}{i2 + 6}", h, " | ".join(fmt(x) for x in k[:-1]), "value changed", fmt(ov), fmt(nv)])
    if cut:
        stats["cut_rows"] += len(cut)
        out.append([name, "", "", "", f"{len(cut)} rows", "rows removed: after the 31 Oct 2026 cut-off", "", ""])
    newk = [k for k in dn if k not in do]
    for k in newk:
        out.append([name, "", f"A{dn[k][1] + 6}", "", " | ".join(fmt(x) for x in k[:-1]), "row added", "", ""])


def compare_cells(name, ro, rn, out, summary_only=False):
    n = 0
    R, C = max(len(ro), len(rn)), max(max((len(r) for r in ro), default=0), max((len(r) for r in rn), default=0))
    for i in range(R):
        for j in range(C):
            a = ro[i][j] if i < len(ro) and j < len(ro[i]) else None
            b = rn[i][j] if i < len(rn) and j < len(rn[i]) else None
            if a != b:
                n += 1
                if not summary_only:
                    out.append([name, f"{get_column_letter(j + 1)}{i + 1}", f"{get_column_letter(j + 1)}{i + 1}", "", "", "cell", fmt(a), fmt(b)])
    if summary_only and n:
        out.append([name, "", "", "", "", f"sheet rewritten: {n} cells differ", "", ""])
    return n


def main(old, new, path, only=None, skip=None):
    """only / skip: comma-separated sheet names (the big Prices OHLCV sheet can be compared in a separate run, then the CSVs joined)."""
    only = set(only.split(",")) if only else None
    skip = set(skip.split(",")) if skip else set()
    wo = openpyxl.load_workbook(old, read_only=True)
    wn = openpyxl.load_workbook(new, read_only=True)
    out = []
    stats = defaultdict(int)
    stats["removed_columns"], stats["new_columns"] = [], []
    for s in wo.sheetnames:
        if s not in wn.sheetnames and not only:
            out.append([s, "", "", "", "", "sheet removed", "", ""])
    for s in wn.sheetnames:
        if s not in wo.sheetnames and not only:
            out.append([s, "", "", "", "", "new sheet", "", f"{wn[s].max_row} rows x {wn[s].max_column} columns"])
    for s in wo.sheetnames:
        if s not in wn.sheetnames or (only and s not in only) or s in skip:
            continue
        ro, rn = rows_of(wo[s]), rows_of(wn[s])
        if s in DATA:
            compare_data(s, ro, rn, out, stats)
        elif s in ("Contents", "Variables"):
            stats[f"cells_{s}"] = compare_cells(s, ro, rn, out)
        elif s == "Gaps":
            stats["cells_Gaps"] = compare_cells(s, ro, rn, out, summary_only=True)
        else:
            compare_table(s, ro, rn, out, stats)
        print(f"  {s}: done", flush=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["sheet", "cell_old", "cell_new", "series", "row_key", "what", "old", "new"])
        w.writerows(out)
    print(f"{path}: {len(out)} rows")
    for k, v in stats.items():
        print(f"  {k}: {v}")
    return stats


if __name__ == "__main__":
    main(*sys.argv[1:4], *(sys.argv[4:6]))
