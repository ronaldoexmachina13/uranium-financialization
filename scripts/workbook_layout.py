"""Layout and styling of thesis_dataset.xlsx. Pure presentation: no data is loaded or changed here.

Every data column is tied to a record of the handoff table (handoff_table.py, IDs V001-V119) so a reader can see which variable it is.
Columns that are not in that table sit last, under "Additional series".
Palette follows the presentation Ronaldo asked the workbook to match: forest green, gold, ivory, sage, charcoal.
"""
import datetime as dt

from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from coverage_map import ROWS as _COV
from handoff_table import CHAT, SECTIONS
import roles

_STEP = {r[0]: r[10] for r in _COV}
_PURPOSE = {r[0]: r[11] for r in _COV}

FOREST, GOLD, IVORY, SAGE, CHAR, GREY = "173F35", "B49A58", "F7F6F2", "E3E9E2", "252B2A", "6B726F"
NAME = "Arial"
TYPE_FILL = {"native": "E3E9E2", "derived": "DCE6F1", "average": "ECECEC", "gap-fill": "FFF2CC", "reconstructed": "FBE5D6"}
TYPE_TEXT = {"native": "Native: as published by the source", "derived": "Derived: month-end or quarter-end value of a higher-frequency series, or a ratio of two series",
             "average": "Period average as published (no month-end value exists)", "gap-fill": "Gap fill from a second source, kept in its own column",
             "reconstructed": "Reconstructed from other published figures"}
STATUS_FILL = {"In workbook": "DDEBD8", "Partly in workbook": "FFF2CC", "Not retrieved": "F6D9D9", "Paid or unavailable": "E4E4E4",
               "Derived, not computed here": "DCE6F1", "Script written, not run": "FBE5D6"}
TAB = {"guide": FOREST, "data": "7FA08F", "event": GOLD, "dump": "A6A6A6"}
thin = Side(style="thin", color="D5D9D6")
BOTTOM = Border(bottom=thin)


def font(size=10, **kw):
    return Font(name=NAME, size=size, **kw)


def fill(hex_):
    return PatternFill("solid", fgColor=hex_)


# ---------------------------------------------------------------- mapping of workbook columns to handoff IDs
_PREFIX = [
    ("Uranium spot, month-end", "V001"), ("Uranium spot, quarter-end", "V001"), ("Uranium long-term, month-end", "V002"), ("Uranium long-term, quarter-end", "V002"),
    ("Conversion spot", "V007"), ("Conversion long-term", "V008"),
    ("SPUT price", "V012"), ("SPUT NAV per unit", "V018"), ("SPUT premium/discount", "V020"), ("SPUT total NAV", "V019"), ("SPUT units outstanding", "V021"),
    ("SPUT units issued", "V022"), ("SPUT gross issuance", "V024"), ("SPUT cash", "V026"), ("SPUT pounds bought", "V027"), ("SPUT holdings", "V029"),
    ("SPUT TSX high", "V013"), ("SPUT TSX low", "V013"), ("SPUT TSX average daily volume", "V014"), ("SPUT TSX volume", "V014"),
    ("SPUT purchases as %", "V032"), ("SPUT purchases in quarter", "V027"),
    ("U-UN.TO", "V011"), ("UPC adj. close", "D44"), ("U-U.TO", "V012"), ("YCA.L", "V033"),
    ("Yellow Cake NAV", "V042"), ("Yellow Cake uranium held", "V040"), ("Yellow Cake lb basis", "V040"), ("Yellow Cake quarter-end spot", "V042"),
    ("Yellow Cake announcement", "V042"), ("Yellow Cake premium", "V043"),
    ("UxC spot volume", "V101"),
    ("Real GDP, US", "V083"), ("Real GDP growth, US", "V083"), ("Industrial production, OECD", "V082"), ("Industrial production, G7", "V082"), ("OECD industrial production growth", "V082"),
    ("Categorical EPU", "V088"), ("Geopolitical risk index, overall", "V089"), ("Geopolitical risk index, daily overall", "V089"), ("Geopolitical risk index, threats", "V090"),
    ("Geopolitical risk index, acts", "V090"), ("Geopolitical risk index, daily threats", "V090"), ("Geopolitical risk index, daily acts", "V090"),
    ("Kilian global real economic activity index", "V084"), ("TED spread", "V076"), ("Commercial paper, 90-day", "V076"), ("T-bill, 3-month", "V076"), ("SOFR", "V076"),
    ("US uranium concentrate production", "V092"), ("US reactor owners inventory", "V095"), ("US supplier inventory", "V096"),
    ("US reactor requirements, maximum anticipated", "V098"), ("US unfilled requirements", "V099"), ("US contracted deliveries", "V100"),
    ("Coal, Australia (IMF)", "V045"), ("Copper (IMF)", "V054"), ("Natural gas, Henry Hub (IMF)", "V051"),
    ("WTI crude, spot, month-end", "V049"), ("Brent crude, spot, month-end", "V048"), ("WTI crude, spot", "V056"), ("Brent crude, spot", "V056"),
    ("Gold futures", "V059"), ("Copper futures", "V059"), ("WTI futures", "V059"), ("Brent futures", "V059"), ("Natural gas futures", "V059"),
    ("Henry Hub natural gas, spot", "V057"),
    ("Fed funds rate, monthly", "V062"), ("Fed funds rate", "V061"), ("3-month Treasury", "V063"), ("US 10y yield", "V064"), ("US 10y real yield", "V065"),
    ("Term spread 10y-3m", "V066"), ("Breakeven inflation", "V067"), ("Consumer price index", "V068"), ("Baa corporate yield", "V070"), ("Aaa corporate yield", "V071"),
    ("Corporate option-adjusted spread", "V073"),
    ("Broad dollar index, real", "V078"), ("Broad dollar index", "V077"), ("USD/CAD", "V079"), ("GBP/USD", "V080"), ("US industrial production", "V081"),
    ("VIX", "V085"), ("Economic policy uncertainty, US, daily", "V087"), ("Economic policy uncertainty", "V086"),
    ("UX futures volume", "V118"), ("UX futures open interest", "V118"),
]
_EQUITY_V119 = ("URA", "URNM", "NLR", "CCJ", "KAP.IL", "UEC", "DNN", "NXE", "PDN.AX", "BOE.AX", "DYL.AX", "SLX.AX", "BWXT", "1164.HK")
_PINK = [("Crude oil, average", "V050"), ("Crude oil, Brent", "V048"), ("Crude oil, WTI", "V049"), ("Coal, Australian", "V045"), ("Coal, South African", "V046"),
         ("Natural gas, US", "V051"), ("Natural gas, Europe", "V052"), ("Liquefied natural gas, Japan", "V053"), ("Gold", "V047"), ("Copper", "V054"),
         ("Total Index", "V055"), ("Energy (WB index)", "V055"), ("Non-energy (WB index)", "V055"), ("Metals", "V055"), ("Precious Metals (WB index)", "V055"), ("Base Metals", "V055")]


def vid(header, pink=False):
    h = header.strip()
    if pink:
        for p, i in _PINK:
            if h.startswith(p):
                return i
        return None
    for p, i in _PREFIX:
        if h.startswith(p):
            return i
    if h.split(" ")[0] in _EQUITY_V119:
        return "V119"
    return None


def sec_no(id_):
    if not id_:
        return 9
    for n, _t, _note, a, b in SECTIONS:
        if a <= id_ <= b:
            return n
    return 9


def sec_label(id_, extra="Additional series (not in the handoff table)"):
    n = sec_no(id_)
    if n == 9:
        return extra
    return f"{n}. {SECTIONS[n - 1][1]}"


def short(header):
    return header.split(" [")[0].strip()


EXTRA_ROLE = "Additional series (not in the handoff table)"


def role_label(id_):
    """Short role name of a handoff record (A. Uranium price, B2. SPUT premium/discount, ...)."""
    st = _STEP.get(id_)
    return roles.ROLES[st][0] if st in roles.ROLES else EXTRA_ROLE


def role_why(id_):
    st = _STEP.get(id_)
    return roles.ROLES[st][1] if st in roles.ROLES else ""


def why_col(id_, header):
    """One line saying why this column matters."""
    if id_:
        return roles.D_PURPOSE.get(id_) or _PURPOSE.get(id_) or role_why(id_)
    for key, txt in roles.EXTRA_WHY:
        if key.lower() in header.lower():
            return txt
    return roles.EXTRA_DEFAULT


def _fmt_key(k):
    if isinstance(k, int):
        return str(k)
    return f"{k[0]}-{k[1]:02d}" if isinstance(k, tuple) else k.isoformat()


def _title(ws, title, subtitle, link=True):
    ws["A1"] = title
    ws["A1"].font = font(15, bold=True, color=FOREST)
    ws["A2"] = subtitle
    ws["A2"].font = font(9, italic=True, color=GREY)
    if link:
        ws["A3"] = "← Contents"
        ws["A3"].hyperlink = "#'Contents'!A1"
        ws["A3"].font = font(9, underline="single", color=FOREST)
    why = roles.SHEET_WHY.get(ws.title)
    if why:
        ws["B3"] = "Why it matters: " + why
        ws["B3"].font = font(10, bold=True, color=FOREST)
    ws.sheet_view.showGridLines = False


# ---------------------------------------------------------------- data sheets (one column per series)
def write_data_sheet(wb, name, title, subtitle, keys, cols, keyhdr, keyfmt, keyval, width=17, tab="data", pink=False, extra_label="Additional series (not in the handoff table)"):
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = TAB[tab]
    cols = [c for _, c in sorted(enumerate(cols), key=lambda t: (sec_no(vid(t[1]["header"], pink)), vid(t[1]["header"], pink) or "~", t[0]))]
    _title(ws, title, subtitle)
    labels = {4: "Role in the thesis", 5: "Variable", 6: "Why it matters", 7: "Series", 8: "Source", 9: "Type · coverage", 10: keyhdr}
    for r, lab in labels.items():
        c = ws.cell(r, 1, lab)
        c.font = font(9 if r < 10 else 10, bold=True, color=CHAR if r == 10 else GREY)
        c.alignment = Alignment(vertical="center", wrap_text=True)
        c.fill = fill(IVORY if r != 10 else "FFFFFF")
    start, prev = 2, None
    keyset = set(keys)
    for j, c in enumerate(cols, 2):
        id_ = vid(c["header"], pink)
        sl = role_label(id_) if id_ else extra_label
        if sl != prev:
            if prev is not None:
                _band(ws, start, j - 1, prev)
            start, prev = j, sl
        if id_ in CHAT:
            var = f"{id_} · {CHAT[id_][0]} — {CHAT[id_][1]}"
        elif id_:  # extra items of the variables doc (D-ids) are described in the coverage map
            r_ = next(r for r in _COV if r[0] == id_)
            var = f"{id_} · {r_[2]} — {r_[3]}"
        else:
            var = "Not in the handoff table"
        vals = {k: v for k, v in c["values"].items() if k in keyset}
        cov = f"{TYPE_TEXT.get(c['kind'], c['kind']).split(':')[0]} · {_fmt_key(min(vals))} to {_fmt_key(max(vals))} · {len(vals)} obs" if vals else "no observations"
        row = {5: var, 6: why_col(id_, c["header"]), 7: short(c["header"]), 8: c["source"], 9: cov, 10: c["unit"]}
        for r, v in row.items():
            cell = ws.cell(r, j, v)
            cell.alignment = Alignment(wrap_text=True, vertical="top" if r < 10 else "center", horizontal="center" if r == 10 else "left")
            cell.font = font(10 if r in (5, 7) else 9, bold=(r in (5, 7)), italic=(r == 6), color=CHAR if r in (5, 7, 10) else (FOREST if r == 6 else GREY))
            cell.fill = fill(IVORY)
        ws.cell(9, j).fill = fill(TYPE_FILL.get(c["kind"], SAGE))
        ws.cell(10, j).fill = fill("FFFFFF")
        ws.cell(10, j).border = Border(bottom=Side(style="medium", color=FOREST))
        note = " ".join(x for x in [c.get("note", ""), f"File: {c['raw']}" if c.get("raw") and c["raw"] != "see note" else ""] if x)
        if note:
            ws.cell(7, j).comment = Comment(note, "Dataset", width=380, height=140)
        ws.column_dimensions[get_column_letter(j)].width = width
    if prev is not None:
        _band(ws, start, len(cols) + 1, prev)
    ws.cell(10, 1).border = Border(bottom=Side(style="medium", color=FOREST))
    for key in keys:
        ws.append([keyval(key)] + [c["values"].get(key) for c in cols])
    from copy import copy
    f10 = font()
    tmpl = ws.cell(11, 1)
    tmpl.number_format, tmpl.font = keyfmt, font(bold=True, color=CHAR)
    styles = [copy(tmpl._style)]
    for j, c in enumerate(cols, 2):
        t = ws.cell(11, j)
        t.font, t.number_format = f10, c["fmt"]
        styles.append(copy(t._style))
    for r in range(11, ws.max_row + 1):
        for j, st in enumerate(styles, 1):
            ws.cell(r, j)._style = copy(st)
    ws.column_dimensions["A"].width = 16
    for r, h in {1: 24, 4: 20, 5: 62, 6: 78, 7: 48, 8: 30, 9: 28, 10: 20}.items():
        ws.row_dimensions[r].height = h
    ws.freeze_panes = "B11"
    return ws


def _band(ws, c1, c2, text):
    ws.cell(4, c1, text)
    if c2 > c1:
        ws.merge_cells(start_row=4, start_column=c1, end_row=4, end_column=c2)
    for j in range(c1, c2 + 1):
        ws.cell(4, j).fill = fill(FOREST if "Additional" not in text and "not in the handoff" not in text else GOLD)
    cell = ws.cell(4, c1)
    cell.font = font(10, bold=True, color="FFFFFF")
    cell.alignment = Alignment(vertical="center", horizontal="left", indent=1, wrap_text=False)


# ---------------------------------------------------------------- record tables (one row per observation or document)
def write_table_sheet(wb, name, title, subtitle, headers, rows, widths, fmts=None, wrap=(), tab="event", first_row=5, plain=False):
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = TAB[tab]
    _title(ws, title, subtitle)
    for j, h in enumerate(headers, 1):
        c = ws.cell(first_row, j, h)
        c.font, c.fill = font(10, bold=True, color="FFFFFF"), fill(FOREST)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
    for i, row in enumerate(rows, first_row + 1):
        for j, v in enumerate(row, 1):
            c = ws.cell(i, j, v)
            if plain:
                continue
            c.font = font()
            if fmts and fmts[j - 1]:
                c.number_format = fmts[j - 1]
            c.border = BOTTOM
            c.alignment = Alignment(wrap_text=(j in wrap), vertical="top")
            if i % 2 == 0:
                c.fill = fill(IVORY)
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
        if plain:
            cd = ws.column_dimensions[get_column_letter(j)]
            cd.font = font()
            if fmts and fmts[j - 1]:
                cd.number_format = fmts[j - 1]
    ws.row_dimensions[first_row].height = 44
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = ws.cell(first_row + 1, 1)
    ws.auto_filter.ref = f"A{first_row}:{get_column_letter(len(headers))}{max(ws.max_row, first_row + 1)}"
    return ws


def write_variables_sheet(wb, rows_by_id, extras):
    """rows_by_id: id -> (workbook status, where, note). extras: list of (id, variable, variant, status, where, note)."""
    ws = wb.create_sheet("Variables")
    ws.sheet_properties.tabColor = TAB["guide"]
    _title(ws, "Variables", "One row per variable of the handoff table, in its order, with its role in the thesis and what this workbook holds. The grey block on the right keeps the handoff's own columns for reference.")
    heads = ["ID", "Variable", "Variant", "Role in the thesis", "Why it matters", "In this workbook", "Where in the workbook", "Note",
             "Source (handoff)", "Status (handoff)", "Remaining issue (handoff)", "Where / how to retrieve (handoff)", "Python (handoff)"]
    widths = [8, 22, 32, 22, 56, 18, 48, 56, 28, 12, 40, 32, 22]
    r = 5
    for j, h in enumerate(heads, 1):
        c = ws.cell(r, j, h)
        c.font, c.fill = font(10, bold=True, color="FFFFFF"), fill(FOREST if j <= 5 else (GOLD if j <= 8 else GREY))
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 30
    r += 1

    def put(r, vals):
        for j, v in vals.items():
            c = ws.cell(r, j, v)
            c.font = font(9 if j > 3 else 10, bold=(j == 2), color=(FOREST if j == 5 else (GREY if j >= 9 else CHAR)), italic=(j == 5))
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = BOTTOM
        wst = vals[6]
        ws.cell(r, 6).fill = fill(STATUS_FILL.get(wst, "FFFFFF"))
        ws.cell(r, 6).font = font(9, bold=True)

    for n, title, note, a, b in SECTIONS:
        ws.cell(r, 1, f"{n}. {title}").font = font(11, bold=True, color=FOREST)
        for j in range(1, len(heads) + 1):
            ws.cell(r, j).fill = fill(SAGE)
        r += 1
        for id_ in sorted(i for i in CHAT if a <= i <= b):
            var, variant, st, src, issue, where, py = CHAT[id_]
            wst, wwhere, wnote = rows_by_id[id_]
            put(r, {1: id_, 2: var, 3: variant, 4: role_label(id_), 5: why_col(id_, ""), 6: wst, 7: wwhere, 8: wnote, 9: src, 10: st, 11: issue, 12: where, 13: py})
            r += 1
    ws.cell(r, 1, "Additional items from the variables document (not in the handoff table)").font = font(11, bold=True, color=FOREST)
    for j in range(1, len(heads) + 1):
        ws.cell(r, j).fill = fill(SAGE)
    r += 1
    for id_, var, variant, wst, wwhere, wnote in extras:
        put(r, {1: id_, 2: var, 3: variant, 4: roles.ROLES["H. Extra items from the variables doc"][0], 5: why_col(id_, ""), 6: wst, 7: wwhere, 8: wnote})
        r += 1
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    for j in range(9, 14):  # handoff columns: reference only, grouped and collapsed
        ws.column_dimensions[get_column_letter(j)].outlineLevel = 1
        ws.column_dimensions[get_column_letter(j)].hidden = True
    ws.freeze_panes = "D6"
    ws.row_dimensions[1].height = 24
    ws.auto_filter.ref = f"A5:{get_column_letter(len(heads))}{r - 1}"
    return ws


def write_contents(wb, sheets, status_counts, built, by=None, extras=None):
    """sheets: list of (group, name, description, frequency, covers, rows, columns)."""
    ws = wb.create_sheet("Contents")
    ws.sheet_properties.tabColor = FOREST
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Uranium financialization: data set"
    ws["A1"].font = font(18, bold=True, color=FOREST)
    ws["A2"] = f"Master's thesis, ISCAP. Built {built}. Values are as published by each source; nothing is interpolated, filled forward or averaged to match another frequency."
    ws["A2"].font = font(9, italic=True, color=GREY)
    ws["A3"] = ("Purpose: the data needed to study how uranium prices relate to the financial vehicles that hold physical uranium (chiefly Sprott Physical Uranium Trust), "
                "compared with other commodities and with macro-financial conditions. This workbook holds data only: no analysis and no hypothesis.")
    ws["A3"].font = font(10, bold=True, color=CHAR)
    ws.merge_cells("A3:F3")
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[3].height = 30
    r = 5
    if by:
        c = ws.cell(r, 1, "What the data is for: the roles")
        c.font = font(11, bold=True, color="FFFFFF")
        for j in range(1, 7):
            ws.cell(r, j).fill = fill(FOREST)
        r += 1
        for j, h in enumerate(["Role", "Why it matters", "Variables", "In this workbook"], 1):
            c = ws.cell(r, j, h)
            c.font, c.fill = font(9, bold=True, color=CHAR), fill(SAGE)
        r += 1
        stat = {}
        for id_, (wst, _w, _n) in by.items():
            stat.setdefault(_STEP.get(id_), []).append(wst)
        for _id, _v, _vr, wst, _w, _n in (extras or []):
            stat.setdefault(_STEP.get(_id), []).append(wst)
        for step, (short_name, why) in roles.ROLES.items():
            ids = sorted(i for i, st in _STEP.items() if st == step)
            sts = stat.get(step, [])
            n_in, n_part = sts.count("In workbook"), sts.count("Partly in workbook")
            n_drv = sum(1 for x in sts if x.startswith("Derived"))
            n_out = len(sts) - n_in - n_part - n_drv
            txt = f"{n_in} in · {n_part} partly · {n_out} missing" + (f" · {n_drv} derived later" if n_drv else "")
            for j, v in enumerate([short_name, why, f"{ids[0]} to {ids[-1]}" if ids else "", txt], 1):
                cc = ws.cell(r, j, v)
                cc.font = font(10 if j == 1 else 9, bold=(j == 1), color=FOREST if j == 1 else CHAR)
                cc.alignment = Alignment(wrap_text=True, vertical="top")
                cc.border = BOTTOM
            ws.row_dimensions[r].height = 30
            r += 1
        r += 1
    heads = ["Sheet", "What it holds", "Frequency", "Covers", "Rows", "Columns"]
    group = None
    for g, name, desc, freq, covers, nrow, ncol in sheets:
        if g != group:
            group = g
            c = ws.cell(r, 1, g)
            c.font = font(11, bold=True, color="FFFFFF")
            for j in range(1, 7):
                ws.cell(r, j).fill = fill(FOREST)
            r += 1
            for j, h in enumerate(heads, 1):
                c = ws.cell(r, j, h)
                c.font, c.fill = font(9, bold=True, color=CHAR), fill(SAGE)
            r += 1
        c = ws.cell(r, 1, name)
        c.hyperlink = f"#'{name}'!A1"
        c.font = font(10, bold=True, underline="single", color=FOREST)
        for j, v in enumerate([desc, freq, covers, nrow, ncol], 2):
            cc = ws.cell(r, j, v)
            cc.font = font(9)
            cc.alignment = Alignment(wrap_text=True, vertical="top")
        for j in range(1, 7):
            ws.cell(r, j).border = BOTTOM
            ws.cell(r, j).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 5).number_format = ws.cell(r, 6).number_format = "#,##0"
        r += 1
        ws.row_dimensions[r - 1].height = 30
    r += 1
    c = ws.cell(r, 1, "How to read the data sheets")
    c.font = font(11, bold=True, color=FOREST)
    r += 1
    rules = [
        "Each data sheet has one column per series and one row per period. Blank cell = no observation. A series is in the sheet of its own publication frequency.",
        "The header block names, for every column: its role in the thesis, the variable and variant from the handoff table (ID V001 to V119), why it matters, the exact series, the source, its type and coverage, and the unit.",
        "Month-end = the last available observation in the month. Period averages appear only where no daily source exists, and are labelled as averages.",
        "Columns that do not correspond to a record of the handoff table sit last, under Additional series.",
        "Hover over a series name for its note and raw file. The raw files are in data/raw and are never edited.",
    ]
    for t in rules:
        c = ws.cell(r, 1, "•  " + t)
        c.font = font(9)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 26
        r += 1
    r += 1
    ws.cell(r, 1, "Type colours (row “Type · coverage”)").font = font(11, bold=True, color=FOREST)
    r += 1
    for k, t in TYPE_TEXT.items():
        ws.cell(r, 1, k.capitalize()).fill = fill(TYPE_FILL[k])
        ws.cell(r, 1).font = font(9, bold=True)
        c = ws.cell(r, 2, t)
        c.font = font(9)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        r += 1
    r += 1
    ws.cell(r, 1, "Status of the 119 handoff variables").font = font(11, bold=True, color=FOREST)
    r += 1
    for k, n in status_counts:
        ws.cell(r, 1, k).fill = fill(STATUS_FILL.get(k, "FFFFFF"))
        ws.cell(r, 1).font = font(9, bold=True)
        ws.cell(r, 2, n).font = font(9)
        ws.cell(r, 2).alignment = Alignment(horizontal="left")
        r += 1
    for j, w in enumerate([26, 78, 14, 40, 10, 10], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[1].height = 30
    return ws
