"""Derived monthly columns for thesis_dataset.xlsx (V006, V031, V069, V072, V074, V110-V115, D13, D27, D35, D45).

Every value is computed from series already in the workbook; nothing is filled or interpolated.
A value is blank when any of its inputs is missing.
"""
import math
import statistics as st


def _prev(k, n=1):
    y, m = k
    t = y * 12 + (m - 1) - n
    return (t // 12, t % 12 + 1)


def _find(M, sid):
    for c in M:
        if c["sid"] == sid:
            return c["values"]
    raise KeyError(sid)


def _logret(v):
    out = {}
    for k, x in v.items():
        p = v.get(_prev(k))
        if p and x and p > 0 and x > 0:
            out[k] = math.log(x / p)
    return out


def _roll(vals, k, n, minobs):
    xs = [vals.get(_prev(k, i)) for i in range(n)]
    return xs


def _rolling_std(r, n=12):
    out = {}
    for k in r:
        w = [r.get(_prev(k, i)) for i in range(n)]
        if all(x is not None for x in w):
            out[k] = st.stdev(w) * math.sqrt(12) * 100
    return out


def _rolling_corr(a, b, n=36):
    out = {}
    for k in a:
        wa = [a.get(_prev(k, i)) for i in range(n)]
        wb = [b.get(_prev(k, i)) for i in range(n)]
        if all(x is not None for x in wa) and all(x is not None for x in wb):
            if st.pstdev(wa) > 0 and st.pstdev(wb) > 0:
                out[k] = st.correlation(wa, wb)
    return out


def derived_columns(M, col, NUM4, INT):
    spot = _find(M, "spot_usd_lb")
    lt = _find(M, "longterm_usd_lb")
    cpi = _find(M, "CPIAUCSL")
    baa, aaa = _find(M, "BAA"), _find(M, "AAA")
    out = []
    src = "computed in scripts/derived.py from columns of this workbook"

    def add(header, vals, sid, unit, note, role):
        out.append(col(header, vals, "derived", src, sid, "monthly", "computed", unit, note, role, "", NUM4))

    add("Spot-long-term price difference, month-end [USD/lb]", {k: lt[k] - spot[k] for k in spot if k in lt and lt[k] is not None and spot[k] is not None},
        "d_spread_lt_spot", "USD per lb", "Cameco long-term minus Cameco spot, same month-end", "V006")

    r = _logret(spot)
    add("Uranium spot log return, monthly [percent]", {k: v * 100 for k, v in r.items()}, "d_ret_spot", "percent", "ln(spot t / spot t-1) x 100, Cameco month-end spot", "V110")
    mom = {}
    for k in spot:
        p = spot.get(_prev(k, 12))
        if p and spot[k] and p > 0:
            mom[k] = math.log(spot[k] / p) * 100
    add("Uranium spot momentum, 12-month log return [percent]", mom, "d_mom12", "percent", "ln(spot t / spot t-12) x 100; the 12-month window is a choice, change it in the analysis if needed", "V112")
    add("Uranium spot volatility, 12-month rolling, annualised [percent]", _rolling_std(r), "d_vol12", "percent",
        "standard deviation of the last 12 monthly log returns x sqrt(12); needs 12 consecutive months", "V113")

    infl_m = {k: (cpi[k] / cpi[_prev(k)] - 1) * 100 for k in cpi if cpi.get(_prev(k))}
    infl_y = {k: (cpi[k] / cpi[_prev(k, 12)] - 1) * 100 for k in cpi if cpi.get(_prev(k, 12))}
    add("US inflation, month on month [percent]", infl_m, "d_infl_m", "percent", "CPI-U seasonally adjusted, percent change on the previous month", "V069")
    add("US inflation, year on year [percent]", infl_y, "d_infl_y", "percent", "CPI-U seasonally adjusted, percent change on the same month a year earlier", "V069")
    base_k = max(cpi)
    add(f"Uranium spot in real terms, {base_k[0]}-{base_k[1]:02d} dollars [USD/lb]", {k: spot[k] * cpi[base_k] / cpi[k] for k in spot if k in cpi},
        "d_spot_real", "USD per lb", "spot x CPI(latest month) / CPI(month); deflator: US CPI-U seasonally adjusted", "V111")

    add("Credit spread, Baa minus Aaa [percentage points]", {k: baa[k] - aaa[k] for k in baa if k in aaa}, "d_baa_aaa", "percentage points", "Moody's Baa yield minus Aaa yield, monthly", "V072")

    for sid, name in [("DGS10", "US 10y yield"), ("DGS3MO", "US 3m Treasury yield"), ("DFF", "Fed funds rate")]:
        s = _find(M, sid)
        add(f"{name}, monthly change of month-end [percentage points]", {k: s[k] - s[_prev(k)] for k in s if s.get(_prev(k)) is not None}, f"d_chg_{sid}",
            "percentage points", "month-end level minus previous month-end level", "V074")

    # futures basis
    close = _find(M, "close") if False else next(c["values"] for c in M if c["header"].startswith("UX futures close (Stooq)"))
    add("Futures basis, Stooq UX close minus Cameco spot [USD/lb]", {k: close[k] - spot[k] for k in close if k in spot}, "d_basis", "USD per lb",
        "UX close (Stooq monthly bar, front contract, roll undocumented) minus Cameco month-end spot; months Stooq lacks are blank", "D35")

    # cross-market co-movement (rolling 36 months)
    peers = [("GC=F", "gold"), ("SI=F", "silver"), ("HG=F", "copper"), ("DCOILWTICO", "WTI oil")]
    for sid, name in peers:
        pr = _logret(_find(M, sid))
        add(f"Rolling 36-month correlation of returns, uranium vs {name} [correlation]", _rolling_corr(r, pr), f"d_corr_{sid}", "correlation",
            "monthly log returns, 36 consecutive months, Cameco spot vs month-end price of the peer", "V114")
        ar = {k: abs(v) for k, v in r.items()}
        ap = {k: abs(v) for k, v in pr.items()}
        add(f"Rolling 36-month correlation of absolute returns, uranium vs {name} [correlation]", _rolling_corr(ar, ap), f"d_vcorr_{sid}", "correlation",
            "volatility co-movement proxy: correlation of absolute monthly log returns over 36 months", "V115")

    # fund volume scaling
    vol = _find(M, "U-U.TO + U-UN.TO volume")
    uo = _find(M, "units_outstanding")  # may pick reported; prefer reconstructed
    rec = next((c["values"] for c in M if c["header"].startswith("SPUT units outstanding, month-end, reconstructed")), None)
    if rec:
        add("SPUT TSX volume / units outstanding, month [ratio]", {k: vol[k] / rec[k] for k in vol if k in rec and rec[k]}, "d_turnover", "ratio",
            "month total TSX volume divided by reconstructed month-end units outstanding (Jun 2022 to Dec 2025 only)", "V031")

    # SPUT discount flag
    pdv = _find(M, "premium_discount_pct")
    add("SPUT trades below NAV, flag [1 = discount]", {k: (1 if v < 0 else 0) for k, v in pdv.items()}, "d_discount_flag", "flag",
        "1 when the month-end premium/discount is below zero", "D45")

    months = sorted(spot)
    first, last = months[0], months[-1]
    allm = []
    y, m = first
    while (y, m) <= last:
        allm.append((y, m))
        m += 1
        if m == 13:
            y, m = y + 1, 1
    for hdr, sid, start, note in [("Regime: post-2005, flag", "d_reg_2005", (2005, 1), "1 from Jan 2005"),
                                   ("Regime: SPUT listed, flag", "d_reg_sput", (2021, 7), "1 from Jul 2021 (SPUT began trading 26 Jul 2021)"),
                                   ("Regime: Russia-Ukraine war, flag", "d_reg_war", (2022, 2), "1 from Feb 2022 (invasion 24 Feb 2022)")]:
        add(hdr, {k: (1 if k >= start else 0) for k in allm}, sid, "flag", note, "D13")
    add("Event month: Kazakhstan unrest, flag", {k: (1 if k == (2022, 1) else 0) for k in allm}, "d_ev_kaz", "flag", "1 in Jan 2022 (unrest from 2 Jan 2022)", "D27")
    add("Event month: Russian invasion of Ukraine, flag", {k: (1 if k == (2022, 2) else 0) for k in allm}, "d_ev_inv", "flag", "1 in Feb 2022 (24 Feb 2022)", "D27")
    return out
