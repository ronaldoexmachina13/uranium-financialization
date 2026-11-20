"""NOT USED. Kept only as reference (D-029 practice: removed code is moved here, never deleted).

On 20 Nov 2026 Ronaldo decided that the six ratio columns built by Claude leave the workbook from v6 on:
the five SPUT statement ratios (net proceeds, average issue price, pounds per unit, uranium oxide fair value per lb,
cash as % of NAV) and Quarterly I (SPUT purchases as % of UxC spot volume). Their source figures stay in the workbook.
The thesis calculations are Ronaldo's own (D-029). The lines below are copied unchanged from scripts/build_dataset.py
as it built thesis_dataset_v5.xlsx (19 Nov 2026). This file is not imported by any script.
Section 3 (added 20 Nov 2026, answers): the Yellow Cake premium/discount column, removed from v6 the same day (Ronaldo's answer 2), copied from the v6 build.
"""

# ---- 1. scripts/build_dataset.py, uxc_volume(): Quarterly column I (v5)
#     ev = S["events"]
#     last_ev = max(e[0] for e in ev)
#     sq, share = {}, {}
#     for qd, v in qt.items():
#         if qd < dt.date(2021, 9, 30) or qd > last_ev:
#             continue
#         a = sum(e[1] or 0 for e in ev if dt.date(qd.year, qd.month - 2, 1) <= e[0] <= qd)
#         sq[qd] = a / 1e6
#         share[qd] = a / (v * 1e6) * 100
#     Q.append(col("SPUT purchases in quarter [million lb]", sq, ...))          # kept in v6
#     Q.append(col("SPUT purchases as % of UxC spot volume, quarter", share, "derived", "Sprott fund page and UxC via Yellow Cake", "SPUT lb / UxC volume", "quarterly", "quarter", "percent",
#                  "numerator from Sprott's purchase list (trade dates), denominator the first-reported UxC quarter volume; UxC volume is a provider-defined total and may include SPUT's own purchases", "H3", raw, NUM))

# ---- 2. scripts/build_dataset.py, main(): the five ratio columns of the SPUT statements sheet (v5 columns P to T)
#     heads = [..., "Source", "Note", "Net proceeds, proceeds minus costs [USD thousand]", "Average issue price [USD]", "Pounds per unit [lb]",
#              "Uranium oxide fair value per lb [USD]", "Cash as % of NAV"]
#     rows.append(v + [r["source"], note_, pr_ - co_ if pr_ is not None and co_ is not None else None, pr_ * 1000 / iss_ if pr_ is not None and iss_ else None,
#                      lb_ / units_ if lb_ and units_ else None, fv_ * 1000 / lb_ if fv_ and lb_ else None, cash_ / nav_ * 100 if cash_ is not None and nav_ else None] + cells_)
#     widths: [13, 14, 16, 16, 14, 16, 14, 18, 16, 16, 18, 18, 18, 36, 46, 18, 14, 14, 16, 12] + [16] * 7
#     fmts:   ["yyyy-mm-dd", None, INT, INT, "0.0000", NUM, INT, INT, INT, INT, INT, INT, INT, None, None, INT, NUM, "0.0000", NUM, NUM] + [...]
#     subtitle: "...; the five columns after Note are ratios of the others. ..."

# ---- 3. 20 Nov 2026 (answers): scripts/build_dataset.py, load(): Quarterly column P of v6 (Q in v5), "Yellow Cake premium/discount to NAV, quarter-end".
#     Removed on 20 Nov 2026 after Ronaldo's answer 2 (a ratio built in the workbook; D-029, D-058). Its inputs stay in the workbook: YCA.L close
#     (Daily, Monthly, Prices OHLCV; data/raw/yahoo/YCA.L.csv) and Yellow Cake NAV per share (Quarterly; data/raw/yellowcake/nav_quarterly.csv).
#     Copied unchanged from scripts/build_dataset.py as it built thesis_dataset_v6.xlsx at 12:42 on 20 Nov 2026.
#     yca = sorted((day(r["date"]), num(r["close"])) for r in equities["YCA.L"] if num(r["close"]))
#     prem = {}
#     for r in yc:
#         qd, nav = day(r["quarter_end"]), num(r["nav_per_share_gbp"])
#         px = [c for d_, c in yca if d_ <= qd]
#         if nav and px and (qd - max(d_ for d_, _ in yca if d_ <= qd)).days <= 7:
#             prem[qd] = round((px[-1] / 100.0 / nav - 1) * 100, 2)
#     Q.append(col("Yellow Cake premium/discount to NAV, quarter-end [percent]", prem, "derived", "YCA.L close (pence) divided by 100, against the announced NAV", "close / nav_per_share_gbp", "quarterly",
#                  "last close on or before the quarter end (within 7 days)", "percent", "price in GBP vs NAV in GBP; negative = discount", "H3", yraw + " and data/raw/yahoo/YCA.L.csv", NUM))
#     Register row V043 as it read in v6 (scripts/build_dataset.py, OVERRIDE):
#     "V043": (DERIVED, "Quarterly: Yellow Cake premium/discount to NAV, quarter-end (a ratio built in this workbook from the YCA.L close and the announced NAV, 33 quarter-ends)",
#              "quarter-end price vs quarter-end NAV. The column is built in this workbook although the variable is a derived one; whether it stays is for Ronaldo (D-029)."),

#     20 Nov 2026 (text fix): the two rules of scripts/workbook_layout.py that served only this column (no effect once it was removed), copied unchanged:
#     in _PREFIX (column -> handoff ID):   ("Yellow Cake premium", "V043"),
#     in the timing rules:                 (("Quarterly",), "Yellow Cake premium", "CONST", "Built here from the YCA.L price and the quarter-end NAV per share (column note)."),


def _not_used():
    raise SystemExit("reference only: this code is not used (removed from the workbook on 20 Nov 2026, Ronaldo's decision)")
