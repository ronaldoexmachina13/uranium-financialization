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
    ("UxC spot volume", "V101"), ("World nuclear electricity produced", "V103"), ("World nuclear operated capacity", "V104"), ("World nuclear energy availability factor", "V105"), ("World nuclear load factor", "V105"), ("World nuclear construction starts", "V106"), ("World nuclear grid connections", "V106"), ("World nuclear permanent shutdowns", "V106"), ("World nuclear reactors in operation", "V106"), ("World nuclear unit capability factor", "V105"), ("World nuclear forced loss rate", "V105"),
    ("Real GDP, US", "V083"), ("Real GDP growth, US", "V083"), ("Industrial production, OECD", "V082"), ("Industrial production, G7", "V082"), ("OECD industrial production growth", "V082"),
    ("Categorical EPU", "V088"), ("Geopolitical risk index, overall", "V089"), ("Geopolitical risk index, daily overall", "V089"), ("Geopolitical risk index, threats", "V090"),
    ("Geopolitical risk index, acts", "V090"), ("Geopolitical risk index, daily threats", "V090"), ("Geopolitical risk index, daily acts", "V090"),
    ("Kilian global real economic activity index", "V084"), ("Baltic Dry Index", "D18"), ("Funding spread, commercial paper", "V076"), ("TED spread", "V076"), ("Commercial paper, 90-day", "V076"), ("T-bill, 3-month", "V076"), ("SOFR", "V076"),
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
# 19 Nov 2026: tag four series that already sit in the register (Data Curator, 19 Nov 2026), so their columns stop saying "Not in the handoff table".
_PREFIX += [("Monetary-policy surprise", "V075"), ("Newcastle coal futures", "V060"), ("S&P 500", "D17"), ("Google Trends search interest", "D21")]
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


# ---------------------------------------------------------------- timing codes and release dates (labels only)
# 19 Nov 2026: one timing code per data column (approved by Ronaldo), so a model does not mix month-end levels with averages, flows or report-date values by mistake.
TIMING_CODES = [
    ("CLOSE", "Daily: end-of-day value for the row's date (exchange close, fund NAV or futures settlement). Close times differ by exchange, so one date is not one moment."),
    ("DAY", "Daily: value for the row's date as published (rates, fixings, spot assessments, daily indices); not an exchange close."),
    ("ME", "Month-end level: the last available value in the month of a daily or weekly source."),
    ("ME-ASSESSED", "Month-end of an assessed price (set by price reporters, not traded); the convention is in the cell note."),
    ("QE", "Quarter-end level, as at the last day of the quarter."),
    ("QE-ASSESSED", "Quarter-end of an assessed price; the convention is in the cell note."),
    ("WE", "Week-end level: value as at the Friday of the row."),
    ("YE", "Year-end level (stock at 31 December)."),
    ("AVG", "Average over the period, as published."),
    ("FLOW", "Sum or count over the period (day, month, quarter or year)."),
    ("EXT", "Highest or lowest value within the period."),
    ("BAR", "Field of a monthly open-high-low-close bar of a continuous futures contract, as published; only the close is an end-of-period value."),
    ("IDX", "Index or measure for the whole period, as published (e.g. CPI, industrial production, EPU, GPR); not a point in time."),
    ("REL", "Value as at a report date only (blank between reports); forward-looking figures as of a stated date carry this code too."),
    ("CONST", "Constructed by arithmetic in the build from published figures (difference, ratio, product, sum of products or reconstruction)."),
    ("EVENT", "Sum over dated events in the period; blank when there is no event."),
    ("INFO", "Descriptive field about another column (text or a date), not a measurement."),
    ("CHECK", "Could not be labelled with confidence; the cell note says why."),
]
TIMING_TEXT = dict(TIMING_CODES)
TIMING_ROW, RELEASE_ROW, KEY_ROW, DATA_ROW = 10, 11, 12, 13  # header rows 4 to 12; data from row 13 on every data sheet
TIMING_BY_ID = {}   # handoff ID -> [(sheet, code)], filled while the data sheets are written; read by the Variables sheet
RELEASE = {}        # series -> (header text, cell comment), from data/release_dates.csv

_CAM = ("Cameco month-end = average of the UxC and TradeTech month-end assessments, not UxC alone. Dec 2017 holds the month-end from Cameco's 6-K (cell comment).")
_TSX = "Last TSX trading day of the month; TSX close and Sprott NAV about 4 pm ET."
_FRED_ME = "Last daily value of the month (FRED daily series)."
_FUT_ME = "Last close of the month of the Yahoo front-month continuous contract; no back-adjustment, so a month-end can sit either side of a roll."
_EQ_ME = "Last day with trading volume in the month (zero-volume days blank); exchange close in local time, so the same date is not the same moment."
_UX = ("UX contracts before Oct 2024 settle on the UxC month-end spot published on the last Monday of the month, so prices are tied to the assessed spot by construction; "
       "the rule after that is open (D-050).")
_TEXTIDX = "Newspaper or text-count index for the month."
_FRED_D = "FRED daily business-day series."
_EQ_D = "Exchange close in local time (NYSE 4 pm ET; ASX, HKEX 4 pm local; LSE for KAP.IL, YCA.L); the same date is not the same moment."
_FUT_D = "Yahoo continuous front-month close, dated in the exchange time zone; rolls not adjusted."

# (sheets, header prefix, code, note). First match wins; headers are the Series row (no unit).
_TIMING_RULES = [
    # Monthly
    (("Monthly",), "Uranium spot, month-end", "ME-ASSESSED", _CAM),
    (("Monthly",), "Uranium long-term, month-end", "ME-ASSESSED", _CAM),
    (("Monthly",), "UX final settlement, reconstructed", "ME-ASSESSED", "UxC month-end spot as published on the last Monday of the month (not calendar month-end), used as the UX final settlement; reconstructed from later reports."),
    (("Monthly",), "U-UN.TO adj. close, month-end", "ME", _TSX), (("Monthly",), "SPUT price, month-end", "ME", _TSX), (("Monthly",), "U-U.TO adj. close, month-end", "ME", _TSX),
    (("Monthly",), "SPUT NAV per unit, month-end", "ME", _TSX), (("Monthly",), "SPUT premium/discount, month-end", "ME", _TSX),
    (("Monthly",), "SPUT TSX high", "EXT", "Monthly high on the TSX, from the prospectus trading tables (printed months later in a supplement)."),
    (("Monthly",), "SPUT TSX low", "EXT", "Monthly low on the TSX, from the prospectus trading tables (printed months later in a supplement)."),
    (("Monthly",), "SPUT TSX average daily volume", "AVG", "Average daily TSX volume over the month (prospectus trading tables)."),
    (("Monthly",), "SPUT TSX volume, month total", "FLOW", "Daily TSX volume of both listings summed over the month in the build."),
    (("Monthly",), "SPUT total NAV, reported", "REL", "Only at SPUT report dates (SPUT statements sheet); release dates in data/release_dates.csv."),
    (("Monthly",), "SPUT units outstanding, reported", "REL", "Only at SPUT report dates (SPUT statements sheet); release dates in data/release_dates.csv."),
    (("Monthly",), "SPUT cash, reported", "REL", "Only at SPUT report dates (SPUT statements sheet); release dates in data/release_dates.csv."),
    (("Monthly",), "SPUT units outstanding, month-end, reconstructed", "CONST", "Reported units plus dated issuances to month-end; the issuance dates were published later (data/release_dates.csv)."),
    (("Monthly",), "SPUT units issued in month", "FLOW", "Sum of the dated issuances in the month; the dates were published later in prospectus documents (data/release_dates.csv)."),
    (("Monthly",), "SPUT gross issuance value in month", "CONST", "Sum of price x units of the month's dated issuances, built here; dates published later (data/release_dates.csv)."),
    (("Monthly",), "SPUT pounds bought in month", "FLOW", "Pounds of the purchase days (trade dates) in the month."),
    (("Monthly",), "SPUT holdings, month-end", "ME", "Last weekly (Friday) holdings value of the month; trade-date basis; rounded to 0.1 million lb."),
    (("Monthly",), "YCA.L adj. close, month-end", "ME", "Last LSE trading day of the month; LSE close, hours before North American closes."),
    (("Monthly",), "Coal, Australia (IMF)", "AVG", "IMF monthly average (FRED copy); accepted as an average (D-034)."),
    (("Monthly",), "Natural gas, Henry Hub (IMF)", "AVG", "IMF monthly average (FRED copy)."),
    (("Monthly",), "Copper (IMF)", "AVG", "IMF monthly average (FRED copy)."),
    (("Monthly",), "Uranium (IMF)", "AVG", "IMF monthly average (FRED copy); cross-check only."),
    (("Monthly",), "Brent crude, spot, month-end", "ME", "Last EIA daily spot assessment of the month (FRED)."),
    (("Monthly",), "WTI crude, spot, month-end", "ME", "Last EIA daily spot assessment of the month (FRED)."),
    (("Monthly",), "Henry Hub natural gas, spot, month-end", "ME", "Last EIA daily spot assessment of the month (FRED)."),
    (("Monthly",), "Gold futures", "ME", _FUT_ME), (("Monthly",), "Copper futures", "ME", _FUT_ME), (("Monthly",), "WTI futures", "ME", _FUT_ME),
    (("Monthly",), "Brent futures", "ME", _FUT_ME), (("Monthly",), "Natural gas futures", "ME", _FUT_ME), (("Monthly",), "Silver futures", "ME", _FUT_ME),
    (("Monthly",), "Platinum futures", "ME", _FUT_ME),
    (("Monthly",), "Newcastle coal futures close", "BAR", "Close of the Investing.com NCFMc1 monthly bar; roll undocumented."),
    (("Monthly",), "Fed funds rate, month-end", "ME", _FRED_ME),
    (("Monthly",), "Fed funds rate, monthly", "AVG", "FRED FEDFUNDS: monthly average of daily rates (FRED definition, grade B, not checked in a kept file)."),
    (("Monthly",), "Consumer price index", "IDX", "CPI for the month (prices collected through the month), seasonally adjusted; FRED latest vintage, revised. Oct 2025 is empty at the source."),
    (("Monthly",), "Baa corporate yield", "AVG", "Moody's monthly series on FRED: average of daily yields (FRED definition, grade B, not checked in a kept file)."),
    (("Monthly",), "Aaa corporate yield", "AVG", "Moody's monthly series on FRED: average of daily yields (FRED definition, grade B, not checked in a kept file)."),
    (("Monthly",), "Monetary-policy surprise", "EVENT", "Sum over the FOMC announcements in the month; months without one are blank."),
    (("Monthly",), "Funding spread", "CONST", "Month-end of the daily spread built here (commercial paper minus T-bill); the two rates' last values can fall on different days."),
    (("Monthly",), "Broad dollar index, real", "IDX", "Fed real broad dollar index, one value per month (how the month is formed not checked in a kept file)."),
    (("Monthly",), "US industrial production", "IDX", "Output over the month; revised (latest vintage)."),
    (("Monthly",), "Industrial production, OECD", "IDX", "Output over the month; revised (latest vintage); levels end 2024-03."),
    (("Monthly",), "Industrial production, G7", "IDX", "Output over the month; revised (latest vintage); levels end 2024-03."),
    (("Monthly",), "OECD industrial production growth", "IDX", "Year-on-year growth of output over the month; revised (latest vintage)."),
    (("Monthly",), "Kilian global real economic activity index", "IDX", "Monthly global real activity index (freight-rate based); the last 19 dates were re-set in the build (column note)."),
    (("Monthly",), "Economic policy uncertainty", "IDX", _TEXTIDX), (("Monthly",), "Categorical EPU", "IDX", _TEXTIDX),
    (("Monthly",), "Geopolitical risk index", "IDX", _TEXTIDX), (("Monthly",), "Energy uncertainty index", "IDX", _TEXTIDX),
    (("Monthly",), "UxC spot volume, month, first reported", "FLOW", "UxC spot volume for the month, first print; first quoted 20 to 242 days after the month (data/release_dates.csv)."),
    (("Monthly",), "UX futures volume (Stooq)", "FLOW", "Volume of the Stooq monthly bar; zero in all months, not used."),
    (("Monthly",), "UX futures volume, CME monthly report", "FLOW", "Contracts traded in the month (CME monthly report)."),
    (("Monthly",), "UX futures open interest", "CHECK", "Open interest as printed; the day of the month it refers to is not documented (Stooq: only 0 or 310; CME monthly report basis not verified)."),
    (("Monthly",), "UX futures open (Stooq)", "BAR", "Stooq monthly bar of the continuous UX front contract (roll undocumented). " + _UX),
    (("Monthly",), "UX futures high (Stooq)", "BAR", "Stooq monthly bar of the continuous UX front contract (roll undocumented). " + _UX),
    (("Monthly",), "UX futures low (Stooq)", "BAR", "Stooq monthly bar of the continuous UX front contract (roll undocumented). " + _UX),
    (("Monthly",), "UX futures close (Stooq)", "BAR", "Stooq monthly bar of the continuous UX front contract (roll undocumented). " + _UX),
    (("Monthly",), "UX futures close (Investing, gap fill)", "BAR", "Close of the Investing.com UXXc2 monthly bar (second contract, kept apart, D-035). " + _UX),
    (("Monthly",), "Baltic Dry Index, month-end", "ME", "Last daily value of the month, read from a chart source."),
    (("Monthly",), "Google Trends search interest", "IDX", "Relative search interest for the month, rescaled to 0-100 over the window fetched (the whole series changes at each fetch)."),
    (("Monthly",), "S&P 500 (FRED), month-end", "ME", _FRED_ME),
    (("Monthly",), "3-month Treasury yield", "ME", _FRED_ME), (("Monthly",), "US 10y", "ME", _FRED_ME), (("Monthly",), "Term spread 10y-3m", "ME", _FRED_ME),
    (("Monthly",), "Breakeven inflation", "ME", _FRED_ME), (("Monthly",), "Corporate option-adjusted spread", "ME", _FRED_ME), (("Monthly",), "TED spread", "ME", _FRED_ME),
    (("Monthly",), "Commercial paper", "ME", _FRED_ME), (("Monthly",), "T-bill, 3-month", "ME", _FRED_ME), (("Monthly",), "SOFR", "ME", _FRED_ME),
    (("Monthly",), "Broad dollar index, month-end", "ME", _FRED_ME), (("Monthly",), "USD/CAD", "ME", _FRED_ME), (("Monthly",), "GBP/USD", "ME", _FRED_ME),
    (("Monthly",), "VIX", "ME", _FRED_ME), (("Monthly",), "Baa-10y credit spread", "ME", _FRED_ME),
    (("Monthly",), "adj. close, month-end", "ME", _EQ_ME, "contains"), (("Monthly",), "S&P 500 (^GSPC), month-end", "ME", _EQ_ME),
    # Quarterly
    (("Quarterly",), "Uranium spot, quarter-end", "QE-ASSESSED", _CAM), (("Quarterly",), "Uranium long-term, quarter-end", "QE-ASSESSED", _CAM),
    (("Quarterly",), "Conversion", "QE-ASSESSED", "Quarter-end conversion price in Cameco's quarterly reports (average of TradeTech and UxC)."),
    (("Quarterly",), "SPUT purchases in quarter", "FLOW", "Pounds of the purchase days (trade dates) in the quarter."),
    (("Quarterly",), "SPUT purchases as %", "CONST", "Ratio built here: SPUT purchases over the first-reported UxC quarter volume, which is published weeks later."),
    (("Quarterly",), "Yellow Cake uranium held", "QE", "Holdings at quarter-end as stated (see lb basis); announced weeks later (announcement date column; data/release_dates.csv)."),
    (("Quarterly",), "Yellow Cake lb basis", "INFO", "Says how the pounds in the held columns were obtained."),
    (("Quarterly",), "Yellow Cake NAV per share", "QE", "Estimated NAV at quarter-end; announced 3 to 39 days later (announcement date column; data/release_dates.csv)."),
    (("Quarterly",), "Yellow Cake quarter-end spot used", "QE-ASSESSED", "Spot price Yellow Cake used for the quarter-end NAV, as given or derived for the announcement."),
    (("Quarterly",), "Yellow Cake announcement date", "INFO", "Date the quarter's Yellow Cake figures were announced: a release date."),
    (("Quarterly",), "Yellow Cake premium", "CONST", "Built here from the YCA.L price and the quarter-end NAV per share (column note)."),
    (("Quarterly",), "Real GDP growth", "IDX", "Annualised growth over the quarter, as published; revised (latest vintage)."),
    (("Quarterly",), "Real GDP, US", "FLOW", "Output over the quarter at a seasonally adjusted annual rate; revised (latest vintage)."),
    (("Quarterly",), "UxC spot volume, quarter, first reported", "FLOW", "UxC spot volume for the quarter, first print (data/release_dates.csv)."),
    (("Quarterly",), "World spot market volume", "FLOW", "Spot volume in the quarter as reported in Cameco's MD&A (filed weeks later)."),
    # Weekly
    (("Weekly",), "SPUT holdings", "WE", "Holdings as at the Friday, from the Sprott fund page; trade-date basis."),
    # Daily
    (("Daily",), "U-UN.TO adj. close", "CLOSE", "TSX close, about 4 pm ET."), (("Daily",), "SPUT price", "CLOSE", "TSX close reported by Sprott, about 4 pm ET."),
    (("Daily",), "U-U.TO adj. close", "CLOSE", "TSX close, about 4 pm ET."),
    (("Daily",), "SPUT NAV per unit", "CLOSE", "Sprott NAV struck about 4 pm ET; unchanged on 216 days."), (("Daily",), "SPUT premium/discount", "CLOSE", "From the Sprott NAV struck about 4 pm ET."),
    (("Daily",), "WTI crude, spot", "DAY", "EIA daily spot assessment (FRED)."), (("Daily",), "Brent crude, spot", "DAY", "EIA daily spot assessment (FRED)."),
    (("Daily",), "Henry Hub natural gas, spot", "DAY", "EIA daily spot assessment (FRED)."),
    (("Daily",), "futures (", "CLOSE", _FUT_D, "contains"),
    (("Daily",), "Fed funds rate", "DAY", "Effective fed funds rate for the day (FRED); the source carries values over weekends and holidays."),
    (("Daily",), "Funding spread", "CONST", "Commercial paper minus T-bill, built here for each day both exist."),
    (("Daily",), "Broad dollar index", "DAY", "Fed H.10 noon New York rates (grade B, not checked in a kept file)."), (("Daily",), "USD/CAD", "DAY", "Fed H.10 noon New York rates (grade B, not checked in a kept file)."),
    (("Daily",), "GBP/USD", "DAY", "Fed H.10 noon New York rates (grade B, not checked in a kept file)."),
    (("Daily",), "VIX", "CLOSE", "Cboe VIX close (FRED)."),
    (("Daily",), "Economic policy uncertainty, US, daily", "DAY", "Daily newspaper-count index, 7 days a week."),
    (("Daily",), "Geopolitical risk index, daily", "DAY", "Daily newspaper-count index, calendar days; weekend values from 1990."),
    (("Daily",), "UX futures volume, CME bulletin", "FLOW", "Contracts traded on the settlement day (CME daily bulletin)."),
    (("Daily",), "UX futures open interest, CME bulletin", "DAY", "Open interest at the end of the settlement day (CME daily bulletin)."),
    (("Daily",), "UX front-month settlement, CME bulletin", "CLOSE", "Daily settlement (CME daily bulletin). " + _UX),
    (("Daily",), "UX futures close, daily (Stooq)", "CLOSE", "Stooq UX.F continuous close, roll undocumented. " + _UX),
    (("Daily",), "UX futures close, daily (Investing UXXc2", "CLOSE", "Investing.com UXXc2 (second contract) close, kept apart (D-035). " + _UX),
    (("Daily",), "S&P 500 (FRED", "CLOSE", "S&P 500 close (FRED)."),
    (("Daily",), "Baa-10y credit spread", "DAY", _FRED_D),
    (("Daily",), "3-month Treasury yield", "DAY", _FRED_D), (("Daily",), "US 10y", "DAY", _FRED_D), (("Daily",), "Term spread 10y-3m", "DAY", _FRED_D),
    (("Daily",), "Breakeven inflation", "DAY", _FRED_D), (("Daily",), "Corporate option-adjusted spread", "DAY", _FRED_D), (("Daily",), "TED spread", "DAY", _FRED_D),
    (("Daily",), "Commercial paper", "DAY", _FRED_D), (("Daily",), "T-bill, 3-month", "DAY", _FRED_D), (("Daily",), "SOFR", "DAY", _FRED_D),
    (("Daily",), "YCA.L adj. close", "CLOSE", "LSE close, hours before North American closes."),
    (("Daily",), "adj. close", "CLOSE", _EQ_D, "contains"), (("Daily",), "S&P 500 (^GSPC) close", "CLOSE", _EQ_D),
    # Annual
    (("Annual",), "US uranium concentrate production", "FLOW", "Total over the calendar year (EIA DUPR)."),
    (("Annual",), "US uranium concentrate shipments", "FLOW", "Total over the calendar year (EIA DUPR)."),
    (("Annual",), "US producer sales deliveries", "FLOW", "Total over the calendar year (EIA DUPR)."),
    (("Annual",), "US producer weighted-average sales price", "AVG", "Weighted average over the calendar year (EIA DUPR)."),
    (("Annual",), "inventory", "YE", "Stock at year-end (EIA UMAR).", "contains"),
    (("Annual",), "US reactor requirements", "REL", "Forward-looking figures by year, as of the date in the header (EIA survey)."),
    (("Annual",), "US unfilled requirements", "REL", "Forward-looking figures by year, as of the date in the header (EIA survey)."),
    (("Annual",), "US contracted deliveries", "REL", "Forward-looking figures by delivery year, as of the date in the header (EIA survey)."),
    (("Annual",), "World nuclear electricity produced", "FLOW", "Total over the calendar year (IAEA PRIS)."),
    (("Annual",), "World nuclear operated capacity", "IDX", "Net capacity of the reactors PRIS counts as operated in the year, not year-end capacity."),
    (("Annual",), "World nuclear reactors in operation, year-end", "YE", "Count at year-end (IAEA RDS-2)."),
    (("Annual",), "World nuclear construction starts", "FLOW", "Count over the calendar year (IAEA RDS-2)."),
    (("Annual",), "World nuclear grid connections", "FLOW", "Count or capacity over the calendar year (IAEA RDS-2)."),
    (("Annual",), "World nuclear permanent shutdowns", "FLOW", "Count or capacity over the calendar year (IAEA RDS-2)."),
    (("Annual",), "World nuclear", "IDX", "Ratio for the calendar year (IAEA PRIS)."),
    # World Bank
    (("World Bank monthly",), "", "AVG", "Monthly average as published (World Bank Pink Sheet); the indices are built from monthly averages."),
]


def timing_of(sheet, header):
    """Timing code and note of a data column. 19 Nov 2026: a column with no rule gets CHECK, so nothing is labelled by default."""
    for rule in _TIMING_RULES:
        sheets, key, code, note = rule[:4]
        if sheet not in sheets:
            continue
        hit = (key in header) if len(rule) > 4 else header.startswith(key)
        if hit:
            return code, note
    return "CHECK", "No timing rule matches this column yet."


def load_release(path):
    """19 Nov 2026: read data/release_dates.csv (a register of release dates) to label the Release row; no data value is read from it."""
    import csv
    RELEASE.clear()
    if not path.exists():
        return
    by = {}
    for r in csv.DictReader(open(path, encoding="utf-8")):
        by.setdefault(r["series"], []).append(r)
    for s, rs in by.items():
        n = sum(1 for r in rs if r["release_date"])
        if n == len(rs):
            RELEASE[s] = (f"Dated: {n} release dates in data/release_dates.csv", None)
        elif n:
            RELEASE[s] = (f"Dated: {n} of {len(rs)} periods in data/release_dates.csv (rest not verified)", None)
        elif all(r["note"].startswith("typical lag") for r in rs):
            RELEASE[s] = ("Not dated: typical lag in data/release_dates.csv", rs[0]["note"] + ". Source: " + rs[0]["release_date_source"])
        else:
            RELEASE[s] = ("Not dated (not verified): see data/release_dates.csv", rs[0]["note"])


def release_of(sheet, header):
    if header in RELEASE:
        return RELEASE[header]
    if sheet == "World Bank monthly" and "World Bank monthly (all columns)" in RELEASE:
        return RELEASE["World Bank monthly (all columns)"]
    return ("", None)


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
    # 19 Nov 2026: rows Timing and Release added between Type · coverage and the unit row (labels only; data now start in row 13).
    labels = {4: "Role in the thesis", 5: "Variable", 6: "Why it matters", 7: "Series", 8: "Source", 9: "Type · coverage", TIMING_ROW: "Timing", RELEASE_ROW: "Release", KEY_ROW: keyhdr}
    for r, lab in labels.items():
        c = ws.cell(r, 1, lab)
        c.font = font(9 if r < KEY_ROW else 10, bold=True, color=CHAR if r == KEY_ROW else GREY)
        c.alignment = Alignment(vertical="center", wrap_text=True)
        c.fill = fill(IVORY if r != KEY_ROW else "FFFFFF")
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
        code, tnote = timing_of(name, short(c["header"]))
        if id_:
            TIMING_BY_ID.setdefault(id_, []).append((name, code))
        rtext, rnote = release_of(name, short(c["header"]))
        row = {5: var, 6: why_col(id_, c["header"]), 7: short(c["header"]), 8: c["source"], 9: cov, TIMING_ROW: code, RELEASE_ROW: rtext, KEY_ROW: c["unit"]}
        for r, v in row.items():
            cell = ws.cell(r, j, v)
            cell.alignment = Alignment(wrap_text=True, vertical="top" if r < KEY_ROW else "center", horizontal="center" if r == KEY_ROW else "left")
            cell.font = font(10 if r in (5, 7) else 9, bold=(r in (5, 7, TIMING_ROW)), italic=(r == 6), color=CHAR if r in (5, 7, KEY_ROW, TIMING_ROW) else (FOREST if r == 6 else GREY))
            cell.fill = fill(IVORY)
        ws.cell(9, j).fill = fill(TYPE_FILL.get(c["kind"], SAGE))
        ws.cell(TIMING_ROW, j).comment = Comment(f"{code}: {TIMING_TEXT[code]} {tnote}".strip(), "Dataset", width=340, height=120)
        if rnote:
            ws.cell(RELEASE_ROW, j).comment = Comment(rnote, "Dataset", width=340, height=120)
        ws.cell(KEY_ROW, j).fill = fill("FFFFFF")
        ws.cell(KEY_ROW, j).border = Border(bottom=Side(style="medium", color=FOREST))
        note = " ".join(x for x in [c.get("note", ""), f"File: {c['raw']}" if c.get("raw") and c["raw"] != "see note" else ""] if x)
        if note:
            ws.cell(7, j).comment = Comment(note, "Dataset", width=380, height=140)
        ws.column_dimensions[get_column_letter(j)].width = width
    if prev is not None:
        _band(ws, start, len(cols) + 1, prev)
    ws.cell(KEY_ROW, 1).border = Border(bottom=Side(style="medium", color=FOREST))
    for key in keys:
        ws.append([keyval(key)] + [c["values"].get(key) for c in cols])
    from copy import copy
    f10 = font()
    tmpl = ws.cell(DATA_ROW, 1)
    tmpl.number_format, tmpl.font = keyfmt, font(bold=True, color=CHAR)
    styles = [copy(tmpl._style)]
    for j, c in enumerate(cols, 2):
        t = ws.cell(DATA_ROW, j)
        t.font, t.number_format = f10, c["fmt"]
        styles.append(copy(t._style))
    for r in range(DATA_ROW, ws.max_row + 1):
        for j, st in enumerate(styles, 1):
            ws.cell(r, j)._style = copy(st)
    ws.column_dimensions["A"].width = 16
    for r, h in {1: 24, 4: 20, 5: 62, 6: 78, 7: 48, 8: 30, 9: 28, TIMING_ROW: 18, RELEASE_ROW: 30, KEY_ROW: 30}.items():
        ws.row_dimensions[r].height = h
    ws.freeze_panes = ws.cell(DATA_ROW, 2)
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


SPOT_LINK = {}


def load_spot_link(path):
    """19 Nov 2026: labels from the Market Expert's screen (data/spot_link.csv): does a variable contain the uranium spot price by construction? Labels only, no calculation."""
    import csv as _csv
    SPOT_LINK.clear()
    with open(path, newline="", encoding="utf-8") as f:
        for row in _csv.DictReader(f):
            SPOT_LINK[row["variable_id"]] = row


def spot_link_cell(id_):
    row = SPOT_LINK.get(id_)
    if not row:
        return None
    txt = row["spot_link"]
    alt = (row["clean_alternative"] or "").strip()
    if alt and not alt.lower().startswith("none"):
        txt += " · " + alt
    return txt


def timing_cell(id_):
    """19 Nov 2026: timing codes of the data columns tagged with this ID, e.g. 'Monthly ME-ASSESSED; Quarterly QE-ASSESSED' (blank if no column)."""
    seen = []
    for sh, code in TIMING_BY_ID.get(id_, []):
        t = f"{sh} {code}"
        if t not in seen:
            seen.append(t)
    return "; ".join(seen) or None


def write_variables_sheet(wb, rows_by_id, extras):
    """rows_by_id: id -> (workbook status, where, note). extras: list of (id, variable, variant, status, where, note)."""
    ws = wb.create_sheet("Variables")
    ws.sheet_properties.tabColor = TAB["guide"]
    _title(ws, "Variables", "One row per variable of the handoff table, in its order, with its role in the thesis and what this workbook holds. The grey block on the right keeps the handoff's own columns for reference.")
    # 19 Nov 2026: column "Timing" (codes of the columns tagged with the ID, by sheet) added after Note; the handoff block moves one column right.
    heads = ["ID", "Variable", "Variant", "Role in the thesis", "Why it matters", "In this workbook", "Where in the workbook", "Note", "Timing", "Spot link",
             "Source (handoff)", "Status (handoff)", "Remaining issue (handoff)", "Where / how to retrieve (handoff)", "Python (handoff)"]
    widths = [8, 22, 32, 22, 56, 18, 48, 56, 26, 30, 28, 12, 40, 32, 22]
    r = 5
    for j, h in enumerate(heads, 1):
        c = ws.cell(r, j, h)
        c.font, c.fill = font(10, bold=True, color="FFFFFF"), fill(FOREST if j <= 5 else (GOLD if j <= 10 else GREY))
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 30
    from openpyxl.comments import Comment as _Comment
    ws.cell(r, 10).comment = _Comment("Spot link (Market Expert screen, 19 Nov 2026; labels only, no calculation; file data/spot_link.csv). A = contains the uranium spot price by construction (never use it to predict the spot; 'A (target)' is the spot itself). B = partly linked (moves on the same news or shares inputs; use only lagged and with care). C = no construction link. Mixed rows give the mark per part. The text after the mark gives the Expert's advice, e.g. the clean alternative.", "Dataset", width=380, height=170)
    r += 1

    def _asof(r, id_):
        row = SPOT_LINK.get(id_)
        if row and row.get("as_of_day_note"):
            ws.cell(r, 9).comment = _Comment("Which day the month-end stands for: " + row["as_of_day_note"], "Dataset", width=380, height=150)

    def put(r, vals):
        for j, v in vals.items():
            c = ws.cell(r, j, v)
            c.font = font(9 if j > 3 else 10, bold=(j == 2), color=(FOREST if j == 5 else (GREY if j >= 11 else CHAR)), italic=(j == 5))
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
            put(r, {1: id_, 2: var, 3: variant, 4: role_label(id_), 5: why_col(id_, ""), 6: wst, 7: wwhere, 8: wnote, 9: timing_cell(id_), 10: spot_link_cell(id_), 11: src, 12: st, 13: issue, 14: where, 15: py})
            _asof(r, id_)
            r += 1
    ws.cell(r, 1, "Additional items from the variables document (not in the handoff table)").font = font(11, bold=True, color=FOREST)
    for j in range(1, len(heads) + 1):
        ws.cell(r, j).fill = fill(SAGE)
    r += 1
    for id_, var, variant, wst, wwhere, wnote in extras:
        put(r, {1: id_, 2: var, 3: variant, 4: roles.ROLES["H. Extra items from the variables doc"][0], 5: why_col(id_, ""), 6: wst, 7: wwhere, 8: wnote, 9: timing_cell(id_), 10: spot_link_cell(id_)})
        r += 1
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    for j in range(11, 16):  # handoff columns: reference only, grouped and collapsed
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
        # 19 Nov 2026: say what the two new header rows mean and how the Monthly date column is to be read.
        "Row Timing gives one code per column saying what kind of number it is (codes below; hover for the convention). Row Release says whether the dates on which the figures became public are in data/release_dates.csv, so a predictive model does not use a figure before it was known.",
        "The Monthly date column shows the 1st of each month for every row, while most values are month-end or for the whole month: join it to month-end-dated data by year and month, not by day.",
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
    # 19 Nov 2026: the timing vocabulary (row "Timing" of the data sheets, column "Timing" of Variables).
    ws.cell(r, 1, "Timing codes (row “Timing”)").font = font(11, bold=True, color=FOREST)
    r += 1
    for k, t in TIMING_CODES:
        ws.cell(r, 1, k).font = font(9, bold=True)
        c = ws.cell(r, 2, t)
        c.font = font(9)
        c.alignment = Alignment(wrap_text=True, vertical="top")
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
