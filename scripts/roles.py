"""Context texts for thesis_dataset.xlsx: the role of each group of variables, why a sheet matters, and why additional series are kept.

Roles follow the grouping of the handoff table (coverage_map.py, column 'step'). No hypothesis is stated or implied here:
the texts say what a series measures and why a study of uranium financialization would look at it.
"""

# step label in coverage_map.py -> (short role name, why this group matters)
ROLES = {
    "A. Uranium price (the dependent variable)": ("A. Uranium price", "The price the thesis studies. Spot and long-term assessments separate prompt pricing from contract-market pricing."),
    "A. Fuel-chain prices (conversion, enrichment)": ("A. Fuel-chain prices", "Conversion and enrichment prices show conditions elsewhere in the nuclear fuel chain, so uranium-specific moves can be told apart from chain-wide ones."),
    "B1. SPUT secondary trading": ("B1. SPUT trading", "How actively SPUT units trade on the exchange: the financial-market side of the physical-uranium trust."),
    "B2. SPUT premium or discount to NAV": ("B2. SPUT premium/discount", "The gap between SPUT's market price and the value of the uranium it holds. A premium is the condition under which the trust can sell units and buy uranium."),
    "B3. SPUT unit issuance and proceeds": ("B3. SPUT issuance", "How SPUT raises money by selling new units: the link from investor demand to uranium purchases."),
    "B4. SPUT cash and purchase capacity": ("B4. SPUT cash", "Cash raised but not yet spent on uranium: how much buying SPUT can still do."),
    "B5. SPUT physical acquisitions and holdings": ("B5. SPUT holdings", "Pounds of uranium SPUT buys and holds: the physical footprint of the financial vehicle in the spot market."),
    "B6. SPUT scaled participation": ("B6. SPUT participation", "SPUT trading and buying scaled by units or market volume: whether the vehicle is large enough to matter for the market."),
    "C. Comparison vehicle: Yellow Cake": ("C. Yellow Cake", "A second listed physical-uranium vehicle (London). It checks whether SPUT findings are specific to one fund."),
    "D. Comparison commodities": ("D. Comparison commodities", "Gold, oil, gas, coal and copper: benchmarks to separate uranium-specific behaviour from general commodity trends."),
    "E. Macro-financial and uncertainty controls": ("E. Macro-financial controls", "Interest rates, the dollar, inflation, uncertainty and activity: the wider financial conditions that move all commodity prices."),
    "F. Physical market conditions": ("F. Physical market", "Production, inventories, requirements and contracts: the supply and demand fundamentals that financial activity has to be separated from."),
    "G. Constructed measures": ("G. Constructed measures", "Returns, volatility, correlations and spillovers built from the series above. Computed in the analysis, not stored here."),
    "H. Extra items from the variables doc": ("H. Extra items", "Additional items from the thesis variables document: earlier fund, attention, shipping, events."),
}

# Purpose of the D-items (the handoff table has none); V-items use coverage_map's own 'purpose' text
D_PURPOSE = {
    "D13": "Period markers (SPUT launch, war, post-2005) so results can be compared across regimes.",
    "D17": "Global equity returns: general risk appetite, a control for financial conditions.",
    "D18": "Global shipping and trade activity: a demand-side control that moves with commodity demand.",
    "D21": "Investor attention to uranium ETFs: a proxy for retail financial interest.",
    "D27": "Dates of supply and geopolitical shocks for event windows.",
    "D34": "Tests for explosive price behaviour, computed from the uranium price.",
    "D35": "Futures minus spot: the market's forward-looking premium over the physical price.",
    "D44": "Uranium Participation Corp, the listed uranium vehicle before SPUT: extends the premium/discount history back to 2010.",
    "D45": "A flag for months when SPUT traded below NAV, when it could not issue units.",
    "D47": "Dated SPUT sales of new units: the moments when investor demand became uranium purchases.",
    "D48": "A broker-based daily uranium price and the price implied by the fund: a daily alternative to the monthly spot.",
}

# Additional series (no handoff ID): header keyword -> why it is kept
EXTRA_WHY = [
    ("Uranium (IMF)", "Second uranium price (IMF monthly average): a cross-check on the Cameco series."),
    ("UX futures", "Uranium futures price: the forward market, to compare with spot."),
    ("UX final settlement", "Reconstructed futures settlement used only to fill months missing from Stooq."),
    ("Silver", "Another precious-metal benchmark for the commodity comparison."),
    ("Platinum", "Another precious-metal benchmark for the commodity comparison."),
    ("Baa-10y", "Credit-market stress: corporate yield over the Treasury yield."),
    ("S&P 500", "Global equity returns: general risk appetite, a control for financial conditions."),
    ("Geopolitical risk index, historical", "Geopolitical risk with a longer history (back to 1900), for supply-shock context."),
    ("Google Trends", "Search interest in uranium ETFs: a proxy for investor attention."),
]
EXTRA_DEFAULT = "Additional series kept for reference; no handoff record."

# One line per sheet: why it matters for the thesis
SHEET_WHY = {
    "Monthly": "The main analysis frequency: uranium prices lined up with SPUT activity, comparators, rates and risk measures, month by month.",
    "Quarterly": "Yellow Cake and conversion prices are published quarterly; this sheet lines them up with SPUT purchases and the size of the spot market.",
    "Annual": "Supply and demand fundamentals that exist only yearly (production, inventories, requirements): they show whether price moves match the physical market.",
    "Weekly": "SPUT's holdings of uranium, published weekly: the physical footprint of the financial vehicle.",
    "Daily": "Daily prices, rates and risk measures for event studies and volatility work; the Monthly sheet takes its month-end values from here.",
    "SPUT purchases": "Every day SPUT bought uranium and how much: the direct link between financial demand and the physical market.",
    "SPUT issuance": "When and at what price SPUT sold new units: shows when investor demand turned into uranium purchases.",
    "SPUT statements": "Audited and interim accounts: exact units, NAV, cash and proceeds at reporting dates, the anchors for the daily and weekly series.",
    "UxC volumes": "Size of the spot market as quoted by UxC: the yardstick for how large SPUT's buying is.",
    "Yellow Cake reports": "The second listed physical-uranium vehicle: its purchases, financing and cash, as a comparison for SPUT.",
    "UPC monthly NAV": "Uranium Participation Corp, the listed uranium vehicle before SPUT: extends the fund history back to 2010.",
    "Physical sources": "Production, requirements and inventories by country and year, from official sources: the physical fundamentals behind prices.",
    "Prices OHLCV": "Open, high, low, close and volume for every traded ticker: the source of the price columns on the Daily and Monthly sheets.",
    "World Bank monthly": "Monthly average prices for coal, oil, gas, metals and fertiliser: comparators where no month-end series exists.",
    "Variables": "The register: what each variable measures, why it matters to the thesis and whether this workbook holds it.",
    "Gaps": "What is missing or weak, and what each gap blocks, so it is clear what the data can and cannot support.",
}
