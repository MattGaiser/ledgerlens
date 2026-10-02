# Financial evidence

The bundled snapshot contains 81 facts: MSFT, AAPL, and NVDA; FY2023–FY2025; revenue, gross profit, operating income, net income, operating cash flow, capital expenditure, diluted EPS, total assets, and cash/equivalents. It was acquired on 2026-10-02 from the [SEC companyfacts API](https://www.sec.gov/search-filings/edgar-application-programming-interfaces).

`data/financials.json` records the original response URL, SHA-256, acquisition time, and normalized facts. Original responses are included in the source archive under `data/raw`. `node scripts/verify-sources.mjs` independently reconciles all source fields, scaling, annual durations, eligible filing dates, and accession URLs. Live sync revalidates the selected company's 27 facts and persists them locally.

| Company | Fiscal calendar | FY2025 revenue (USD millions) |
| --- | --- | ---: |
| Microsoft | June | 281,724 |
| Apple | September | 416,161 |
| NVIDIA | January | 130,497 |

All currency metrics use USD millions. Diluted EPS uses USD/share. Each source keeps its raw unscaled value. Annual flow facts must span 330–380 days; balance-sheet facts are instants. The supported issuers' fiscal-year labels follow the year of their reporting end date, not the XBRL record's filing-context `fy` field. Latest available eligible 10-K comparative values are selected as of acquisition, so a FY2025 fact can be cited to a 2026 filing.

NVIDIA's capital-expenditure fallback uses `PaymentsToAcquireProductiveAssets`; Microsoft and Apple use `PaymentsToAcquirePropertyPlantAndEquipment`. These labels are similar but not identical. Operating cash flow less capex is a calculated measure, not a standardized GAAP line item. Different fiscal calendars further limit direct comparisons. Forecasts are scenario calculations from user assumptions and must not be described as reported results or company guidance.

The universe and fiscal mapping are deliberately bounded. Adding companies requires validation of fiscal naming, concepts, currency, duration, restatements, and units. The current live sync uses each metric's established concept; it does not silently switch concepts when an issuer changes reporting.
