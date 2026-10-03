# Financial evidence

The bundled snapshot contains 81 facts: MSFT, AAPL, and NVDA; FY2023–FY2025; revenue, gross profit, operating income, net income, operating cash flow, capital expenditure, diluted EPS, total assets, and cash/equivalents. It was acquired on 2026-10-02 from the [SEC companyfacts API](https://www.sec.gov/search-filings/edgar-application-programming-interfaces).

`data/financials.json` records the original response URL, SHA-256, acquisition time, and normalized facts. Live sync revalidates the selected company's 27 facts and persists them locally. The default tests use this bundled snapshot and small test fixtures; they require no SEC downloads.

Original-response reconciliation is optional. If you have the archived companyfacts responses matching the hashes in `data/financials.json`, place them in `data/raw` as `msft.json`, `aapl.json`, and `nvda.json`, then run `node scripts/verify-sources.mjs` or `.\scripts\test.ps1 -OriginalSources`. This checks source fields, scaling, annual durations, eligible filing dates, and accession URLs. The archives are excluded from Git and release packages. New SEC downloads can differ from the recorded hashes as filings change; they are not substitutes for the original inputs.

| Company | Fiscal calendar | FY2025 revenue (USD millions) |
| --- | --- | ---: |
| Microsoft | June | 281,724 |
| Apple | September | 416,161 |
| NVIDIA | January | 130,497 |

All currency metrics use USD millions. Diluted EPS uses USD/share. Each source keeps its raw unscaled value. Annual flow facts must span 330–380 days; balance-sheet facts are instants. The supported issuers' fiscal-year labels follow the year of their reporting end date, not the XBRL record's filing-context `fy` field. Latest available eligible 10-K comparative values are selected as of acquisition, so a FY2025 fact can be cited to a 2026 filing.

NVIDIA's capital-expenditure fallback uses `PaymentsToAcquireProductiveAssets`; Microsoft and Apple use `PaymentsToAcquirePropertyPlantAndEquipment`. These labels are similar but not identical. Operating cash flow less capex is a calculated measure, not a standardized GAAP line item. Different fiscal calendars further limit direct comparisons. Forecasts are scenario calculations from user assumptions and must not be described as reported results or company guidance.

The universe and fiscal mapping are deliberately bounded. Adding companies requires validation of fiscal naming, concepts, currency, duration, restatements, and units. The current live sync uses each metric's established concept; it does not silently switch concepts when an issuer changes reporting.
