# LedgerLens

**A sourced financial research workspace inside Microsoft Excel.** Built as an independent pre-interview demonstration for the AlphaSense Excel engineering role, using public SEC filings and optional OpenAI research.

![LedgerLens running in Excel](docs/images/excel-workspace.png)

The workflow is simple: inspect a financial fact, follow its filing evidence, research a company, preview nine historical model updates, and apply them without overwriting analyst assumptions or forecast formulas. Every reviewed import appends a durable source audit to the workbook. Undo checks for subsequent edits before restoring values.

## Run the Windows release

1. Extract the entire `LedgerLens-1.0.2-windows.zip` into a writable local folder.
2. Double-click **Start LedgerLens.cmd**. It starts the local service and opens a new analyst workbook with the research pane.
3. On **Model**, choose **LedgerLens → Review updates**, preview the nine reported values, then apply. Change the blue assumptions to explore forecasts.
4. Use **Research** for cited explanations and **Health** to demonstrate offline mode, recovery, and notifications.

Windows desktop Excel, .NET Framework 4.8, and Microsoft Edge WebView2 Runtime are required. The service runtime is included; Visual Studio, Node, and the .NET SDK are unnecessary for running the release. The validated host is Microsoft 365 Excel x64, version 16.0, Application.Build 20430. The native add-in is unsigned. If your organization's policy blocks unsigned XLLs, use the browser preview or arrange an approved development environment; the launcher does not change Office security policy.

The existing user-level `OPENAI_API_KEY` is read locally by the service. Without a key, the app provides explicitly labeled calculated analysis. Use OpenAI is an explicit option; `LL.ASK` also requests AI and may incur usage charges. No key is included in the workbook, source, or release. See [setup and troubleshooting](docs/OPERATIONS.md).

```powershell
.\Start-LedgerLens.ps1 -Browser -NoExcel  # Browser preview
.\Start-LedgerLens.ps1 -Validate          # Real Excel launch/save/close check
.\Stop-LedgerLens.ps1                     # Stop service; leave workbooks open
```

## What to demonstrate

| Job requirement | Working evidence |
| --- | --- |
| C# desktop integration | .NET Framework 4.8 XLL, ribbon, WebView2 task pane, Excel COM adapter |
| Calculation engine and performance | Async UDFs, shared bounded caches, cancellation, 2,000-cell coalescing test |
| Complex state and recovery | Preview dependencies, optimistic conflicts, transactional writes, guarded rollback |
| Financial workflow | Three companies, three fiscal years, nine metrics, scenarios, charts, filing audit |
| Cloud and AI integration | Out-of-process .NET 10 API, live SEC refresh, OpenAI Responses, validated citation IDs |
| Reliability and testing | Offline snapshots, bounded retries, circuit breaker, WebSocket cleanup, real Excel automation |
| Cross-platform strategy | Shared web UI and tested Office.js adapter; actual Mac/Office.js host validation remains pending |

Example formulas:

```excel
=LL.METRIC("MSFT","Revenue","FY2025")
=LL.SOURCE("MSFT","Revenue","FY2025")
=LL.STATUS("MSFT","Revenue","FY2025")
=LL.TABLE("AAPL","FY2025")
=LL.ASK("MSFT","Compare FY2024 and FY2025 operating margins")
=LL.LIVE()
```

`LL.METRIC` returns USD millions, except `DilutedEPS` in USD/share. `LL.TABLE` returns an array; spill behavior depends on the Excel version. `LL.LIVE` displays service notifications, not security prices. Optional formula revision arguments allow explicit refresh without volatile UDFs.

## Evidence and review

- [Three-minute demo](docs/DEMO.md)
- [Architecture and design decisions](docs/ARCHITECTURE.md)
- [Validation report and compatibility scope](docs/VALIDATION.md)
- [Code review and regression fixes](docs/CODE-REVIEW.md)
- [Office.js preview](docs/OFFICE-PREVIEW.md)
- [Data provenance](docs/DATA.md)

This is a portfolio prototype, not an AlphaSense product or a representation of its internal implementation. Enterprise SSO, AWS deployment, signed installation, managed updates, and organization-wide rollout are design discussion topics, not implemented features. The project was developed with substantial AI assistance; review and understand the implementation before presenting it as your work.

## Build and test

The source package includes .NET and npm dependency lockfiles. Use the x64 .NET SDK pinned by `global.json`, Node 22+, and Windows PowerShell 5.1. Edge is used by Playwright.

```powershell
npm ci
.\build.ps1 -SelfContained
.\Start-LedgerLens.ps1 -NoExcel
.\scripts\test.ps1
# Real Excel/WebView2 integration, run separately from service load tests:
.\scripts\test.ps1 -Native
.\scripts\package.ps1
.\scripts\test-release.ps1  # Fresh extraction, real Excel launch, saved formula inspection
```

The native test starts and closes its own Excel instance. It enables WebView2 debugging only inside the test harness. Performance tests must run without another LedgerLens client issuing requests.

A pre-interview project is a small work sample tailored to an employer's actual problem, prepared to demonstrate how you would contribute. This follows the idea of a [value validation project](https://cultivatedculture.com/value-validation-projects-my-best-job-search-strategy-ep-20/): show a concrete result the team can inspect. Here the strongest story is safe, sourced model updates under unreliable network conditions.
