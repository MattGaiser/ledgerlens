# LedgerLens delivery contract

An independent pre-interview demonstration for the AlphaSense Excel engineering role. Real Windows Excel integration, public SEC financial evidence, and OpenAI research. No AlphaSense affiliation or access is implied.

## Required product experience

- A styled analyst workbook with company comparisons, historical financials, scenario inputs, forecast formulas, charts, source records, and an operating guide.
- Native C# Excel-DNA add-in: ribbon, modern WebView2 research pane, asynchronous financial UDFs, source and freshness UDFs, explicit AI research, streaming update subscriptions, and cancellation.
- Trace every numeric fact to its SEC accession, reporting dates, XBRL concept, units, filing URL, and acquisition timestamp. Never present synthetic scenario adjustments as reported financials.
- Review proposed model changes before applying. Detect edits made after preview, preserve forecast formulas and analyst inputs, restore Excel state on every exit path, and support guarded rollback of the last applied transaction.
- Out-of-process C# service with API contracts reusable by an Office.js client. Ship a working browser interface and Office.js adapter/manifest; describe the Mac validation limitation honestly.
- Ground AI answers in supplied facts, validate citation IDs, expose provider/model and provenance, keep keys out of workbooks/logs/bundles, and support a clearly labeled deterministic explanation when AI is unavailable.
- Offline cache, explicit freshness, timeouts, bounded concurrency, request coalescing, bounded retries, circuit breaker, event streaming, and fault injection observable from a diagnostics page.
- Simple launch/stop/build/test/package scripts, versioned artifacts, clean source repository, screenshots, concise demo script, architecture/deployment notes, and machine-readable validation results.

## Completion gates (evidence required)

1. Build all projects with warnings as errors. Pin and lock dependencies; inspect dependency vulnerabilities.
2. Unit/contract tests: normalization, period selection, units, cache freshness, concurrency/coalescing, cancellation, retries/circuit recovery, AI citation checks, JSON parsing, and refresh conflict/rollback invariants.
3. Service integration tests: real HTTP authentication, invalid input, cache/offline behavior, fault injection/recovery, streaming disconnect/reconnect, AI failure responses, and request limits.
4. Native Excel tests: load XLL; asynchronous formulas complete; malformed requests show actionable errors; workbook close cancels work; pane opens; ribbon commands work; refresh/rollback preserve user data; host remains responsive.
5. Browser tests: keyboard navigation, narrow pane layout, task flows, source links, honest failure states, and visual screenshots. Exercise the actual Excel bridge as well as browser fallbacks.
6. Performance: cold/warm requests, repeated keys across thousands of cells, bounded backend requests, latency during slow provider operation, repeated lifecycle memory observations. Report measured results and scope, not invented targets or universal claims.
7. Live OpenAI generation succeeds with valid citations, and negative tests cover unavailable key, rejected key, timeout, refusal, and malformed output without disclosing secrets.
8. Independent source checks reconcile workbook values to downloaded SEC evidence; verify fiscal year and reporting currency explicitly.
9. Start from a freshly extracted release folder and run the documented launch workflow. Inspect the packaged workbook, UI, logs, source, and test evidence; remove transient files and secrets from the release.

## Architecture

`Excel (net48 / Excel-DNA / WebView2) -> loopback authenticated HTTP + WebSocket -> .NET 10 service -> SEC + OpenAI`

`LedgerLens.Core` holds contracts, financial validation, cache coordination, and refresh planning without an Excel dependency. Only the Excel adapter may touch the COM object model; writes run on Excel's main thread. Web UI is shared between the task pane and browser; Office.js is a separate bridge.

## Completed evidence

All nine delivery gates have evidence in [VALIDATION.md](VALIDATION.md) and the packaged machine-readable reports.

1. Native x64/x86 and self-contained Windows service build successfully with warnings as errors; dependency audits report no known vulnerabilities.
2. 36 .NET unit/contract tests pass, including cache, cancellation, source selection, state, and AI failure fixtures.
3. 13 service checks pass, including authentication, request limits, Host/Origin protection, offline recovery, and WebSocket reconnect/cleanup.
4. 14 real Excel/WebView2 checks pass, including workbook-close cancellation, formulas, guarded updates, audit, rollback, and lifecycle.
5. Eight browser tests and nine Office.js adapter tests pass; native UI screenshots were inspected. Office manifest validation passes.
6. The 2,000-formula benchmark coalesces to one provider request; latency and four lifecycle samples are recorded with their measurement scope.
7. Live OpenAI generation succeeds; cited numerical claims were inspected. Negative fixtures cover unavailable, rejected, timeout, refusal, incomplete, and malformed responses.
8. All 81 facts reconcile independently to three hashed original SEC responses. Live SEC sync revalidated 27 Microsoft facts.
9. A freshly extracted self-contained release, in a path containing spaces, loaded the XLL, created the workbook, resolved a financial formula, saved/reopened it, opened WebView2, and unloaded cleanly. The example workbook has six sheets, 64 formulas, and no cached formula errors. Package credentials are scanned and file hashes recorded.

The tested product target is Windows Microsoft 365 Excel x64. Mac/Office.js host operation, older Excel versions, enterprise SSO, AWS, signed deployment, and managed updates remain explicitly outside the validated prototype scope. They are not represented as completed product capabilities.
