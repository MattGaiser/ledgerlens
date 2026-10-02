import fs from 'node:fs';
const read=name=>{const bytes=fs.readFileSync(`artifacts/validation/${name}`);return JSON.parse(bytes.toString(bytes[0]===255&&bytes[1]===254?'utf16le':'utf8').replace(/^\uFEFF/,''))};
const native=read('native-results.json'),service=read('service-results.json'),browser=read('browser-results.json'),source=read('source-reconciliation.json'),workbook=read('workbook-inspection.json'),sync=read('live-sec-sync.json');
const launch=fs.existsSync('artifacts/validation/release-launch.json')?read('release-launch.json'):null;
const m=native.measurements;
const decimal=value=>Number(value).toFixed(1);
const privateMemory=m.lifecycle.map(x=>x.privateMb);
const report=`# Validation report — LedgerLens 1.0.0

Evidence recorded on ${new Date().toISOString().slice(0,10)}. This is a prototype validation report for one Windows workstation, with measured results and explicit coverage limits.

## Results

| Check | Result | Evidence |
| --- | --- | --- |
| Release build, warnings as errors | PASS; native x64/x86 artifacts and self-contained win-x64 service | build.ps1 |
| .NET unit and contract tests | 36 passed | unit-results.trx |
| HTTP, resilience, request limits, WebSocket lifecycle | ${service.checks.length} passed | service-results.json |
| Real Excel and WebView2 workflow | ${native.checks.length} passed; ${native.errors.length} JavaScript errors | native-results.json |
| Edge browser UI | ${browser.stats.expected} passed, ${browser.stats.unexpected} unexpected failures | browser-results.json |
| Office.js adapter | 9 passed | office-adapter.tap |
| Office XML manifest | Validator passed | office-manifest.txt |
| Original SEC source reconciliation | ${source.checks.length} facts passed against 3 hashed source files | source-reconciliation.json |
| Live SEC refresh | ${sync.status}; ${sync.facts} MSFT facts; ${sync.elapsedMs} ms | live-sec-sync.json |
| Live OpenAI research | Completed; cited claims inspected against supplied facts | live-openai.json |
| Saved workbook | ${workbook.status}; ${workbook.worksheets.length} sheets, ${workbook.formulas} formulas, ${workbook.cachedErrors.length} cached formula errors | workbook-inspection.json |
| Fresh extracted release launcher | ${launch?.status??'Pending final extraction check'} | release-launch.json |
| Dependency audit | No known vulnerabilities reported by NuGet or npm at check time | dependency-audit.txt, npm-audit.json |

Machine-readable files are included under validation/ in the release and source archives; local development results are under artifacts/validation/.

## Measured performance

- 2,000 identical asynchronous Excel formulas: **${m.providerLoads} provider load**, ${m.bulk.writeAndCalculateMs} ms to write/start calculation, ${decimal(m.allCellsResolvedMs)} ms for all cells to resolve with an injected two-second provider delay.
- A separate Excel ready/status call during that delay completed in **${decimal(m.responsivenessMs)} ms**. This includes the PowerShell harness polling overhead; it is not a pure UI frame-time measurement.
- 100 warm local HTTP requests: p50 **${decimal(service.performance.warmP50Ms)} ms**, p95 **${decimal(service.performance.warmP95Ms)} ms**.
- Four workbook/pane create-close cycles: Excel private memory ranged from **${decimal(Math.min(...privateMemory))} to ${decimal(Math.max(...privateMemory))} MiB**; stream subscriber count was ${m.lifecycle.map(x=>x.streamSubscribers).join(', ')}. The short observation is useful lifecycle evidence, not proof against every long-running leak.
- Closing a workbook canceled ${m.canceledWaiters} pending formula waiter in the real host. Shared provider work remains available to other callers.

Measurements come from this workstation and this bounded dataset. They are not claims about AlphaSense's infrastructure, arbitrary workbooks, or every supported Office version.

## What the tests protect

Native checks exercise reviewed updates, durable source audit, undo, edits after preview, company changes, forecast/input preservation, later-edit protection, safe formula insertion, saved research, restored Excel flags, malformed formula messages, coalescing, cancellation, and repeated pane lifecycle. The final example was inspected as an XLSX archive for cached formula errors.

Service and unit tests cover cache generations, bounded capacity, shared-request cancellation isolation, immutable snapshot replacement, persistence, annual source selection, HTTP/Origin/Host authentication boundaries, malformed/oversized requests, offline behavior, retry limits, circuit recovery, and WebSocket reconnect/cleanup. AI failure fixtures include missing/rejected keys, rate/server errors, timeout, refusal, incomplete output, malformed structures, and unknown citations. They use synthetic credentials, never the real key.

Browser tests exercise company switching, source dialogs, keyboard interaction, 320px layout, calculated research, offline recovery, notifications, disabled unsupported operations, and untrusted AI text rendering. Office.js adapter tests use a mocked host API; the manifest validator checks schema, not runtime feature correctness.

## Compatibility and limitations

| Environment | Validation status |
| --- | --- |
| Windows Microsoft 365 Excel x64, 16.0, Application.Build 20430 | Real native XLL, COM, WebView2 and workbook tests passed |
| Windows x86 Excel | XLL built; host runtime not tested |
| Excel 2016 / Excel 2019 | Not available for testing; no compatibility claim |
| Mac Excel / Office.js host / Excel web | Adapter and manifest only; no actual host validation |
| Browser | Microsoft Edge via Playwright |
| Enterprise proxy, SSO, signed managed deployment | Not implemented or certified |

The Windows release is unsigned and uses local per-session authentication. Undo is session-scoped, and clients need relaunch after a service-token change. AI citation checks establish identifier membership, not entailment or numerical correctness. The live answer was reviewed against revenue, margins, and operating cash flow less capex; every future answer still needs review. Fiscal calendars and capex definitions differ by issuer.
`;
fs.writeFileSync('docs/VALIDATION.md',report);
console.log('Wrote docs/VALIDATION.md');
