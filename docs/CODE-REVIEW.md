# Code review and regression testing — 1.0.1

This pass reviewed the calculation cache, workbook transaction boundaries, identity and lifecycle, UI dispatch, event streaming, evidence export, test harness, and release reporting. It includes behavioral fixes and readable, enforced source formatting.

## Corrections

| Finding | Change | Regression evidence |
| --- | --- | --- |
| A fast cache load could complete between another caller's cache miss and in-flight reservation | Reserve shared work in the same critical section as the cache lookup; register cleanup once per shared load | 100 rounds of 64 concurrent immediate-result callers; existing failure, eviction, invalidation and cancellation tests |
| A historical cell already equal to its source was absent from preview conflict checks | Capture unchanged historical values as dependencies while continuing to preserve analyst inputs | Unit test failed before the fix; real Excel test now rejects that edit without partial writes |
| A resolver could return a valid fact belonging to another issuer | Check returned company, metric, and period against the requested key | Unit test failed before the fix |
| Workbook copies shared their saved document GUID and could replace or remove one another's transaction state | Assign an independent in-memory session identity to each open workbook object; release it on close/unload | Unit test plus real Excel copy/apply/close/original-apply/undo workflow |
| Parallel event publishers could enqueue sequence numbers out of order | Serialize publication and enqueue each subscription's handshake before subsequent events | Concurrent-publisher test failed before the fix; bounded-queue and handshake tests pass |
| An abruptly failed WebSocket receive could leave the send loop waiting | Cancel the subscription loop in the receive task's finally block; reject unsupported protocols | HTTP/WebSocket disconnect/reconnect tests and deterministic subscriber cleanup |
| Cancellation could report a canceled result after a synchronous workbook write had already begun | Separate queued/canceled/executing states; cancellation prevents queued work, while started work reports its actual outcome | Tests for cancellation before dispatch, during execution, and dispatcher failure |
| Late pane callbacks consulted a disposed service singleton | Capture the allowed service origin in the pane; reject messages during disposal; isolate WebView profiles per release root | Twelve real workbook/pane cycles with subscription and memory observations |
| Native formula insertion omitted the read-only check | Check the selected cell's owning workbook before writing | Real read-only Excel workbook test |
| Fresh launch saved as soon as one formula resolved, leaving 25 pending formula errors in the file | Wait for calculation completion and zero formula errors across all worksheets before saving | Fresh extracted-release launch and inspection of every worksheet in the saved XLSX |
| Research export checked claim IDs but not completeness/validity of its embedded source records | Validate source presence, uniqueness, metadata and citation coverage before creating a sheet | Missing, duplicate, foreign-URL and wrong-unit evidence tests |
| Validation documentation contained hard-coded test counts and could inherit an old release launch result | Read counts from test artifacts; require passing suite evidence; match launch proof to the archive version | Generated validation report and fresh 1.0.1 extraction check |

The before-fix run reproduced three deterministic failures in the newly added tests. Its TRX and the successful full-suite log are included under `validation/quality-review` in the packages.

## Maintainability

C# formatting is defined in `.editorconfig` and enforced with `dotnet format whitespace --verify-no-changes`. JavaScript, HTML, CSS, and JavaScript tests use a pinned Prettier dependency and a checked-in configuration. The previously compressed stylesheet is now readable source. Both checks run in the standard test command. Product version reporting derives from the built Core assembly rather than separate handwritten constants.

The UI-dispatch cancellation policy and workbook-session registry are independent of COM and have direct unit tests. The native adapter retains responsibility for main-thread COM access. Test automation resolves the service endpoint dynamically, so the review service and Excel instance can run separately from an already-open demonstration.

## Final coverage

- 50 .NET unit/contract tests.
- 15 service integration checks, including 2,000 mixed requests across all 81 fact keys with at most four provider slots.
- 19 real Excel/WebView2 checks, including asynchronous array spilling, protected audit sheets, copied workbooks, read-only insertion, and twelve lifecycle cycles.
- Eight browser tests and nine Office.js adapter tests.
- Independent reconciliation of 81 SEC facts, dependency audits, manifest validation, source formatting, XLSX cached-error inspection, package credential scanning, and fresh extracted-release launch.

These results cover the installed Windows Microsoft 365 Excel x64 host. They do not establish Mac, older Excel, enterprise SSO/proxy, signed deployment, or indefinite leak-free operation. See [VALIDATION.md](VALIDATION.md) for measurements and retained provider-evidence timestamps.
