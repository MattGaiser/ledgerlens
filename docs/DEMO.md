# Three-minute interview demo

Open the release with **Start LedgerLens.cmd**. Before the call, confirm Health shows the service connected. Keep the calculated-analysis option available if the AI provider is slow. Start with a fresh model so the FY2025 import column is empty.

**0:00–0:30 — The analyst's problem.** “An analyst should be able to get a reported value, verify its provenance, and update a model without losing assumptions. I built this independent demonstration around that workflow using SEC evidence and OpenAI.” Show the Dashboard comparison and chart. Explain USD millions and different fiscal calendars.

**0:30–1:00 — Follow a number.** In the pane, click operating income. Show its XBRL concept, accession, period, filing date, raw value, and SEC link. Select an empty cell and insert the formula. Explain that HTTP work occurs asynchronously outside Excel's COM thread.

**1:00–1:45 — Protect the model.** Switch to Model. Preview updates, point out the nine historical values, then apply. Change the blue revenue-growth assumption: the forecast updates. Open Sources and show the appended import audit. Undo the update: historical values revert while the analyst's assumption remains. For a deeper demonstration, preview again, manually edit E10, and attempt apply; the conflict blocks the entire transaction.

**1:45–2:15 — AI with evidence.** Ask “Compare FY2024 and FY2025 operating margins and cash generation.” Show the provider label, claims, and clickable source citations. Save the answer to a worksheet. Explain that citation membership is validated, but numeric reasoning still needs review. Calculated analysis is explicitly labeled when AI is not used.

**2:15–2:45 — Reliability.** Open Health, select Offline, and show saved evidence with its freshness label. Return Online. Replay a filing notification and show the WebSocket event. Explain that replay changes no reported values. Mention the measured 2,000-formula coalescing result from the validation report.

**2:45–3:00 — Engineering judgment.** “The native adapter owns Excel lifecycle and calculation concerns. A separate service owns providers and resilience. The shared UI also has an Office.js adapter, but I have only validated the native Windows host here.” Invite questions about conflicts, cancellation, deployment, and the migration seam for AlphaSense APIs.

## Prepare to explain

- Why automatic retries apply to read operations and not ambiguous workbook mutations.
- Why canceling one formula waiter must not cancel a provider request shared by other cells.
- Why async HTTP cannot touch Excel COM objects on a worker thread.
- How the import's company, metric map, and period become conflict dependencies.
- Why a last-applied transaction is session state while its evidence audit belongs in the workbook.
- Why Excel-DNA was selected for UDFs and the C# object model, and how a VSTO shell could reuse Core and the service.
- What changes for production: SSO and entitlements, signed deployment, supported Office test matrix, proxy tests, observability, and source-specific licensing.

Use your own words. Be explicit about AI assistance and the parts you reviewed, modified, and can defend. This artifact is a conversation starter and work sample; it does not establish years of production ownership.
