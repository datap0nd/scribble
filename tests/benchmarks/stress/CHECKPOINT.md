# Native acceptance checkpoint — 2026-09-22

Acceptance requires **three consecutive unassisted live passes each** for XA01 (Excel to four-slide deck) and PP01 (repair existing six-slide deck). Each pass must verify factual accuracy, preserved source work, and rendered slide quality. CI and draft creation alone do not count.

| Case | Last live failure | Fix / next hypothesis | Consecutive passes |
| --- | --- | --- | --- |
| XA01 | 2.0.352, 10:32 UTC: before any write, Qwen repeatedly passed `rows`/`columns` to read-only `list_worksheets` despite reasoning that it needs `{}`; the task stopped after repeated invalid arguments. Earlier 2.0.346 run left an invalid-formula draft uncertain. | Ignore only those inert inventory hints at the read-only validation boundary. A known post-write formula rejection now keeps its incomplete marked sheet but permits a new numbered draft; unknown write outcomes remain blocked. Next live run tests source reads, the Excel draft, and four rendered slides. | 0/3 |
| PP01 | 2.0.352: one slide was appended after six originals; reviewer falsely rejected its native page number, repair returned unchanged content, and a changed follow-up write hit the recovery boundary. | The host now accepts a page-number-only objection only after verifying the native footer. A same-tool changed continuation now substitutes the exact journaled arguments only when receipts prove safe resume; unrelated writes stay blocked. Next live run tests advancement beyond the cover without weakening other review gates. | 0/3 |

Last recorded OpenRouter key usage: **$18.672035519 / $20**, leaving **$1.327964481** (no reset). The stress checkpoint has been raised to the existing $20 hard cap, but request admission still stops at $19.75, retaining a $0.25 reserve. Do not infer a live pass from CI or a partial native deck. Preserve the existing Office originals and any uncertain draft for inspection; never replay a changed write without receipt reconciliation.
