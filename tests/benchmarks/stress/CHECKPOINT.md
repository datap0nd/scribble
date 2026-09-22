# Native acceptance checkpoint — 2026-09-22

Acceptance requires **three consecutive unassisted live passes each** for XA01 (Excel to four-slide deck) and PP01 (repair existing six-slide deck). Each pass must verify factual accuracy, preserved source work, and rendered slide quality. CI and draft creation alone do not count.

| Case | Last live result | Fix / next hypothesis | Consecutive passes |
| --- | --- | --- | --- |
| XA01 | 2.0.385 live run completed unassisted. Native evaluator passed final XLSX/PPTX, source preservation, exact formula-derived May/June values, four-slide structure, chart data, and text fit. Manual inspection of all four rendered slides found a readable scorecard, chart/table comparisons, and structured data-quality cards without clipping or collision. Evidence: `suite-20260922-163550-3fa83055`, `artifacts/qa-385-xa01-run1/`. | Repeat twice without an intervening XA01 failure, with the same source/accuracy/rendered checks. The evaluator records `needs_native_visual_review`, so retain the manual visual assessment alongside native gates. | 1/3 |
| PP01 | 2.0.395 rendered a readable six-slide deck, removed the slide-4 KPI duplication, and passed every deterministic native gate, including exact chart data, text fit, and byte-for-byte preservation of both source files. It does **not** count: the trace never reached a terminal case event and was paused after 23 minutes. The main stall was repeated full-deck preflight retries caused by a probabilistic reviewer claiming the cover contained KPI callouts even though the submitted JSON and rendered cover had none; later warning-only casing/wording findings also restarted the deck. Evidence: `suite-20260922-184018-0e1541e4`, `artifacts/qa-395-pp01-stopped/`. | Refute cover-metric findings only when the parsed cover demonstrably has no metric structures or callouts and the user did not request them; keep genuine mixed blockers closed. Treat explicit warning-only review findings as advisory. Next hypothesis: PP01 reaches a terminal receipt without repeated absent-cover-callout retries, while the real May/June direction blocker remains enforced. | 0/3 |

At 12:23 UTC, OpenRouter confirmed the user's earlier **$10 top-up** and the active no-reset inference key's total cap was increased from **$20 to $30**. After the user authorized another $10, the active key was independently verified and its no-reset cap was increased from **$30 to $40** at 19:35 UTC. It then reported **$14.47 remaining** and **$25.53 used**. The harness checks before each model request and retains a $0.05 reserve. CI passed for installed 2.0.395; a new candidate is required because the former harness authorization stopped at $30. The live acceptance bar remains XA01 1/3 and PP01 0/3. Do not infer passes from CI or a partial native deck. Preserve Office originals and uncertain drafts for inspection; never replay a changed write without receipt reconciliation.

Candidate 2.0.398 passed CI and installed successfully. Its PP01 attempt (`suite-20260922-194429-80f7e1bb`) stopped during `SnapshotExternalKit` before any case, model request, or document write; it produced no timeline or case result and does not count as a pass or failure. No further paid PP01 reruns are authorized until reliability-plan phases 0–3 pass offline. Preserve the verified **$14.47** balance for the phase-3 pilot and smoke runs.

## Reliability phase and experiment checkpoint format

For every new experiment, record its hypothesis, code commit, configuration
fingerprint, case IDs, expected change, actual deterministic/native/visual
status, defect class, total model requests, prompt/completion tokens, reported
cost, elapsed time, and next decision. A missing or pending gate stays pending;
an installer build or model-authored completion message is not an acceptance
result. Record provider per response if OpenRouter fallbacks remain enabled.

| Phase | Commit / evidence | Offline gate | Native gate | Visual gate | Paid model use | Next decision |
| --- | --- | --- | --- | --- | --- | --- |
| 0: baseline and mechanisms | `ea5c296`, `tests/benchmarks/stress/PHASE0_BASELINE.md`; hosted config SHA-256 `ee2400f4288b6f51754e36aaf801cfa4527bba02b96be60e0c9e00d311e4e941` | Full PR #22 CI run `35790300096` passed; all seven mechanisms have offline coverage | No Office write in phase 0; local 4K vision and 2K schema probes completed without Office mutation | No new candidate output | $0 | Phase 0 passed for the hosted pilot. Report this gate, then start phase 1. Local 27B remains uncertified for Office due RAM headroom. |
