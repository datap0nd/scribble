# Delivery checkpoint evidence

The current local Scribble.dll SHA-256 is `603f43c4d042d6a6d725bfc5e190d69cf6d41cb6f0c5dd9c29971e968e57a537`; GuardrailTests.exe SHA-256 is `ee32d49352326372a2f7f7a7013496a2e0c5e0747f390a5778b9b0cd72c11db1`.

`offline-candidate.json`, `native-candidate.json`, `chartless-acceptance.json`, `production-route.json`, `excel-binding.json`, `analysis-component.json`, `xa01-route.json`, `xa01-failed-workbook.json`, `xa01-malformed-workbook.json`, `xa01-cancelled.json`, `xa01-transport-retry.json`, `xa01-transport-exhausted.json`, `xa01-rejected-review.json`, `xa01-restart-reconcile.json`, `pp01-grading.json`, and `visual-review.json` are the current checkpoint. The full local PP01 capture and XA01 test artifacts remain under the ignored `tests/benchmarks/generated/delivery-native/` directory; their exact paths and hashes are in `DELIVERY_STATUS.md`. The production DLL changed to repair review-time journal rebasing and resumed-call authorization; all listed current-route receipts were renewed against it.

`phase4-references.json` and `phase4-defects.json` are historical prior-binary reference-generator receipts. They are retained for traceability and are **not** part of the current-binary pass. The corrected D2P2/D2P3 references were reviewed previously; current-binary full reference review remains open.

Fake endpoint and hand-authored native checks establish route and artifact mechanics. They do not establish model judgment, unseen-case generalization, human visual approval, paid-model feasibility, or deployment readiness.

The cancellation receipt covers a stop after the source read and before any draft write. It does not qualify cancellation during an uncertain native mutation or restart recovery.

The transport receipt covers one recovered HTTP 503 after the workbook draft, with a byte-identical inference retry and no duplicate native write. It does not qualify nonrecoverable transport or exhausted allowance.

The exhausted-transport receipt covers two HTTP 503 responses after the workbook draft. The identical retry failed, the task paused, the workbook write remained verified, and no deck or review call followed. It does not qualify other transport failures, exhausted task allowance, or restart recovery.

The rejected-review receipt covers two valid blocker findings after native draft creation. It confirms an uncertain deck write and no terminal completion. After encrypted checkpoint reload and host rebind, a changed deck payload was refused before mutation while source values and slide IDs remained fixed. The separate restart receipt shows the original exact payload reconciled the same four slides and all writes, obtained one fresh fake review, and completed. It does not qualify a cold Office process restart, user-edited destination, or real model judgment. An earlier disposable attempt on the prior binary stopped before review on a native chart out-of-memory error; repeatable resource stability remains open.

The current PP01 PDF rendered byte-identically to the previously agent-reviewed PDF on all six pages at 110 and 120 dpi. `visual-review.json` corrects three transcription errors in the old page-hash ledger against both files. This carried review is not human approval.

The XA01 harness reports forced cleanup of its own captured Excel PID after the test. That cleanup protected the pre-existing user Excel session but is not a clean process-exit qualification.
