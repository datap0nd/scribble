# Delivery checkpoint evidence

The current local Scribble.dll SHA-256 is `66ab464d8cc8c13624be148cf0dd18c5c4377a9665ef6c93aa9ab5d35adb9be0`.

`offline-candidate.json`, `native-candidate.json`, `chartless-acceptance.json`, `production-route.json`, `excel-binding.json`, `analysis-component.json`, `xa01-route.json`, `xa01-failed-workbook.json`, `xa01-malformed-workbook.json`, `xa01-cancelled.json`, `xa01-transport-retry.json`, `xa01-transport-exhausted.json`, `xa01-rejected-review.json`, `pp01-grading.json`, and `visual-review.json` are the current checkpoint. The full local PP01 capture and XA01 test artifacts remain under the ignored `tests/benchmarks/generated/delivery-native/` directory; their exact paths and hashes are in `DELIVERY_STATUS.md`. The XA01 harness modes changed GuardrailTests.exe without changing Scribble.dll; the earlier PP01 and XA01 receipts remain bound to their recorded harness hashes and the same production DLL.

`phase4-references.json` and `phase4-defects.json` are historical prior-binary reference-generator receipts. They are retained for traceability and are **not** part of the current-binary pass. The corrected D2P2/D2P3 references were reviewed previously; current-binary full reference review remains open.

Fake endpoint and hand-authored native checks establish route and artifact mechanics. They do not establish model judgment, unseen-case generalization, human visual approval, paid-model feasibility, or deployment readiness.

The cancellation receipt covers a stop after the source read and before any draft write. It does not qualify cancellation during an uncertain native mutation or restart recovery.

The transport receipt covers one recovered HTTP 503 after the workbook draft, with a byte-identical inference retry and no duplicate native write. It does not qualify nonrecoverable transport or exhausted allowance.

The exhausted-transport receipt covers two HTTP 503 responses after the workbook draft. The identical retry failed, the task paused, the workbook write remained verified, and no deck or review call followed. It does not qualify other transport failures, exhausted task allowance, or restart recovery.

The rejected-review receipt covers two valid blocker findings after native draft creation. It confirms an uncertain deck write and no terminal completion, but does not qualify retry or recovery. Its first local attempt stopped earlier on a native chart out-of-memory error; only the clean-session retry reached review.

The XA01 harness reports forced cleanup of its own captured Excel PID after the test. That cleanup protected the pre-existing user Excel session but is not a clean process-exit qualification.
