# Delivery checkpoint evidence

The current local Scribble.dll SHA-256 is `66ab464d8cc8c13624be148cf0dd18c5c4377a9665ef6c93aa9ab5d35adb9be0`.

`offline-candidate.json`, `native-candidate.json`, `chartless-acceptance.json`, `production-route.json`, `excel-binding.json`, `analysis-component.json`, `xa01-route.json`, `xa01-failed-workbook.json`, `xa01-malformed-workbook.json`, `pp01-grading.json`, and `visual-review.json` are the current checkpoint. The full local PP01 capture and XA01 test artifacts remain under the ignored `tests/benchmarks/generated/delivery-native/` directory; their exact paths and hashes are in `DELIVERY_STATUS.md`. The new XA01 harness mode changed GuardrailTests.exe without changing Scribble.dll; the earlier PP01 native receipt remains bound to its recorded harness hash and the same production DLL.

`phase4-references.json` and `phase4-defects.json` are historical prior-binary reference-generator receipts. They are retained for traceability and are **not** part of the current-binary pass. The corrected D2P2/D2P3 references were reviewed previously; current-binary full reference review remains open.

Fake endpoint and hand-authored native checks establish route and artifact mechanics. They do not establish model judgment, unseen-case generalization, human visual approval, paid-model feasibility, or deployment readiness.
The XA01 harness reports forced cleanup of its own captured Excel PID after the test. That cleanup protected the pre-existing user Excel session but is not a clean process-exit qualification.
