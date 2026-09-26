# Delivery checkpoint evidence

The current local Scribble.dll SHA-256 is `ea845efeffa2e0e3d5ad0cf44818d9de2965be44dab8e2cde0b191a7aab6cd2c`.

`offline-candidate.json`, `native-candidate.json`, `chartless-acceptance.json`, `production-route.json`, `excel-binding.json`, `analysis-component.json`, `xa01-route.json`, `pp01-grading.json`, and `visual-review.json` are the current checkpoint. The full local PP01 capture and XA01 test artifacts remain under the ignored `tests/benchmarks/generated/delivery-native/` directory; their exact paths and hashes are in `DELIVERY_STATUS.md`.

`phase4-references.json` and `phase4-defects.json` are historical prior-binary reference-generator receipts. They are retained for traceability and are **not** part of the current-binary pass. The corrected D2P2/D2P3 references were reviewed previously; current-binary full reference review remains open.

Fake endpoint and hand-authored native checks establish route and artifact mechanics. They do not establish model judgment, unseen-case generalization, human visual approval, paid-model feasibility, or deployment readiness.
