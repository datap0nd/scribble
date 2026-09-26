# Scribble 2.0 delivery status

Updated 26 September 2026. This is the execution ledger for the plan merged in [PR #41](https://github.com/datap0nd/scribble/pull/41); [DELIVERY.md](DELIVERY.md) remains the ordered plan and exit contract. Public 2.0.91 remains frozen.

The measured local candidate is Scribble.dll SHA-256 `ea845efeffa2e0e3d5ad0cf44818d9de2965be44dab8e2cde0b191a7aab6cd2c`, built from this branch. The sealed corpus manifest is `64f73305c8cf0cc2efdc5b87c539b833fd442c82c67c0ba2f62f06e0ff689d2e`. Fresh offline and native receipts match that assembly hash; source PPT01 and WB01 hashes matched their manifest entries before native testing.

| Gate | Measured result | Boundary |
| --- | --- | --- |
| D1 offline | Release build and complete guardrail suite passed | Local binary; renew on future source changes |
| D1 PowerPoint | `chartless-v1` apply, preservation, rollback, concurrency and structure passed with no PowerPoint exit | `all_operations_passed=false`; chart edits remain outside this scope |
| D1 Excel | Disposable native target binding, formula handling and interrupted-write recovery passed; existing Excel PID 43664 was preserved | This component receipt does not certify all Excel tasks |
| D2 PP01 positive route | Public tools, real HTTP fake endpoint, six slide inspections, owned unsaved draft and terminal task receipt passed in 11 requests; source deck and workbook preserved | Fake responses prove plumbing, not model judgment; negative paths remain open |
| D2 PP01 artifact | Independent native grader accepted the candidate: hard presentation, chart and artifact checks passed | A soft non-solid chart-fill measurement still requires visual review |
| D2 XA01 positive route | Public Excel and PowerPoint hosts, real HTTP fake endpoint, typed facts, two owned unsaved drafts, verified writes and terminal task receipt passed in four model requests and five review calls; source sheet preserved | Fake responses prove plumbing, not model judgment; negative paths remain open |
| D2 XA01 components | Native typed workbook capture, four-slide output, handoff and recovery checks passed; a test-owned PowerPoint process remained open after the component pilot and was closed before grading | This component receipt does not qualify process cleanup or the full recovery matrix |
| D3 visuals | Agent reviewed the current six-slide PP01 render at full size; essential chart legend is legible. Earlier D2P2/D2P3 corrected references were also reviewed | No human approval or model-review calibration; current-binary full reference approval remains open |
| D4-D9 | Not qualified | Paid model feasibility, unseen generalization, local runtime, broader suite and deployment remain open |

The exact local PP01 native attempt is `tests/benchmarks/generated/delivery-native/20260926T000434555Z-f4f599f619b84467acf0ef555017c186/`. Its `candidate.pptx` SHA-256 is `fa7a885a89345426f533f98b594424fe4af021e134104c8d3510111abbceef12`; its PDF SHA-256 is `381dd709af3ba07a0fb74970e2f9b4ad04bb9372c1abedeadd2b43d1fce405c9`. The native capture is test-owned and was saved only for grading. The production draft remained unsaved. The exact XA01 full-route report is `tests/benchmarks/generated/delivery-native/xa01-route-final`. Reproducible commands are in [DELIVERY.md](DELIVERY.md#9-reproducible-execution-entry-points). Compact receipts and the agent visual review are in [delivery evidence](tests/benchmarks/stress/evidence/delivery-2026-09-26/).

The first unresolved dependency is D2 negative-path qualification: exercise the specified failed tool, malformed proposal, rejected review, cancellation, transport retry, exhausted allowance and restart through production routes, with precise non-success receipts and no blind replay. After that, finish D3 calibration on current visual outputs before any paid D4 pilot. Each code change invalidates the affected candidate receipts and requires a new binary-bound run.
