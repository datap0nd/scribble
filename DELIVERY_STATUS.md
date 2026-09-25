# Scribble 2.0 delivery status

Updated 26 September 2026. This is the execution ledger for the plan merged in [PR #41](https://github.com/datap0nd/scribble/pull/41); [DELIVERY.md](DELIVERY.md) remains the ordered plan and exit contract. Public 2.0.91 remains frozen.

The measured local candidate is Scribble.dll SHA-256 `b729d04a4e7c1064c79a620d404650ad0dc36b0dbbdf74c98a2a3823d1b96c8e`, built from this branch. The sealed corpus manifest is `64f73305c8cf0cc2efdc5b87c539b833fd442c82c67c0ba2f62f06e0ff689d2e`. Fresh offline and native receipts match that assembly hash; source PPT01 and WB01 hashes matched their manifest entries before native testing.

| Gate | Measured result | Boundary |
| --- | --- | --- |
| D1 offline | Release build and complete guardrail suite passed | Local binary; renew on future source changes |
| D1 PowerPoint | `chartless-v1` apply, preservation, rollback, concurrency and structure passed with no PowerPoint exit | `all_operations_passed=false`; chart edits remain outside this scope |
| D1 Excel | Disposable native target binding, formula handling and interrupted-write recovery passed; existing Excel PID 43664 was preserved | This component receipt does not certify all Excel tasks |
| D2 PP01 positive route | Public tools, real HTTP fake endpoint, six slide inspections, owned unsaved draft and terminal task receipt passed in 11 requests; source deck and workbook preserved | Fake responses prove plumbing, not model judgment; negative paths remain open |
| D2 PP01 artifact | Independent native grader accepted the candidate: hard presentation, chart and artifact checks passed | A soft non-solid chart-fill measurement still requires visual review |
| D2 XA01 components | Native typed workbook capture, four-slide output, handoff and recovery checks passed without Office exit | The complete public XA01 request-to-terminal route has not run |
| D3 visuals | Agent reviewed the six-slide PP01 render and the corrected D2P2/D2P3 references at full size; essential legends are legible | No human approval or model-review calibration; full reference approval remains open |
| D4-D9 | Not qualified | Paid model feasibility, unseen generalization, local runtime, broader suite and deployment remain open |

The exact local PP01 native attempt is `tests/benchmarks/generated/delivery-native/20260925T230120388Z-664432b4d6e54f3bbac5a2d3b3b66fd1/`. Its `candidate.pptx` SHA-256 is `a34e6bba43b7d48c15fbc15805f7af932e449c1c25ed77484aa8e905e106a513`; its PDF SHA-256 is `e78fae615c3a27ff9cc1900f2cb46ced3d4e10ecbde2c4ccb592888f6cefe153`. The native capture is test-owned and was saved only for grading. The production draft remained unsaved. Reproducible commands are in [DELIVERY.md](DELIVERY.md#9-reproducible-execution-entry-points). Compact receipts and the agent visual review are in [delivery evidence](tests/benchmarks/stress/evidence/delivery-2026-09-26/).

The first unresolved dependency is D2: run XA01 through the actual request factory, public Excel/PowerPoint hosts, fake HTTP transport, journal and terminal completion; then exercise the specified rejection, interruption, retry, exhaustion and cancellation paths. After that, finish D3 calibration on current visual outputs before any paid D4 pilot. Each code change invalidates the affected candidate receipts and requires a new binary-bound run.
