# Scribble 2.0 scoreboard

This file is the only progress measure for [DELIVERY.md](../../../DELIVERY.md).
Add one row per real-model or native run, newest last, and never edit or
delete an old row. A run that didn't happen still gets a row with the result
`NOT RUN` and the reason.

Fill every column. Write `—` when a value genuinely doesn't exist; don't
leave a blank.

| Column | Contents |
| --- | --- |
| Date | Run date |
| Build | Installed build number and source commit |
| Case | Corpus case ID |
| Route | `old` (model-authored Samsung workflow) or `new` (typed analysis route) |
| Result | `PASS`, `FAIL` or `NOT RUN`, and why |
| First failing stage | One of: capture/binding, planning/schema, calculation, render, native execution, review, completion or harness |
| Requests | Model requests for the whole task |
| Cost | Provider-reported cost in USD |
| Minutes | Wall-clock minutes to the terminal event |
| Trace | Suite or run ID |

## Summary

| Scenario | Bar (DELIVERY.md §1) | Current build |
| --- | --- | --- |
| S1 XA01–XA10 batch | ≥ 9/10 | not run |
| S1 XA01 consecutive | 3 | 0 on the new route |
| S2 PP01–PP10 batch | ≥ 8/10 | not run |
| S2 PP01 consecutive | 3 | 0 |

Remaining OpenRouter balance: $14.4736 at 21:38 UTC on 26 Sep 2026 (existing
`TestLabStressBudget.CheckAsync` against `/api/v1/key`; $25.5264 used of the
$40 no-reset key). After the real XA01 run: $14.1401 at 22:06 UTC. Recheck
before the next paid batch.

## Installed builds

| Date | Version | Source | Installer SHA-256 | Installed Scribble.dll SHA-256 | Verification |
| --- | --- | --- | --- | --- | --- |
| 2026-09-26 | 2.0.694.0 | PR #44 head `99837ca`, CI merge `65ea83a`, run `36273774551` | `96bdc751a347d3d4bd1ac3c2d7b8ccb288351e9b29c7fe657fecdb7c33115811` | `2efd65023777c9c9abc9cf7395653ae0fce245cac59d299d94434583ecc33bc3` | Pilot installer completed; renderer payload present; version and hashes read back from installed files. |

## Runs

The 22 Sep rows come from `CHECKPOINT.md` and `PHASE0_BASELINE.md`, which
predate this file. Their source commits weren't recorded there.

| Date | Build | Case | Route | Result | First failing stage | Requests | Cost | Minutes | Trace |
| --- | --- | --- | --- | --- | --- | ---: | ---: | ---: | --- |
| 2026-09-22 | 2.0.385 | XA01 | old | PASS | — | 26 | $0.20 | — | `suite-20260922-163550-3fa83055` |
| 2026-09-22 | 2.0.388 | PP01 | old | FAIL | not recorded | 87 | $0.52 | — | `suite-20260922-165042-738f6117` |
| 2026-09-22 | 2.0.391 | PP01 | old | FAIL: trace incomplete | not recorded | 74 | $0.49 | — | `suite-20260922-172330-f0091961` |
| 2026-09-22 | 2.0.395 | PP01 | old | FAIL: no terminal event, paused after 23 min | review (a false cover-KPI finding restarted full-deck preflight) | — | — | > 23 | `suite-20260922-184018-0e1541e4` |
| 2026-09-22 | 2.0.398 | PP01 | old | NOT RUN: stopped in `SnapshotExternalKit` before any request | harness | 0 | $0.00 | — | `suite-20260922-194429-80f7e1bb` |
| 2026-09-26 | local `603f43c4` (PR #43 native candidate) | PP01 | new | PASS: fake endpoint production route; terminal receipt, grader accepted, sources unchanged | — | 11 | $0.00 | — | `delivery-2026-09-26/production-route.json` |
| 2026-09-26 | local `603f43c4` (PR #43 native candidate) | XA01 | new | PASS: fake endpoint production route; terminal receipt, sources unchanged | — | 4 model + 5 review | $0.00 | — | `delivery-2026-09-26/xa01-route.json` |
| 2026-09-26 | 2.0.694.0 (`99837ca`, CI `65ea83a`) | XA01 | old | FAIL: typed route not entered despite persisted setting; native outputs passed oracle but no terminal task event; 34 requests exceeds cap | capture/binding | 34 | $0.33 | 7.3 | `suite-20260926-215717-2b996964` / `7c7a41f6720a4d589284c7a259f89a20` |
