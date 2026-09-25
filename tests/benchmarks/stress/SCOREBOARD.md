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

Remaining OpenRouter balance: $14.47 at 19:35 UTC on 22 Sep 2026, with no paid
runs since. Re-verify in stage A5.

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
