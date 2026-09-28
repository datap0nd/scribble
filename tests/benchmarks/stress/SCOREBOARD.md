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
| S1 XA01–XA10 batch | ≥ 9/10 | 0/9 eligible on 2.0.706.0; XA02–XA10 ran as one batch |
| S1 XA01 consecutive | 3 | 0 eligible on the new route; the 2.0.706.0 run used 19 requests (cap 12) |
| S2 PP01–PP10 batch | ≥ 8/10 | not run |
| S2 PP01 consecutive | 3 | 0 |

Remaining OpenRouter balance: $14.4736 at 21:38 UTC on 26 Sep 2026 (existing
`TestLabStressBudget.CheckAsync` against `/api/v1/key`; $25.5264 used of the
$40 no-reset key). After seven real XA01 runs: $12.9402 at 06:22 UTC on 27 Sep. Recheck
before the next paid batch.
After the 28 Sep B4 batch: $11.2796 remaining at 14:01 UTC ($28.7204 used);
recheck before another paid batch.

## Installed builds

| Date | Version | Source | Installer SHA-256 | Installed Scribble.dll SHA-256 | Verification |
| --- | --- | --- | --- | --- | --- |
| 2026-09-26 | 2.0.694.0 | PR #44 head `99837ca`, CI merge `65ea83a`, run `36273774551` | `96bdc751a347d3d4bd1ac3c2d7b8ccb288351e9b29c7fe657fecdb7c33115811` | `2efd65023777c9c9abc9cf7395653ae0fce245cac59d299d94434583ecc33bc3` | Pilot installer completed; renderer payload present; version and hashes read back from installed files. |
| 2026-09-26 | 2.0.695.0 | PR #44 head `dc98db9`, CI run `36276009192` | `f51c1c22d3b234c405d694fba510e2ddb3303a2bb2ed3b207e9919f10bfe243b` | `3173792c2b296b55be3c58650f91cc2d73a922be47ea87125df68be1beec3ead` | Pilot installer exited 0; installed version, DLL and browser host hashes verified; no Office processes existed before or after install. |
| 2026-09-26 | 2.0.696.0 | PR #44 head `63ca17b`, CI run `36277334428` | `3cd4694c7f8f40957a869925bccda4201cbc3fda4c32c7d6f29058b3900c5a6d` | `8a1b82135ccdfbc73efef39f8c2a8adb8f0960b8846d64de7b887793d40d3a9b` | Pilot installer exited 0; installed version, DLL and renderer payload verified; no Office processes existed before or after install. |
| 2026-09-26 | 2.0.697.0 | PR #44 head `b590816`, CI run `36278264802` | `65a6c650f9b0cc52e01a2a10b44e356c047b54993d9378c1b7d38d67954898e8` | `3c034227b7e4f5c7cf55274cbdad37d80a3c8c94307953ff22d5e210bcd993c3` | Pilot installer exited 0; installed version, DLL and renderer payload verified; no Office processes existed before or after install. |
| 2026-09-26 | 2.0.698.0 | PR #44 head `1ce0670`, CI run `36279656273` | `8474259bbb164f0ba52c5441ac13f82b316e3c762ec058834888e934141e64fe` | `1c2508e172b0a5cac3debc375ad74d0437fbe5426a51a8ce2932280fba4ea960` | Pilot installer exited 0; installed version, DLL and renderer payload verified; no Office processes existed before or after install. |
| 2026-09-27 | 2.0.699.0 | PR #44 head `c632ec9`, CI run `36281114429` | `4243c202bc6fbeb08b08f0bf7caa5c0ca8c9065695645881b18caddf9af7ecf4` | `f6eb63e8075595a402ade24674fbee9dc3e0150d9edadadb19aa175b998d5684` | Pilot installer exited 0; installed version, DLL and renderer payload verified; no Office processes existed before or after install. |
| 2026-09-27 | 2.0.701.0 | PR #44 head `e1e1bb5`, CI run `36282424477` | `66ce36603a0b081e71534c9f11dbc7d21de34a2c64815056b2d533729fa59850` | `6ea8f9abfd1247c1cbb38813839c9a3729d55737fed7a2762809f1eaa30b1459` | Pilot installer exited 0; installed version, DLL, browser host and renderer payload verified; no Office processes existed before or after install. |
| 2026-09-28 | 2.0.713.0 | PR #45 head `c97a175`, CI run `36432733798` | `000f74e34c04ec7f95d3551e315545601eb98eecf319c033be6983ed5b573e06` | `20533a0b5089b379570a81278353ad439a586d128033bb423e6b59ef4cefa7c8` | Green CI; pilot installer exited 0; installed DLL and browser host hashes verified. Preexisting owner Excel PID 33464 and start time unchanged; no other Office process remained after install. |

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
| 2026-09-26 | 2.0.695.0 (`dc98db9`, CI `36276009192`) | XA01 | new | FAIL: typed grouped binding rejected formula-valued source metrics, so no deck or terminal event | capture/binding | 14 | $0.11 | 2.7 | `suite-20260926-223645-e2bf561f` / `ec57603e05c74e7bb46364bf735b5cd5` |
| 2026-09-26 | 2.0.696.0 (`63ca17b`, CI `36277334428`) | XA01 | new | FAIL: grouped source facts succeeded; draft supplied the first analysis ID after a second grouped read extended it, so no outputs or terminal event | plan/preflight | 9 | $0.06 | 1.5 | `suite-20260926-225713-99570bc1` / `2fc49c125fa54015b7b5cb21b8277ea0` |
| 2026-09-26 | 2.0.697.0 (`b590816`, CI `36278264802`) | XA01 | new | FAIL: typed workbook draft succeeded; deck plan repeatedly failed numeric/text authority and expanded citations, with no deck or terminal event | planning/schema | 32 | $0.46 | 8.8 | `suite-20260926-231455-2049ec90` / `6b751d2c36ba4e02a7f3e6dc776c6ca6` |
| 2026-09-26 | 2.0.698.0 (`1ce0670`, CI `36279656273`) | XA01 | new | FAIL: typed workbook draft succeeded; model could not produce a valid full deck plan within repeated schema attempts, no deck or terminal event | planning/schema | 33 | $0.44 | 8.8 | `suite-20260926-234245-02f61298` / `41012b165afd41f085c364abb3a08974` |
| 2026-09-27 | 2.0.699.0 (`c632ec9`, CI `36281114429`) | XA01 | new | FAIL: terminal event and native xlsx/pptx; verified May/June formulas landed in F/G rather than B/C, and chart included a secondary series | planning/schema | 10 | $0.09 | 2.7 | `suite-20260927-001224-0954e081` / `04178e05b5114c618e3affbc34bea4d8` |
| 2026-09-27 | 2.0.701.0 (`e1e1bb5`, CI `36282424477`) | XA01 | new | FAIL: all deterministic gates passed and terminal native outputs produced, but native visual review remains; hard-coded request selection makes this build ineligible for S1 | review | 10 | $0.04 | 2.1 | `suite-20260927-062019-291570f4` / `0d4efd81b59c446792a41970ad353b6e` |
| 2026-09-28 | 2.0.704.0 (`62f43d4`, CI `36416255638`) | XA01 | new | FAIL: typed draft formulas verified, but generic `write_cells` changed bound Ledger B3:C3; later deck attempts lacked the original full-period source and no pptx or terminal event resulted | planning/schema (new draft period arguments rejected before the source edit) | 29 | $0.22 | 11.2 | `suite-20260928-114350-030bd01a` / `efa979e2855442ababfc2ce8733ffdee` |
| 2026-09-28 | 2.0.705.0 (`d235ec4`, CI `36419124623`) | XA01 | new | FAIL: terminal native XLSX/PPTX and sources unchanged; extra verified metrics displaced requested Revenue/Cost from B4:C5 and made Units the chart series | planning/schema | 12 | $0.04 | 2.8 | `suite-20260928-121309-ac70a964` / `267247bf77f740e685a39c174f9d1bc8` |
| 2026-09-28 | 2.0.706.0 (`ede1b26`, CI `36421402714`) | XA01 | new | PASS: terminal and unassisted; all native hard checks passed, sources unchanged, B4:C5 and single Revenue chart correct; four full-size slides passed agent visual review (owner review remains for D2) | — | 19 | $0.04 | 6.0 | `suite-20260928-123347-9aa2eec7` / `effa79f13ae447d4b0fc680c46e8a2b9` |
| 2026-09-28 | 2.0.706.0 (`ede1b26`, CI `36421402714`) | XA02 | new | FAIL: formula-valued source metrics did not bind; no terminal deck; exceeds 12-request cap | capture/binding | 42 | $0.300 | 10.8 | `suite-20260928-125813-5ab7f507` / `304158199abd4d768f12ec699a9e500e` |
| 2026-09-28 | 2.0.706.0 (`ede1b26`, CI `36421402714`) | XA03 | new | FAIL: period and dimensional reads did not retain one complete source view; later review/recovery failed; exceeds cap | planning/schema | 28 | $0.103 | 5.0 | `suite-20260928-125813-5ab7f507` / `75d6beee42a0401a96b665d9158d937f` |
| 2026-09-28 | 2.0.706.0 (`ede1b26`, CI `36421402714`) | XA04 | new | FAIL: blank source value rejected grouped binding; no terminal deck; exceeds cap | capture/binding | 53 | $0.306 | 12.7 | `suite-20260928-125813-5ab7f507` / `9e9d1eb6a797453aa1731a774925b3d4` |
| 2026-09-28 | 2.0.706.0 (`ede1b26`, CI `36421402714`) | XA05 | new | FAIL: period and dimensional reads did not retain one complete source view; later review/recovery failed; exceeds cap | planning/schema | 31 | $0.173 | 7.0 | `suite-20260928-125813-5ab7f507` / `df883ef56a524e1d934fab81c99cca19` |
| 2026-09-28 | 2.0.706.0 (`ede1b26`, CI `36421402714`) | XA06 | new | FAIL: chart used Planned Hours instead of requested Actual Hours; also exceeded cap | planning/schema | 13 | $0.036 | 2.2 | `suite-20260928-125813-5ab7f507` / `f18298fce0f647c399198b5a11f83b26` |
| 2026-09-28 | 2.0.706.0 (`ede1b26`, CI `36421402714`) | XA07 | new | FAIL: native checks and agent visual review passed, but 16 requests exceed cap; owner visual review remains D2 | completion | 16 | $0.038 | 2.7 | `suite-20260928-125813-5ab7f507` / `9c33b8374c584c819384204df5250b92` |
| 2026-09-28 | 2.0.706.0 (`ede1b26`, CI `36421402714`) | XA08 | new | FAIL: blank and guarded formula source values did not bind; no terminal deck; exceeds cap | capture/binding | 26 | $0.194 | 6.9 | `suite-20260928-125813-5ab7f507` / `61dca74daf224a8783a935da0ee65c9d` |
| 2026-09-28 | 2.0.706.0 (`ede1b26`, CI `36421402714`) | XA09 | new | FAIL: native checks and agent visual review passed, but 15 requests exceed cap; owner visual review remains D2 | completion | 15 | $0.040 | 2.5 | `suite-20260928-125813-5ab7f507` / `159c3d6de32f408d80d2aed26dfe4e36` |
| 2026-09-28 | 2.0.706.0 (`ede1b26`, CI `36421402714`) | XA10 | new | FAIL: reviewer returned fenced JSON; review parsing and later slide recovery blocked terminal handoff; exceeds cap | review | 23 | $0.152 | 7.3 | `suite-20260928-125813-5ab7f507` / `9be5dd57426f4e1b8c2551ae7df76a00` |
| 2026-09-28 | 2.0.713.0 (`c97a175`, CI `36432733798`) | XA01 | new | FAIL: terminal typed workbook and four native slides; all deterministic hard checks, source preservation and agent visual review passed, but 13 requests exceed the S1 cap of 12; owner visual review remains D2 | completion | 13 | $0.033 | 3.1 | `suite-20260928-141119-ee36d02c` / `d4716076d15e4c81bc7381107379be71` |
