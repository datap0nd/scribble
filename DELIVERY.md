# Scribble 2.0 — delivery plan (revision 2)

Revised 26 September 2026. This replaces revision 1 (`6e822f6`, D0–D10,
merged with PR #41 as `ce85b4f`) as the Scribble 2.0 delivery plan.
Revision 1 and the reliability/phase documents stay in history as design
rationale. They do not set scope, gates or order any more.
Only the owner changes this plan. The executing agent updates **§5 Status**
and `tests/benchmarks/stress/SCOREBOARD.md`, and nothing else.

## 0. Why the plan changed

| Observation (verified 25 Sep) | Consequence |
| --- | --- |
| The live scoreboard has not moved since 22 Sep: XA01 1/3 (old route, 2.0.385), PP01 0/3. The new architecture has **never run with a real model**. | 300 commits since the last paid attempt, 59 of them only record, archive or trace evidence. Nothing in that work was measured against the real task. |
| The next critical-path step (a native Office run) needs permission under Codex's global `AGENTS.md`. Unattended sessions can't get it. | Overnight sessions fall back to offline work, evidence ceremony and re-planning. |
| The acceptance bar grew in every revision: 1 run → 3+3 runs → phases 0–5 → D0–D10 with a calibrated AI reviewer, 24 sealed executions and local deployment. | The agent wrote its own finish line and kept moving it. |
| `PresentationRevisionAcceptance.Enabled` turns the feature off at runtime unless `%LOCALAPPDATA%\Scribble\PowerPointAcceptance.json` matches the running DLL's SHA-256. The new route also needs the `SCRIBBLE_ANALYSIS_PILOT=1` environment variable. | Every commit disables the feature until someone re-runs the native harness. No real user machine has it on. |
| The PP01 repair path is written for the fixture: exactly 6 slides, the chart as the last shape on slide 2, replace slide 4, and literal `" \| sales \| "`, `"Cost EUR"`, `"Revenue EUR / Cost EUR (EUR)"` and `"Period"` (`DocumentDraftHost.PilotCopy.cs`, `PresentationDraftCopy.cs`, `AnalysisTableArtifactBuilder.cs`). | A PP01 pass proves nothing about PP02–PP30. Revision 1 put generalization *after* acceptance, so every acceptance would have to be redone. |
| The sealed corpus already holds untouched variants: XA01–XA10 (10 workbooks, 6 domains) and PP01–PP30 (30 decks, 6–12 slides), each with an oracle. | There's no need to write new held-out tasks. XA02–XA10 and PP02–PP10 have never been used for tuning. |
| Measured real-run cost: XA01 ≈ $0.20 (26 requests). PP01 ≈ $0.50 (74–87 requests, and it never reached a terminal event). About $14.47 of the $40 cap is left (22 Sep figure, no paid runs since). | Real runs are cheap. Not running them has cost far more than the runs would have. |

Keep: the architecture (code owns facts, calculations, identity, layout and
execution; the model picks and narrates), the typed contracts, compilers,
source snapshots, journals, recovery, deterministic graders, the stress corpus
and oracles, budget enforcement, and the chart-crash evidence.

## 1. Definition of done — Scribble 2.0 development candidate

Everything below runs on **one installed build** on this workstation. The model
is hosted `qwen/qwen3.8-27b` through the existing Test Lab configuration. Runs
use `Run-StressSuite.ps1` and the corpus oracles, with no manual help during a
run.

| Scenario | Cases | Pass bar |
| --- | --- | --- |
| **S1** Excel analysis → native workbook + four-slide deck | XA01–XA10 | ≥ 9/10 in one batch, plus XA01 passing 3 times in a row on the same build |
| **S2** Workbook-backed repair of an existing deck | PP01–PP10 (6–12 slides, mixed domains) | ≥ 8/10 in one batch, plus PP01 passing 3 times in a row on the same build |

A case passes when:
- every oracle check passes;
- the task reaches a terminal completion event (S1 within 20 minutes, S2
  within 30);
- it stays within the request caps (S1 ≤ 12, S2 ≤ 18);
- the owner accepts the rendered slides (§3 D2).

**Zero tolerance.** A single occurrence of any of these fails the candidate:
- a source file changes;
- a write lands in the wrong document;
- a wrong number is presented as verified;
- success is reported while an oracle check fails;
- any send, save, close or delete action occurs.

**Not part of 2.0.** Each of these needs its own owner request later:
- local Ollama deployment (a 4K-context probe left 0.35 GiB free RAM on this
  PC, so it's a hardware decision);
- validation on the protected work PC;
- Word and Outlook adapter migration;
- AI-reviewer calibration;
- arbitrary chart editing;
- workbooks without a period column;
- public release (frozen at 2.0.91).

## 2. Owner decisions (approved 26 September 2026)

1. **Standing permission for test-owned Office automation: granted.** The
   repo `AGENTS.md` carries it. It covers launching and driving *test-owned*
   Excel/PowerPoint through the repo harnesses and `Run-StressSuite.ps1`. It
   never covers user-owned processes or documents.
2. **Spending: the remaining balance is authorized for this plan**, with a
   soft cap of $4 per day. The estimate in §4 goes beyond the remaining
   balance, so a top-up decision comes before stage D.
3. **Integration: PR #41 is merged** into `codex/development` (`ce85b4f`).
   From now on, one `codex/` feature branch at a time, targeting
   `codex/development`. Prune the Codex worktrees once each is confirmed clean
   and pushed, keeping any ignored evidence or generated corpus they hold.
4. **Model: hosted Qwen for 2.0 acceptance.** Local deployment is deferred
   (§1).

## 3. Work, in order

Stages run in order. Each has a timebox. At 150 % of a timebox, stop and write
a ≤ 10-line owner note in §5.

### Stage A — unblock (≤ 1 day, no paid runs)

- **A1** Integration: PR #41 is merged. Work from a single checkout and one
  feature branch at a time.
- **A2** Remove the runtime receipt gate. Chartless revision is enabled by
  code, not by `PowerPointAcceptance.json`. Chart-mutation operations stay
  rejected in code: keep `RequireSupportedOperations` and make
  `SupportsCharts` return false. The native harness stays a *test*, not a
  runtime switch. Update the affected guardrail tests.
- **A3** Replace the `SCRIBBLE_ANALYSIS_PILOT` environment variable with a
  persisted setting that the Test Lab turns on. A run's trace must show that
  the new route handled the task.
- **A4** Run `Test-DeliveryCandidate.ps1 -RunNative` once as a smoke test
  (4-hour timebox). Fix only what blocks it. If the timebox runs out, go to
  stage B anyway: a real run tells us more.
- **A5** Check the fresh remaining balance with the existing budget check and
  record it in the scoreboard. Build, install and hash-verify the candidate.

### Stage B — S1 on the new route with the real model (2–3 days)

- **B1** Run XA01 on the installed build and add the scoreboard row.
- **B2** On failure, work this loop:
  1. Name the **first failing stage** from the trace: capture/binding,
     planning/schema, calculation, render, native execution, review or
     completion.
  2. Reproduce it offline where possible, as a regression test.
  3. Fix it, rebuild, install and rerun.
  - At most 6 paid runs per day.
- **B3** Review rules for this route:
  - one review/repair round at most;
  - a model finding can block completion only when a deterministic
    measurement confirms it (overflow, overlap, font size, value mismatch);
  - factual findings from the model are ignored, because code owns facts;
  - no full-deck regeneration for a wording or geometry defect.
- **B4** Once XA01 passes, run XA02–XA10 as one batch (≈ $2). Fix only
  mechanisms that generalize. If the new route has 0 passes after 8 real runs,
  stop and escalate to the owner with the traces. Don't quietly switch back to
  the old route.
- **Exit:** S1 pass bar met on one build.

### Stage C — S2 generic repair (3–4 days)

- **C1** Run one native experiment first (≤ half a day). Create the draft by
  opening the saved source as an untitled copy
  (`Presentations.Open(path, ReadOnly:=msoTrue, Untitled:=msoTrue, WithWindow:=msoTrue)`)
  instead of copying slide by slide, which is what crashed `chart.dll`. Try it
  on PPT01–PPT06. Pass condition: PowerPoint doesn't exit, every
  slide/notes/chart is present, and the source is byte-identical.
  - If it passes, it replaces `PresentationDraftCopy`'s six-slide shape
    copying.
  - If it fails, generalize the current shape-copy + rebuild-chart approach to
    any slide and any chart position.
  - Record the result either way.
- **C2** Remove the fixture assumptions:
  - Slide count and chart/replacement slide indices become inspected roles.
  - Charts bind to workbook series by matching categories and series names.
    Reject an ambiguous match before writing.
  - Tables bind by their header row.
  - Style defects are found by measurement (font size, overflow, overlap),
    not by text.
  - Add a guardrail test that fails if product code contains a corpus case ID,
    company name or fixture label.
- **C3** Run PP01 until it passes, then PP01–PP10 as one batch.
- **Exit:** S2 pass bar met on one build.

### Stage D — candidate (1 day)

- **D1** Freeze the build. Run XA01–XA10 and PP01–PP10 once more on it. This
  is the repeatability evidence, and the earlier batches on the same build
  count toward "3 in a row".
- **D2** Owner visual review. Produce one contact sheet per run with the
  existing render tooling. The owner marks each deck accept or reject, which
  takes about 15 minutes. A rejection goes back to B or C as a renderer fix,
  and the affected batch reruns.
- **D3** Update the PR description to the scoreboard plus known limits, and
  merge to `codex/development`. The public freeze is unaffected.

## 4. Budget estimate (measured costs; re-estimate after stage B)

| Stage | Runs | Estimate |
| --- | --- | --- |
| B | ~8 dev + 10 batch + 2 repeats × $0.20 | ≈ $4 |
| C | ~8 dev + 10 batch × $0.20–0.50 | ≈ $4–9 (lower if the 18-request cap holds) |
| D | 10 XA + 10 PP | ≈ $4–7 |
| **Total** | | ≈ $12–20 against ≈ $14.47 left → **a top-up is likely before D** |

## 5. Status (the executing agent updates this section only)

| Scenario | Route | Build | Result |
| --- | --- | --- | --- |
| S1 XA01 | old | 2.0.385 | 1/3 (22 Sep) |
| S1 XA02–10 | — | — | not run |
| S2 PP01 | old | 2.0.395 | 0/3 (no terminal event after 23 min) |
| S2 PP02–10 | — | — | not run (the pilot code rejects every one) |
| New route, real model | — | — | **0 runs** |

Current stage: **A**. The §2 decisions were approved on 26 Sep and A1's merge
is done. Next: A2.

## 6. Rules for autonomous sessions

These are also in `AGENTS.md`.

- **Every session ends with a new scoreboard row** (a real run, or a native
  run in stage A/C1). If it can't, it stops early and writes a ≤ 10-line
  blocker in §5 that says exactly what the owner must do.
- **Progress is measured only by the scoreboard**, never by commits, tests,
  documents or build numbers.
- **Don't:**
  - write new plan, roadmap, checkpoint or phase documents;
  - add gates or acceptance criteria;
  - create new worktrees or long-lived branches;
  - make commits that only archive evidence or add tracing;
  - write fixture-specific code (case IDs, company names, fixture labels,
    fixed slide counts or indices);
  - add a prompt prohibition to fix a single run;
  - weaken a zero-tolerance check to get a pass.
- **Two-strike rule.** If the same first-failing stage returns after two
  fixes, stop patching. Write down the mechanism-level hypothesis, then change
  the approach or escalate.
- **Needs owner approval:** exceeding the daily or total spend, touching
  user-owned Office processes or documents, changing provider or model, raising
  a cap, merging to anything other than `codex/development`, or any public
  release.
