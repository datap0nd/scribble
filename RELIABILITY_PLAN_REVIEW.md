# Review of the Scribble reliability plan

> **Historical review of the 22 September design baseline.** The current,
> consolidated implementation and acceptance plan is [DELIVERY.md](DELIVERY.md)
> in PR #41. Branch prerequisites, balances and implementation-status statements
> below are dated findings, not instructions for the current delivery candidate.

Review target: [RELIABILITY_PLAN.md](RELIABILITY_PLAN.md), as submitted on
22 September 2026.

Checked against: plan baseline `fad5cdd` and head `b873435` of
`codex/stress-suite-200` (draft PR #20), and `codex/development` at `35157be`.
Line numbers refer to `fad5cdd`.

Review status: corrections applied to `RELIABILITY_PLAN.md`. Phase 0 can start
once PR #20 lands in `codex/development` or the owner approves stacking on it.

## Overall verdict

The architecture holds up. Every type, file and test entry point the plan names
exists at the baseline, each row of its evidence table matches a mechanism in
the code, and its central decision (the host owns facts, calculations,
identity, layout and execution; the model owns interpretation and narrative)
targets those mechanisms directly. The existing review filters show why: each
fixes one symptom with a regex over reviewer JSON.

The corrections concern grounding and feasibility, not direction. Five need a
decision before implementation. The rest are factual fixes and clarifications,
applied in place.

## Priority 0: resolve before implementation

### 1. The baseline is not on the development branch

`fad5cdd` is the second-newest commit on `codex/stress-suite-200`, the head of
draft PR #20. At the baseline that branch is 189 commits ahead of
`codex/development` (97 files, +14,070/−492). `codex/development` lacks
`src/Scribble/Office/WorkbookGroupedTotals.cs`, `tests/benchmarks/stress/` (the
corpus, XA01/PP01 and `CHECKPOINT.md`) and
`src/Scribble/Testing/TestLabStressBudget.cs`. `CLAUDE.md` requires `codex/`
feature branches from `origin/codex/development` with PRs targeting it, so
phase branches cut by the rules would not contain the code the plan extends.

Applied: the header, the §7 branching rule, phase 0, §11 and §12 make landing
PR #20 (or an owner-approved stack) a prerequisite. The header also records the
two post-baseline commits (`ebaad7b`, `b873435`). They only touch the stress
budget retry and do not overlap the plan's targets.

### 2. The run evidence is local only

`.gitignore:4` ignores `artifacts/`, so `qa-385-xa01-run1`, `qa-388-pp01-failed`
and `qa-391-pp01-blocked` exist only on the workstation. The committed
checkpoint history records the 2.0.385 and 2.0.391 runs but never 2.0.388, so
the 87 calls, 3,387,583 prompt tokens and 13 providers cannot be checked from
the repository.

Applied: §2 separates code-verified mechanisms from local-trace figures. Phase 0
archives the traces outside the repository and commits only sanitized derived
metrics (suite IDs, assembly hash, `summarize_usage.py` output).

### 3. The evidence and the goal use different model configurations

The plan's goal names "the configured local model", but every live run used
hosted OpenRouter `qwen/qwen3.8-27b`:

- `TestLabStressBudget.cs:44-49` refuses any other endpoint or model for the
  stress corpus.
- `OpenAiCompatibleClient.cs:1118-1232` applies OpenRouter-only overrides,
  selected by `UsesOpenRouterQwenPolicy` (`:1333-1357`). Reasoning is disabled
  for native deck drafting and set to `none` for compact tool-less calls
  (`:1175-1180`). Draft calls get 32K/8K `max_tokens`, tool calls are serial,
  and only tool-bearing requests are held to a six-provider allow-list with
  `allow_fallbacks: true` (`:1199-1231`).

None of this applies to a local endpoint. The routing also explains the
13-provider figure: tool-less reviewer calls have no provider restriction.

Applied: §1 and a new §6 subsection separate the evaluation configuration from
the deployment target. Phase 0 chooses the acceptance configuration. Phase 6
changes the stress gate only if a local endpoint is the target. §6 notes that
pinning an OpenRouter provider means `allow_fallbacks: false`.

### 4. The live ladder does not fit the remaining budget

`TestLabStressBudget.MaximumTotalUsd` is a $30 hard cap
(`TestLabStressBudget.cs:20`) with a $0.05 reserve. The latest checkpoint
(`b873435`) reports about $5.62 remaining at 17:51 UTC. Balance deltas in the
checkpoint history imply roughly $0.35–0.50 per XA01/PP01 run: $0.71 between
15:40 and 16:47 UTC (XA01 and PP01 on 2.0.385) and $0.96 between 16:47 and
17:32 UTC (PP01 on 2.0.388 and 2.0.391). The ladder needs at least 30 runs
(6 smoke, 24 held-out), about $10–15 at those rates.

Applied: §8 states the budget fit and makes it a recorded decision before
phase 6: demonstrate efficiency first, obtain a new authorization, or evaluate
on a local runtime. §10 states the cap, reserve and balance. Phase 9 notes that
the full corpus needs separate funding, and §11 adds the risk.

### 5. The guardrail scan pins tool names

`scripts/Test-Guardrails.ps1:499-548` requires the exact set of
`public const string` names in `WorkbookToolCatalog`, `PresentationToolCatalog`,
`CrossAppToolCatalog` and `WordToolCatalog`. Phase 2 adds "adapters beside
`WorkbookToolCatalog`, `CrossAppToolCatalog`"; any new tool constant there fails
CI.

Applied: a §7 tool-surface rule (prefer typed arguments on existing tools;
otherwise a reviewed allow-list update in the same PR). The phase 2 exit gate
includes the scan, and §11 and §12 repeat the rule.

## Priority 1: factual corrections applied

| Plan statement | Finding at the baseline | Change |
| --- | --- | --- |
| Phase 0: "Add sanitized offline regressions" for percentages/rankings, dates and page numbers | Already present, written against the current filters: `SamsungRecoveryTests.cs:287-292` (55.74%/55.76%), `SamsungWorkflowTests.cs:81-92` (22,675/22,044), `SamsungWorkflowTests.cs:64-73` (ISO due dates), `SamsungRecoveryTests.cs:305-316` and `SamsungWorkflowTests.cs:530-534` (folio). Missing: context growth, geometry routing, Excel ordering | §2 and phase 0: restate the existing tests as mechanism-level assertions; add the three missing ones using `FakeEndpoint` (`Program.cs:8768`) and `CrossAppFixture` |
| Phase 0 gate: "All five failure mechanisms" | §2 lists seven; the §12 handoff names four | Gate and handoff cover all seven |
| "The existing PowerPoint path requires vision; if the chosen local model lacks it, the architectural pilot uses native assertions plus an identified human visual reviewer" | The path refuses to run without vision: `SLIDE_VISION_REQUIRED` at `DocumentDraftHost.PowerPoint.cs:53-54` and `:230-231`. `ModelCatalog.IsVisionCapable` (`ModelCatalog.cs:180`) is name-based, and `ModelContractProbe.cs:70` records `vision_certified = false` | §6: probe vision in phase 0; the non-vision mode needs a new flagged code path |
| "Support capability-probed strict structured output" | No `response_format`, `json_schema` or `strict` anywhere in `src/Scribble`, and the probe does not test them. `OpenAiCompatibleClient.cs:1195`: `require_parameters` cannot be used for Qwen 3.8 routes | §6: marked as new work, with the OpenRouter constraint |
| "For six-slide repair, work in a marked draft copy of the original deck … and patch only requested defects" | The current PP01 route, `ShouldDraftRepairedDeck` (`DocumentDraftHost.PowerPoint.cs:19-24`, an instruction regex), creates a new deck and regenerates all six slides. The PP01 prompt asks for "a repaired, editable draft of every slide" in the Samsung MD style | §6 and phase 4: stated as a change of approach built on `PresentationRevision` (at most 24 operations, `PresentationRevision.cs:46`). Phase 4 defines patches versus requested recreation, and its gate requires PP01 grader acceptance |
| Invariant: "Bind execution to captured document identities" | `write_cells` uses `application.ActiveSheet` (`WorkbookDraftWriter.cs:418`) and `write_draft_sheet` uses `application.ActiveWorkbook` (`:95-97`). Neither keeps a before-image | §3: the invariants bind the new path; phase 5 closes the existing gaps; the pilot writes only to new documents |
| "Generic read capture flattens string leaves" | It also drops object keys and numeric, boolean and null leaves (`TaskSources.cs:141-148`). The attached `.xlsx` extractor drops empty cells without cell references (`EmailAttachmentReader.cs:1950`) and emits date serials and no formulas; `.xlsb` also skips empty values (`XlsbTextExtractor.cs:233`, `:247`). Live reads use `Value2` only (`WorkbookToolHost.cs:970`, `:1033`). `CompleteWorkbookTotals` runs only on complete attached workbooks (`TaskSources.cs:79-87`) | §2 capture-gap list, phase 1 integration points, and a new offline test bullet |
| "Page metadata must include an integer `expected_page_number`" | Correct, but a field named `expected_page` already exists with a different meaning, the page's elements (`DocumentDraftHost.SlideRepair.cs:462-464`, `:492-494`). The identities already live in `SamsungGenerationJournal.Receipt` (`SamsungGenerationJournal.cs:12`) and `OwnedPageReceipt` (`DocumentDraftHost.SlideRepair.cs:210-218`) | §5: do not overload `expected_page`; carry the journal identities into findings |
| "Introduce `analysis_v1`/`document_plan_v1` schema versions" | Versioning already exists: `DurableTaskState.SamsungWorkflowVersion` (`TaskContextManager.cs:36`, `SamsungAuthoringPolicy.cs:15`, checked at `SamsungGenerationJournal.cs:60`) routes pre-v2 tasks to `LegacySamsung` | §9: persist the new contract version beside it and route by it |
| Phase 7 "Migration and release … Tested release candidate … release rollback" | Public release is frozen at 2.0.91 and only a new explicit owner request lifts it (`docs/release-channels.md`). CI has no release job | Phase 7 becomes development-candidate readiness; §9 states the freeze |
| `SamsungSlideDesign`, `PresentationDraftWriter.Samsung` | Both exist twice: workflow 2 in `src/Scribble/Office/` and the frozen v1 in `src/Scribble/Office/LegacySamsung/` | The header disambiguates; the handoff forbids changing `LegacySamsung/` |
| "These observations were verified in source code and native-run traces" | The traces are not in the repository (Priority 0, item 2) | §2 separates the two sources |
| "The CI build uses MSBuild for the solution and separate browser-extension checks" | Correct but incomplete. CI also runs `GuardrailTests.exe`, `Test-Guardrails.ps1`, the Python evaluator and catalog tests, the Test Lab script tests and the release gates, on `windows-latest` without Office | §8 lists the CI steps and states that native checks are workstation-only |
| Phase 3 integration points | They omit the deterministic evidence checker (`SamsungPresentationReview.ValidateEvidence` and `Numbers`, `SamsungPresentationReview.cs:46-89`, `:232-241`) and `SamsungRepairPolicy` | Both added to phase 3, which also retires the review filters |

## Priority 2: clarifications applied

- §2 names the current heuristic mitigations: `FilterBriefRefutedReview` (the
  baseline commit itself; the 2.0.392 PP01 run stopped at a budget-check
  timeout before authoring, so no native run has validated it),
  `FilterNativeRefutedReview`, `FilterReviewFindings`,
  `OnlyHostOwnedPageNumberBlockers`, `OnlyOtherSlideCoverageBlockers`,
  `OutlineReviewApprovedOrDeterministicallySatisfied`,
  `ReviewApprovedOrSatisfiedPromptConstraint` and `ShouldDraftRepairedDeck`.
  Phase 3 and §9 retire them per capability.
- §6 documents today's nested repair budget: three cycles per native slide
  (`DocumentDraftHost.SlideRepair.cs:280-286`), two proposals per cycle
  (`:88`), two syntax-only JSON retries per parse (`:178-207`), a fact review
  per cycle (`:146-152`), and a separate three-cycle limit for staged
  `revise_slides` batches.
- §5 adds repository conventions: classic csproj registration for source and
  test files, `JavaScriptSerializer` as the only JSON library (so immutability
  comes from IDs and hashes), `ToolContractValidator`, and public contract
  types so guardrail tests need no reflection.
- §5: the new Excel compiler emits invariant `Formula` text only.
  `FormulaLocal` (`WorkbookDraftWriter.cs:269`) and `TryRepairAdjacentRowFormula`
  (`:322`) stay out of the new path.
- §8 step 2 keeps the checkpoint's "three consecutive unassisted passes, reset
  on failure" semantics.
- §8 step 3: held-out tasks are authored fresh, sealed with the existing
  manifest-hash mechanism and kept outside the implementation branch.
  Cross-app held-out tasks are Excel↔PowerPoint only until phase 8 migrates Word
  and Outlook.
- Efficiency targets: phase 0 computes the call floor of the stage design,
  counting chat-loop turns, before the ≤12/≤18 targets are relied on.
- Renderer fixture families map to existing layout IDs. Fixtures build on the
  `SamsungSlideTests` previews, `Test-StressNativeGrading.ps1` and
  `SamsungBenchmarks.json`.
- The existing-assets list adds `PresentationRevision`,
  `LegacyOfficeTextExtractor` (the typed `.xls` precedent,
  `LegacyOfficeTextExtractor.cs:210-212`), workflow versioning,
  `summarize_usage.py`, and the `CrossAppFixture`/`FakeEndpoint` test doubles.
- §11 adds four risks: the PR #20 dependency, the budget, hosted-versus-local
  representativeness, and the guardrail scan.

## Verified as accurate

- `fad5cdd` exists on `codex/stress-suite-200` ("Prefer verified slide
  arithmetic over stale brief numbers").
- All named types exist at the baseline: `TaskSources`, `WorkbookGroupedTotals`,
  `SamsungSlideDesign`, `PresentationDraftWriter.Samsung`,
  `SamsungGenerationJournal`, `DurableExcelTransform`, `ExcelTransformTarget`,
  `WorkbookDraftWriter`, `DocumentDraftHost` (including `.PowerPoint` and
  `.SlideRepair`), `WorkbookToolCatalog`, `CrossAppToolCatalog`,
  `SamsungWorkflowSchema`, `SamsungAuthoringPolicy`, `SamsungEvidence`,
  `PresentationRevision` and `ModelContractProbe`.
- All named test entry points exist, including the Test Lab runner described
  in `tests/benchmarks/stress/README.md`.
- `CompleteWorkbookTotals` matches exactly `RowID/Period/Group/RevenueEUR/CostEUR`
  (`TaskSources.cs:192-196`), and identical text shares provenance (`:44-47`).
- The date-suffix mechanism is documented in the code itself
  (`SamsungEvidence.cs:65-68`).
- 55.74% versus 55.76% is the recorded root cause of the 2.0.391 PP01 stop
  (`CHECKPOINT.md`).
- Excel validation follows mutation (`WorkbookDraftWriter.cs:245-322`); the
  partial state lands on a new, marked `Scribble Draft` sheet.
- Some author and reviewer calls disable reasoning
  (`OpenAiCompatibleClient.cs:1175-1180`).
- `DurableExcelTransform` is specialized staged capture/write/readback. COM work
  runs on the `OfficeThread` STA, and evidence lives in `TaskCheckpointStore`
  behind `read_task_evidence`.
- XA01 is Excel to a four-slide deck, and PP01 repairs a six-slide deck
  (`CHECKPOINT.md`).
- The two Anthropic references resolve and support the points cited.

## Not verifiable from the repository

- 87 model calls, 3,387,583 prompt tokens and 13 providers for the 2.0.388 PP01
  run: local trace only. The provider spread is consistent with the routing
  code.
- The vLLM strict-mode link: `docs.vllm.ai` was blocked by the review
  environment's network policy. The plan already treats strict mode as a
  capability to verify on the chosen server.
- Per-run cost: inferred from checkpoint balance deltas, assuming no other
  spend in those windows.

## Method

The baseline was checked out in a separate worktree and compared with
`origin/codex/development` and `origin/codex/stress-suite-200`. Every cited path
was read at the baseline. The Linux review environment cannot build or run
Scribble, so no build, test or model call was made; CI remains the compile
gate.
