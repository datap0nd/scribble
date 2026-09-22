# Scribble reliability: architecture and implementation plan

Prepared 22 September 2026. Reviewed and corrected the same day against the
repository; [RELIABILITY_PLAN_REVIEW.md](RELIABILITY_PLAN_REVIEW.md) lists each
correction and its evidence.

**Planning baseline:** `fad5cdd859713561af704e08674936d719b2d5ca` on
`codex/stress-suite-200`, the head branch of draft
[PR #20](https://github.com/datap0nd/scribble/pull/20) into `codex/development`.
Repository `datap0nd/scribble`, local clone
`C:\Users\keeoh\Documents\ChatGPT\Scribble`.

At review time the branch head was `b873435`. The two commits after the baseline
(`ebaad7b`, `b873435`) add bounded retries to the stress budget check and touch
only `src/Scribble/Testing/TestLabStressBudget.cs`,
`tests/GuardrailTests/StressGradingTests.cs` and
`tests/benchmarks/stress/CHECKPOINT.md`. They do not overlap the files this plan
refactors.

**The baseline is not on `codex/development`.** At the baseline, PR #20 is 189
commits ahead of `codex/development` (97 files, +14,070/−492 lines).
`codex/development` does not yet contain `WorkbookGroupedTotals.cs`, the 200-case
stress corpus in `tests/benchmarks/stress/` (including XA01, PP01 and
`CHECKPOINT.md`), or `TestLabStressBudget.cs`. Repository rules (`CLAUDE.md`,
`docs/release-channels.md`) require `codex/` feature branches from
`origin/codex/development` with PRs targeting it. Implementation therefore starts
after PR #20 lands, or after the owner explicitly approves stacking phase branches
on `codex/stress-suite-200` until it does.

The existing implementation task is active. This is a design and coding-agent handoff, not an instruction to overwrite its current work. Reconcile changes made after this baseline before implementation. Creating this plan does not change Scribble, stop another task, or authorize additional API spending.

File paths are relative to the repository at the baseline. `SamsungSlideDesign`,
`PresentationDraftWriter` (including `PresentationDraftWriter.Samsung.cs`),
`SamsungPresentationReview` and `MetoTheme` mean the workflow-2 files in
`src/Scribble/Office/`. The same-named files in `src/Scribble/Office/LegacySamsung/`
are the frozen v1 implementation for tasks started before workflow 2; do not
modify them.

## 1. Outcome and engineering decision

Make routine Office work reliable with the user-configured model: create an accurate, editable Excel analysis; use the same analysis in a well-designed PowerPoint deck; and make bounded edits without corrupting source work or entering repetitive review loops. The deployment target is a local Qwen-family model behind an OpenAI-compatible endpoint. All live evidence so far comes from the hosted OpenRouter `qwen/qwen3.8-27b` configuration that the stress harness requires; §6 separates the two.

The core change is to make Scribble own facts, calculations, document identity, layout mechanics, and execution. The model owns interpretation, narrative, and choices among supported operations. Extend the existing implementation rather than replacing the Office integration or building another renderer.

The first proof is **one verified analysis → Excel report and four-slide deck**, followed by **repairing a six-slide deck using an authoritative workbook**. These are development targets, not the boundaries of the eventual product. Generalization must be demonstrated with unseen inputs.

Success means correct native outputs, preserved sources, acceptable rendered quality, bounded execution, and honest completion status. A new build, a model saying “done,” or a single lucky run does not establish success.

## 2. Evidence and root causes

The mechanisms below were verified in source code at the baseline. The run figures (call and token counts, provider count) come from native-run traces on the local workstation. `artifacts/` is gitignored, so those figures cannot be reproduced from the repository until phase 0 archives them. The plan addresses mechanisms, not individual benchmark values.

| Observed failure | Architectural cause | Required correction |
| --- | --- | --- |
| Reviewer demands 55.74%; host calculation requires 55.76% | Model-authored brief competes with verified data as an authority | Separate requirements, narrative plans, and facts; calculations have a single authoritative result |
| Reviewer says 22,044 exceeds 22,675 and rejects the correct ranking | Probabilistic reviewer can overrule deterministic comparisons | Rankings and arithmetic belong to code; review objections must identify a checkable source or binding error |
| Date suffixes such as `-10` and `-16` (from `2026-07-10`, `2026-07-16`) become unsupported numbers | Numeric scanning (`SamsungPresentationReview.Numbers`) cannot distinguish dates from quantities | Typed dates, identifiers, quantities, and text segments |
| Slide six is reviewed against footer one | Native identity, logical identity, and display order are conflated in review payloads: each logical slide is composed as page `- 1 -`, and the reviewer's `expected_page` field carries page elements, not a number | Explicit slide/page identities and typed page-number metadata |
| Model is asked to fix fixed-position titles and chart-highlight boxes | Repair authority does not match defect ownership; every visual repair asks the model for a complete slide | Geometry repairs belong to the renderer; model repairs change supported semantic fields |
| Six-slide run (PP01, 2.0.388) makes 87 model calls and processes 3,387,583 prompt tokens, including repeated/cached context | Duplicated evidence, nested reviews, large payloads, and regeneration | Scoped context, fact references, bounded stage calls, targeted patches |
| Ordinary Excel write can discover formula errors after some cells were written | Validation and mutation are interleaved: `WorkbookDraftWriter` writes the value grid, then formulas, then detects rejected or error-valued formulas | Stage, validate, commit, read back, and reconcile against an explicit document target |

Trace baselines, local only: `artifacts/qa-388-pp01-failed/`, `artifacts/qa-385-xa01-run1/`, and `artifacts/qa-391-pp01-blocked/`. The committed checkpoint records the 2.0.385 XA01 run (`suite-20260922-163550-3fa83055`) and the 2.0.391 PP01 run (`suite-20260922-172330-f0091961`). The 2.0.388 PP01 run appears only in the local trace. It used 13 hosted providers for the same Qwen model. That is consistent with current routing: `OpenAiCompatibleClient` restricts only tool-bearing OpenRouter requests to a six-provider allow-list with fallbacks enabled, and tool-less internal reviewer calls use OpenRouter's default routing. Reasoning is disabled for native deck-drafting calls, set to `none` for compact tool-less internal calls (at most 2,048 output tokens) and `low` otherwise. This is not a controlled benchmark of the best local runtime configuration.

**Current mitigations are heuristics.** The baseline commit itself adds `FilterBriefRefutedReview`, which drops reviewer findings that repeat stale brief numbers. Its first native PP01 run (2.0.392) stopped at a budget-check timeout before authoring, so no native run has validated it yet. Similar post-hoc filters over reviewer JSON cover native state, page numbers, other-slide coverage and deterministic calculations: `FilterNativeRefutedReview`, `FilterReviewFindings`, `OnlyHostOwnedPageNumberBlockers`, `OnlyOtherSlideCoverageBlockers`, `OutlineReviewApprovedOrDeterministicallySatisfied` and `ReviewApprovedOrSatisfiedPromptConstraint`. An instruction regex (`ShouldDraftRepairedDeck`) chooses the PP01 repair route. This plan replaces them with typed contracts and retires each one when its capability migrates.

**Existing offline regressions.** Four mechanisms already have guardrail regressions, written against those filters: stale brief (`tests/GuardrailTests/SamsungRecoveryTests.cs:287-292`), ranking (`SamsungWorkflowTests.cs:81-92`), ISO due dates (`SamsungWorkflowTests.cs:64-73`) and native folio numbering (`SamsungRecoveryTests.cs:305-316`, `SamsungWorkflowTests.cs:530-534`). None covers context growth, geometry-repair routing or Excel write-then-validate ordering.

### Preserve these existing assets

- `TaskSources` (`src/Scribble/Chat/TaskSources.cs`): captured source text and provenance spans.
- `WorkbookGroupedTotals` (`src/Scribble/Office/WorkbookGroupedTotals.cs`, PR #20 only): deterministic decimal aggregation that counts and excludes blank or non-numeric cells rather than treating them as zero.
- `SamsungSlideDesign` and `PresentationDraftWriter.Samsung` (workflow 2): native editable layouts and rendering.
- `SamsungGenerationJournal`: write receipts carrying native `SlideId`, logical `SourceId`, page index, page ordinal and fingerprint; resume and user-edit checks. `OwnedPageReceipt` (`DocumentDraftHost.SlideRepair.cs`) keeps the same identities for review.
- `PresentationRevision` (`revise_slides`, `revert_scribble_changes`): staged edits on an unsaved working copy, mandatory target IDs and original fingerprints, journaled batches of at most 24 operations, and native revert.
- `DurableExcelTransform`: staged capture/write/readback machinery for specialized column transformations such as workbook translation.
- `LegacyOfficeTextExtractor`: `.xls` capture through ExcelDataReader that already keeps column positions, ISO dates and a per-cell type, including `blank`.
- Workflow versioning: `DurableTaskState.SamsungWorkflowVersion` routes tasks started before workflow 2 to the frozen `LegacySamsung` path.
- `ModelContractProbe` and `tests/benchmarks/stress/summarize_usage.py` (per-response model, provider, token and cost accounting).
- Native artifact readbacks, stress fixtures, `Test-StressNativeGrading.ps1`, `tests/NativeAcceptance/SamsungBenchmarks.json`, Test Lab, the `CrossAppFixture` and `FakeEndpoint` test doubles, and existing source-preservation checks.

The problem is how these components exchange information and authority. Their presence makes an incremental migration practical.

The first data phase must fix these capture gaps, not add more benchmark-specific header recognizers:

- `TaskSources.CompleteWorkbookTotals` recognizes only the exact headers `RowID/Period/Group/RevenueEUR/CostEUR` with `YYYY-MM` periods, and only for complete attached workbooks (`CaptureInput`). Live Excel reads never reach it.
- `TaskSources.CaptureRead` flattens tool results through `Collect`, which keeps only string leaves. Object keys and numeric, boolean and null leaves are dropped; the rest is joined into untyped text.
- The attached `.xlsx` extractor (`EmailAttachmentReader.ExtractXlsxSheet`) skips empty cells without tracking cell references, so a blank cell shifts later values under the wrong header. It emits stored values only: dates as serial numbers, no number formats, no formulas. `XlsbTextExtractor` also skips empty values.
- Live Excel reads (`WorkbookToolHost`) render `Value2` as TSV text. Dates arrive as serial numbers; number formats, formulas and cell types are not captured.
- `TaskSources.Add` identifies a source by its content fingerprint and returns the existing span IDs for identical text, so two documents with the same text share one provenance label.

Replace the special case with typed table capture, starting at the extractors and COM reads, and reusable aggregate operations. Preserve distinct provenance when two documents contain identical text.

## 3. Scope and invariants

### Initial production scope

1. Excel summaries, grouped totals, comparisons, live formulas, tables, and charts.
2. New corporate slide decks from verified workbook analysis.
3. Source-preserving repair of an existing deck, including factual updates from a specified workbook.
4. Explicitly requested, bounded Excel edits on a captured workbook and worksheet.

Word, Outlook-origin requests, and other source types retain existing behavior during the pilot. Their adapters migrate after the shared analysis path passes. Receiving a request from Outlook must not require regenerating the analysis just to hand it to Excel or PowerPoint.

### Non-negotiable invariants

- Never turn model-generated prose into a stronger authority than source evidence or verified calculations.
- A fact ID cannot change its meaning, value, unit, period, or source revision in place.
- Preserve missing values; distinguish zero, blank, error, unknown, and not applicable.
- Retain live formulas when requested. Never substitute answer constants to make a test pass.
- Default to marked, unsaved drafts. Modify an original only within the user's existing explicit edit authorization.
- Do not send mail, save original documents, publish, or change unrelated content through this work.
- Bind execution to captured document identities; changing the active window must not redirect a write.
- A retry cannot duplicate or silently replace an uncertain write.
- Native COM work stays serialized on the owning Office thread.
- No output is marked complete until its requested checks have actually finished.
- User/source text remains data. It cannot introduce tool instructions or change authorization.

These invariants bind the new path from its first commit. Existing write paths do not all meet them yet. `write_cells` writes into `application.ActiveSheet` at call time, and `write_draft_sheet` appends to `application.ActiveWorkbook` unless it creates a new workbook; neither keeps a before-image. Phase 5 closes these gaps. Until then the pilot writes only to new, marked workbooks and presentations.

## 4. Target architecture

```text
User request
    ↓
Task requirements + authorized source/destination bindings
    ↓
Captured source snapshots → verified facts/tables → analysis calculations
    ↓
Model proposes narrative + semantic document operations
    ↓
Typed plans reference the verified analysis
    ├── Excel compiler → formulas, ranges, formats, charts
    └── PowerPoint compiler → content blocks, layouts, charts, citations
    ↓
Staged native execution → readback → deterministic checks
    ↓
Scoped visual/narrative review → supported targeted repair, if needed
    ↓
Verified draft + precise completion receipt
```

“Compiler” here means ordinary C# code that converts a small, typed plan into the existing Office writer operations. It does not require a new language platform.

### Authority rules

| Authority | Owns | Cannot do |
| --- | --- | --- |
| User request and accepted clarifications | Purpose, scope, source precedence, output constraints | Make contradictory source values mathematically consistent |
| Captured source records | What the source says, where, and at which revision | Establish that every source assertion is objectively true |
| Verified analysis | Results calculated from declared inputs and operations | Invent the interpretation of ambiguous columns or periods |
| Narrative plan | Selection, order, emphasis, wording | Override facts or create new user requirements |
| AI review | Unsupported prose, ambiguity, relevance, visual observations | Override arithmetic, native identities, or user authorization |

If a reviewer finds a possible binding error, route it back to the source/analysis stage with exact IDs. Resolve or expose the conflict there. Do not correct the number by accepting the reviewer's replacement prose.

## 5. Canonical data and plan contracts

Names below are proposed types; adapt to repository conventions:

- .NET Framework 4.8 and C# 7.3 (no target-typed `new`, ranges or switch expressions); string concatenation rather than interpolation. Keep this unless a separate change is justified.
- Classic csproj files: add each new source file to `src/Scribble/Scribble.csproj` and each new test file to `tests/GuardrailTests/GuardrailTests.csproj`, and register new tests in `Main` in `tests/GuardrailTests/Program.cs`.
- JSON is `JavaScriptSerializer` throughout; no other JSON library is referenced. It requires public setters, so enforce immutability with host-issued IDs, content hashes and the absence of update operations, not read-only DTOs. Validate tool arguments with `ToolContractValidator`.
- Guardrail tests reach code through public APIs or reflection (no `InternalsVisibleTo`). Make the new contract types public.

### SourceSnapshot

Store `snapshot_id`, source instance identity, source type, capture revision, capture time, content hash, supported coverage, calculation state, and locators. Excel locators identify workbook instance, worksheet identity, range, formula/value, raw cell type, number format, cached value, and relevant table headers. PowerPoint locators distinguish presentation identity, stable SlideID, shape ID, and displayed ordinal. A stale formula cache is not a verified recalculated value.

Live Excel reads currently capture only `Value2`; typed capture adds `Formula`/`HasFormula`, `NumberFormat`, the displayed `Text` and the raw value type. COM returns doubles, so keep the raw value and the displayed text separately. `inspect_slide` already returns stable presentation, slide and shape identifiers and fingerprints for the PowerPoint locators.

Content hashes establish integrity, not source identity: two files containing the same numbers must remain distinguishable. Unsaved Office documents need runtime identities. Fingerprint the relevant read/write scope; unrelated edits should not unnecessarily invalidate the whole task.

### Fact and TableDataset

A fact includes:

- Host-issued ID and snapshot binding.
- Kind: observed source value, user-supplied sample, derived calculation, or unresolved assertion.
- Value type: decimal, integer, date, text, boolean, missing, or error.
- Metric/field identity, unit, currency where applicable, period, dimensions, and display precision.
- Source locators and status describing what has actually been verified.

Tables carry typed columns, row identities, dimension keys, and cells referencing facts. The host issues short readable aliases for model use; clients cannot register invented facts as verified records.

Extraction is not automatically verification. For ambiguous headers, PDF text, or prose, retain the exact evidence and unresolved mapping. Use a targeted clarification only when intent cannot be established from the available sources. Do not introduce routine approval prompts for clear requests.

### Calculation and AnalysisArtifact

Calculations contain an operation, input fact/dataset IDs, grouping/filter semantics, rounding policy, and a host-computed output. Start with the operations already required by the pilot: sum, filtered/grouped sum, count with explicit null policy, difference, ratio, growth, margin, ranking, and comparisons.

The model proposes expressions but does not supply the authoritative result. Check unit compatibility, periods, grouping dimensions, zero denominators, duplicates, missing values, ties in rankings, percentage versus percentage-point changes, and calculation cycles. Keep source precision until presentation formatting.

An immutable `AnalysisArtifact` bundles snapshot references, verified datasets, calculations, assumptions, unresolved conflicts, and schema version. Excel and PowerPoint consume its ID. A source revision creates a new analysis revision and invalidates affected plans.

### Document plans

`WorkbookPlan` uses semantic columns, rows, formulas, and charts. The compiler allocates cell addresses and generates native formulas. Supported formula operations must cover the actual pilot, including aggregate rates and references to source data. Advanced formulas may use an explicit validated expression path; unsupported operations return a capability error rather than silently becoming values. The compiler emits invariant-culture `Formula` text only. Do not carry `WorkbookDraftWriter`'s `FormulaLocal` fallback or its post-write `TryRepairAdjacentRowFormula` header-offset repair into the new path.

`DeckPlan` separates user requirements from narrative choices. Slide blocks are typed: headline, metric, table, chart, explanatory text, image, or source disclosure. A metric/chart/table references analysis objects. Text can contain typed fact references that the host formats; dates and document identifiers are not scanned as free-floating quantities.

Illustrative model-facing plan fragment:

```json
{
  "analysis_id": "analysis_17",
  "slides": [{
    "id": "june_results",
    "layout": "scorecard",
    "title": "June performance",
    "metrics": ["revenue_june", "cost_june", "margin_june"]
  }]
}
```

The final names and JSON shape should be tested with the target model. The important property is that numeric values and citations are resolved by the host rather than copied into multiple model fields.

### Findings and patches

Use structured findings with `code`, owner (`analysis`, `content`, `renderer`, `execution`), target ID, evidence/measurement, severity, and supported action. Example: `CHART_HIGHLIGHT_BOUNDS` targets a chart and is owned by the renderer; `UNSUPPORTED_CAUSAL_SENTENCE` targets a text block and is owned by content.

Patches operate on a specific field or block. Do not regenerate immutable data and citations when shortening a title. Page metadata must include an integer `expected_page_number`, separate from native slide ID and logical slide ID. Do not reuse the existing `expected_page` review field, which carries the page's text, table and chart elements. The journal already stores native `SlideId`, logical `SourceId`, page index and ordinal; carry those into findings rather than parsing the `- N -` footer string.

## 6. Runtime, review, and recovery

### Model interaction

- Expose only the tools and schema needed at the current stage.
- Return relevant typed facts and summaries; fetch additional evidence by ID when needed.
- Keep full evidence in the existing protected evidence store (`TaskCheckpointStore`, `read_task_evidence`), not repeatedly inside prompts.
- Support capability-probed strict structured output where the selected server supports it. Keep a bounded parse/validation fallback for servers that do not. JSON validity does not establish factual correctness. This is new work: nothing in `src/Scribble` sends `response_format`, `json_schema` or tool `strict` today, and `ModelContractProbe` does not test them. On OpenRouter, `provider.require_parameters` cannot currently be used for Qwen 3.8 routes because none advertises `parallel_tool_calls` (see the routing comment in `OpenAiCompatibleClient`), so provider-side enforcement cannot be assumed there.
- Record model, provider/runtime build, tokenizer/chat template, quantization, context limit, reasoning settings, sampling, and output limits for comparisons. `summarize_usage.py` already extracts per-response model, provider, tokens and reported cost; extend it rather than building parallel accounting.
- For local evaluation, pin one configuration. For hosted diagnostics, pin a supported provider where possible: on OpenRouter that means a single-provider `order` with `allow_fallbacks: false`, which gives up the current fallback-based availability. If fallbacks stay on, record the provider that served each request and compare only like-for-like runs. Do not mix configurations and attribute the difference solely to code.
- Prompt corrections must describe one structured defect. Avoid growing the universal system prompt after each failure.

### Evaluation configuration and deployment target

Every live run in the evidence used hosted OpenRouter `qwen/qwen3.8-27b`. `TestLabStressBudget` refuses any other endpoint or model for the stress corpus, and `OpenAiCompatibleClient.UsesOpenRouterQwenPolicy` applies OpenRouter-only overrides: reasoning settings, 32K/8K `max_tokens` for draft calls, the provider allow-list and serial tool calls. None of these applies to a local endpoint, so hosted results do not transfer automatically. Phase 0 decides which configuration is the acceptance target. Running the stress corpus against a local endpoint needs a reviewed change to the stress budget gate and its guardrail test, without weakening the paid-key cap for the hosted endpoint.

Discover the actual local endpoint, model, hardware, vision support, context capacity, and schema-enforcement behavior in phase 0. The existing PowerPoint path refuses to run without vision: `DocumentDraftHost.PowerPoint.cs` throws `SLIDE_VISION_REQUIRED` during argument validation and again before the native write. `ModelCatalog.IsVisionCapable` decides this from the model name and catalog, not from a probe, and `ModelContractProbe` records `vision_certified = false`. Phase 0 must therefore probe vision with a real image. If the chosen model lacks vision, the pilot's "native assertions plus an identified human visual reviewer" mode needs an explicit, flagged code path with its own guardrail test; it does not exist today. Any separate local vision model must be an explicit, measured configuration. Do not silently route visual review to a hosted service. Fully automated visual review remains unproven until its configured path passes calibration.

### Review boundaries

Deterministic checks own numeric values, dates, periods, totals, rankings, required counts, data bindings, object identity, and source-preservation comparisons. Model review handles narrative and visual judgments, using only relevant evidence and rendered pages.

A reviewer may flag a wrong label-to-column mapping or unsupported conclusion. It must identify the affected binding; it cannot replace a verified value from memory. Aesthetic observations route to supported layout alternatives or renderer operations. Unresolvable defects remain explicit failed criteria; the system must not conceal them by deleting requested content.

At the baseline this rule is approximated by the post-hoc filters listed in §2. Replace them with typed findings that cannot carry replacement values for fact-bound fields, and retire each filter when its capability moves to the new contract.

Today each native slide gets up to three repair cycles (`samsung_repairs:<SlideID>`); each cycle allows two repair proposals, each parse up to two syntax-only JSON retries, and a fact review. Staged `revise_slides` batches have their own three-cycle limit. Proposed global repair policy: at most one normal corrective model patch per affected block, with an explicit task-level repair budget. Schema retries, page reviews, and final review all consume that budget. There is no nested reset that grants three more cycles inside another three-cycle loop. Renderer fitting can try a finite ordered list of valid variants without an LLM call.

### Native execution

Use a durable state machine:

`Captured → Planned → Validated → Staged → Applying → ReadbackVerified → Complete`

Exceptional outcomes: `Conflict`, `Unsupported`, `NeedsInspection`, or `Failed`. “Applying” after a crash is uncertain until reconciled; it is not evidence that an operation succeeded or failed.

Each operation records an operation ID, target binding, expected pre-state, intended patch, result fingerprint, and verification receipt. Before mutation, confirm the relevant source, destination, and formula-dependency fingerprints. Retain before-images of every touched property needed for supported recovery: formulas, values, formats, dimensions, and owned objects as applicable. Formula staging must preserve workbook reference semantics; an isolated scratch cell is insufficient for formulas that depend on sheets, names, tables, or workbook settings.

COM does not provide a general atomic transaction. Use staged preparation and compensating restoration where proven safe. On interruption, compare actual state with before/after expectations; resume or acknowledge a verified operation once. If concurrent user edits prevent safe restoration, preserve the document and explain the exact conflict. Do not claim rollback succeeded without readback evidence.

Declare capability handling for merged/protected ranges, spill arrays, external links, volatile functions, and UDFs. Preserve or reject unsupported cases explicitly; do not pretend that an arbitrary formula has been validated. Keep “applied in memory,” “verified,” and “saved” as distinct states.

For six-slide repair, work in a marked draft copy of the original deck, preserve unchanged notes/artwork/hyperlinks, and patch only requested defects. Unsupported objects remain preserved or are explicitly identified as unsupported. Do not approximate them silently while reconstructing a deck.

This changes current behavior. Today `ShouldDraftRepairedDeck` routes PP01 to a new, separate deck and regenerates all six slides through `add_draft_slides`. Copy-and-patch builds on `PresentationRevision`: an unsaved working copy, mandatory target IDs and original fingerprints, journaled batches of at most 24 operations, and `revert_scribble_changes`. PP01 asks for "a repaired, editable draft of every slide" in the Samsung MD visual language, with a recreated native chart and fixes for stale categories, undersized text, overflow and overlaps. Phase 4 must therefore define which requested repairs are patches and which are explicitly requested object recreation, and confirm that the PP01 grader accepts a draft-copy output. The existing rejections of recomposition that would lose protected artwork, actions, animations, hyperlinks or a non-solid background still apply.

## 7. Implementation sequence and gates

Each phase ends with a reviewable commit/PR and a recorded gate. Avoid full installation/model runs for pure contract or arithmetic changes; use native Office only for checks that require native behavior.

- **Branches.** Create `codex/` feature branches from `origin/codex/development` and target PRs at `codex/development` once PR #20 has landed. Until then, stack on `codex/stress-suite-200` only with the owner's explicit approval.
- **Tool surface.** `scripts/Test-Guardrails.ps1` asserts the exact set of `public const string` tool names in `WorkbookToolCatalog`, `PresentationToolCatalog`, `CrossAppToolCatalog` and `WordToolCatalog`, so any added constant there fails the scan. Prefer typed arguments on existing tools. If a new model-facing tool is justified, update the allow-list in the same PR as a reviewed security change. Never rename existing tool names, the `Scribble Draft` sheet name or the `[Scribble draft]` slide marker.

| Phase | Work and existing integration points | Exit gate |
| --- | --- | --- |
| 0. Freeze evidence and establish baseline | Resolve the PR #20 prerequisite. Archive the gitignored local traces (`qa-385`, `qa-388`, `qa-391`) outside the repository and commit only sanitized derived metrics: suite IDs, assembly hash, `summarize_usage.py` output. Record the exact commit and configuration. Choose the acceptance configuration (hosted or local); probe vision, structured output and context capacity. Restate the existing regressions (stale brief, ranking, ISO dates, native folio) as mechanism-level assertions, and add regressions for context growth (`FakeEndpoint` call and payload counts), geometry-repair routing and Excel write-then-validate ordering (`CrossAppFixture`). Compute the minimum call count of the target stage design. Update checkpoint format. | All seven failure mechanisms in §2 reproducible without paid inference; traces archived; runtime capabilities and baseline metrics recorded, not inferred |
| 1. Canonical analysis contracts | Extend `TaskSources`, `WorkbookGroupedTotals`, and task persistence. Fix capture at `EmailAttachmentReader.ExtractXlsxSheet`, `XlsbTextExtractor` and `WorkbookToolHost` reads, using `LegacyOfficeTextExtractor` as the typed precedent. Add snapshot/fact/dataset/calculation/analysis types and versioning. | Deterministic extraction/calculation, source identity, blank-cell alignment, invalidation, and serialization tests pass |
| 2. First complete path | Add semantic workbook/deck plan adapters beside `WorkbookToolCatalog`, `CrossAppToolCatalog`, `PresentationToolCatalog` and `SamsungWorkflowSchema`, following the tool-surface rule above. Resolve fact references into existing writers. Use isolated new drafts and a development feature flag. | One hand-authored analysis produces native Excel plus four slides without duplicated author-supplied numbers or citations; both outputs match the independent oracle and existing isolation/recovery requirements; the guardrail scan passes with an explicitly reviewed tool list |
| 3. Correct review and repair ownership | Refactor `SamsungAuthoringPolicy`, `SamsungEvidence`, `SamsungPresentationReview` (`ValidateEvidence`, `Numbers`), `SamsungRepairPolicy`, `DocumentDraftHost.PowerPoint`, and `DocumentDraftHost.SlideRepair`. Add typed findings/page metadata and a shared repair budget. Retire the §2 review filters for migrated capabilities. | Wrong AI arithmetic cannot override data, enforced by the finding contract rather than regex filters; contradictory review exits precisely; geometric defects are handled by renderer operations; requests stay within context/call limits |
| 4. Certify common slide layouts and six-slide repair | Reuse `SamsungSlideDesign`, `PresentationDraftWriter.Samsung`, `PresentationRevision`, and `SamsungGenerationJournal`. Use `revise_slides` operations as the patch vocabulary. Add tested layout variants and source-preserving patch plans. Build fixtures on the `SamsungSlideTests` layout previews, `Test-StressNativeGrading.ps1` and `SamsungBenchmarks.json`. | Native editable chart/table/scorecard/text layouts pass approved render fixtures; six-slide repair preserves unrelated content, is accepted by the PP01 grader, and completes on the new contract |
| 5. Generalize safe Excel execution | Extend concepts in `DurableExcelTransform`/`ExcelTransformTarget`; migrate ordinary draft and bounded edit paths in `WorkbookDraftWriter` and `DocumentDraftHost`, binding `write_cells` and `write_draft_sheet` to captured identities instead of the active sheet or workbook. | Formula/address generation, native recalculation, target binding (including an active-window change), failure injection, reconciliation, and supported restoration tests pass |
| 6. Controlled model evaluation | Extend `ModelContractProbe`, request serialization, Test Lab, `summarize_usage.py` and benchmark reporting; change the `TestLabStressBudget` gate only if a local endpoint is the acceptance target. Freeze development and held-out splits. Confirm the live budget before starting. | Target cases and unseen cases pass the acceptance matrix below on a pinned configuration, with measured costs and latency |
| 7. Migration and development-candidate readiness | Complete compatibility routing, session version handling, candidate evidence, user-facing completion status, and fallback rules. Public release stays out of scope: 2.0.91 remains frozen until a new explicit owner release request (`docs/release-channels.md`). | Tested development candidate from CI artifacts; no pending visual review mislabeled as passed; feature-flag rollback and native smoke checks documented |
| 8. Remaining app adapters | Migrate Word report content and Outlook-origin handoffs onto the shared analysis and staged output contracts. Keep mail reading/drafting permissions unchanged. | Excel/PowerPoint/Word agree on values and source revision regardless of the initiating pane; Word native layout/source checks pass; no mail is sent |
| 9. Capability expansion | Certify the remaining layouts, broader Excel operations, source types, and the existing 200-case corpus in budgeted batches. The full corpus does not fit the current $30 cap and needs a separate budget authorization or a local runtime. | A published internal capability matrix distinguishes supported, tested, and unsupported combinations; broader reliability claims have matching evidence |

Phases 3 and layout fixture work can proceed alongside phase 2 after the data contracts are agreed. Pure Excel transaction work can proceed separately after target/operation IDs are agreed. Only one process should own live Office test sessions at a time. Implementation agents use isolated branches/worktrees; do not run two writers against the same Office documents.

Phase 2 is limited to new, marked, disposable draft destinations using existing proven staging/journaling. It does not modify source workbooks or enable ordinary in-place edits. If a live-formula pilot needs execution machinery that does not yet exist, implement that minimum dependency first. Generalized bounded edits cannot become release-supported before phase 5 passes.

### Milestone decision

After phases 0–3, run a small architectural-feasibility pilot. Complete the hand-authored shared-output experiment before expanding schemas broadly. Continue expanding only if correctness and call/context costs improve. If the simplified model-facing contract still fails, isolate interpretation, planning, serialization, and execution in separate experiments before selecting another model or changing the architecture again. A successful early pilot supports continued implementation; visual certification, general recovery claims, and deployment still require phases 4–7.

Do not promise an elapsed completion time before this milestone. The critical path is demonstrable native behavior, not the number of commits produced.

## 8. Tests and acceptance criteria

### Offline deterministic tests

Use independent expected results, not values generated by the same production aggregation function. Test changed amounts/labels/order as well as the original fixture. Cover:

- Typed dates, identifiers containing numbers, negative quantities, percentages versus fractions, currency/unit mismatch, precision, and zero denominators.
- Mixed years, absent periods, missing versus zero, duplicate records, filtered ranges, reordered columns, duplicate sheet names across workbooks, and unsaved documents.
- Blank cells in attached `.xlsx`/`.xlsb` sheets keep their column; formulas, number formats and date serials are captured as typed values.
- Correct arithmetic paired with a stale/wrong model brief; correct ranking paired with an incorrect AI objection.
- Snapshot changes, source conflicts, incomplete capture, and unsupported extraction.
- Formula dependency cycles, relative/absolute references, locale-independent formula emission, recalculation errors, and requested live links.
- Unsupported prose does not become verified merely by referring to a valid fact ID.

### Renderer fixtures

Begin with six common families: scorecard, comparison chart/table, grouped-data chart, evidence cards, dense table, and cover/closing. They map to existing layout IDs in `SamsungSlideDesign.Layouts`: `scorecard`; `chart`, `two_pane` or `table`; `chart` or `annotated_chart`; `cards`; `table` or `matrix`; `cover` and `closing`. Do not add layout IDs unless a family needs one. Exercise short/long titles, varied digit widths, negative/zero/missing values, long labels, and realistic density boundaries. Preserve other layouts behind their existing route until certified.

Start with three density variants per family: normal, long labels/titles, and maximum supported content, yielding 18 reference fixtures. Calibrate the reviewer on these plus at least 18 deliberately defective variants. Include false arithmetic objections and wrong expected-page metadata as reviewer-contract tests, not aesthetic judgments.

Freeze calibration thresholds before running it: zero acceptance of seeded critical defects, zero permitted arithmetic/native-identity overrides, at least 90% recall of noncritical visual blockers, and no more than 10% false blockers on approved fixtures. Report each category separately. These small-set thresholds qualify the reviewer for the pilot, not universal visual judgment. Until calibration passes, an identified human must review the exact candidate outputs. Pilot release also retains that independent visual attestation; automated review alone does not replace it.

Create native rendered reference outputs using deliberately designed specifications. An identified reviewer approves the baseline against the user's design direction. Store both native geometry/readback and images. Use tolerances for font/Office-version rendering differences; pixel equality alone is not the acceptance criterion. Never automatically bless a new screenshot after a regression.

A visual pass requires readable text, no unintended overlap or clipping, correct emphasis and chart annotations, balanced composition, consistent typography, and useful content retained. Automated checks plus an AI “approved” verdict are insufficient until the reviewer's quality has itself been evaluated against known good/bad examples.

### Recovery tests

Inject failure before mutation, after partial native writes, after native success but before receipt persistence, during readback, and during resume. Change the active sheet/window and edit a target concurrently. Verify no duplicate writes, no wrong-document edits, preserved sources, explicit uncertainty where applicable, and restoration only when it can be verified.

### Live evaluation ladder

Before selecting held-out tasks, freeze the supported capability matrix: Excel operations/formula families, protected or merged ranges, spill arrays, external links/UDFs, presentation objects, layout density, source types, and runtime configuration. The positive held-out set exercises declared-supported capabilities. `Unsupported` on such a case is a non-completion. Expected rejections for unsupported inputs belong to a separate negative suite. Do not narrow support after observing failures and retain the old pass-rate claim.

1. Use a small development set with the existing XA01/PP01 cases and varied Excel/deck inputs.
2. Require three consecutive unassisted passes each for XA01 and PP01 on a fixed candidate/configuration, resetting the count on failure, as `tests/benchmarks/stress/CHECKPOINT.md` already defines. This is a smoke/repeatability gate, not proof of broad reliability.
3. Hold out 12 distinct tasks, four each for Excel, PowerPoint, and cross-app work. Include new amounts, schemas, periods, long labels, missing data, and user edits. Author them fresh rather than generating them from the `build_catalog.py` templates that development has iterated on. Seal them with a manifest SHA-256 (the existing `--kit-sha256` mechanism) and keep prompts and oracles outside the implementation branch until the evaluation run. Until phase 8, cross-app held-out tasks are Excel↔PowerPoint handoffs only, because Word and Outlook paths are not yet on the new contract. Run each twice: 24 held-out executions. Keep expected answers outside model context.
4. Release gate: all primary smoke runs pass; at least 23/24 held-out runs complete correctly and meet visual criteria; zero wrong-source writes, silent corruption, false success, or incorrect numeric outputs presented as verified. Any non-completion must be safe, explicit, and diagnosed. Do not tune on the held-out set and continue calling it held-out—replace exposed cases.
5. Run the existing broader Office/Outlook/browser regression suites. Claim only the capabilities and model/runtime configurations actually evaluated. Expand the 200-case live suite in later batches when the pilot and budget support it.

**Budget fit.** Checkpoint deltas imply roughly $0.35–0.50 per XA01 or PP01 run at current token volumes. The six smoke runs plus 24 held-out runs would cost about $10–15, against about $5.62 remaining under the $30 hard cap (17:51 UTC, `b873435`). The ladder therefore fits only after the efficiency targets below are demonstrated, with a new explicit budget authorization, or on a local runtime. Record that decision before phase 6.

### Proposed efficiency targets for the bounded pilot

These are initial engineering gates, not measured achievements or promises for arbitrary documents. The counts include chat-loop turns as well as internal reviews. Phase 0 computes the call floor of the proposed stage design (for example, per-slide rendered review versus a montage review through `SamsungDeckOverview.Montage`); if the floor already exceeds a target, revise the target before implementation, not after measuring.

| Metric | XA01: Excel + four slides | PP01: six-slide repair |
| --- | --- | --- |
| Total model requests, including internal reviewers/retries | ≤12 | ≤18 |
| Aggregate prompt tokens, counting repeated/cached inputs | ≤150,000 | ≤250,000 |
| Individual review input | ≤12,000 tokens, plus explicitly bounded images | ≤12,000 tokens, plus explicitly bounded images |
| Full-slide regeneration for a wording/geometry defect | 0 | 0 |
| Silent data loss or false completion | 0 | 0 |

Measure actual endpoint token usage and image accounting. If a review exceeds its budget, select a smaller relevant evidence slice or record an explicit budget failure; never silently truncate necessary facts. Compare latency on the same fixed runtime/hardware, report median and p95, and aim for at least 50% median reduction against a reproducible baseline. If a baseline cannot be reproduced, report absolute timings without claiming the percentage improvement.

### Test commands and evidence

Use the repository's current scripts rather than inventing a parallel test harness. Relevant entry points are `scripts/Test-Guardrails.ps1`, `tests/GuardrailTests` (`GuardrailTests.exe`; register tests in `Main`), `tests/NativeAcceptance/Test-PowerPointWorkflow.ps1`, `tests/NativeAcceptance/Test-OfficeReadiness.ps1`, `tests/benchmarks/evaluator/evaluate.py`, `tests/benchmarks/stress/Test-StressNativeGrading.ps1`, `tests/benchmarks/stress/Run-StressSuite.ps1`, `tests/benchmarks/stress/summarize_usage.py`, and the Test Lab runner described in `tests/benchmarks/stress/README.md`. Inspect current parameters at implementation time.

CI (`.github/workflows/build.yml`, `windows-latest`, no Office) builds the solution with MSBuild and runs the browser-extension checks, Python evaluator and catalog tests, Test Lab script tests, `GuardrailTests.exe`, `Test-Guardrails.ps1`, and the release-evidence and public-release-freeze gates. Native Office checks, including stress runs, run only on the Windows workstation.

Report deterministic status, native execution status, visual status, source-preservation status, and overall acceptance separately. Preserve an auditable final attestation that binds manual visual review to the exact output hashes, reviewer, and rubric version. Do not leave a JSON result saying pending while a prose checkpoint says fully passed.

## 9. Migration, compatibility, and release

- Version through the existing mechanism. `DurableTaskState.SamsungWorkflowVersion` (currently 2, from `SamsungAuthoringPolicy.WorkflowVersion`) already routes tasks started before workflow 2 to the frozen `LegacySamsung` path, and `SamsungGenerationJournal.CanResume` checks it. Persist the new analysis/document-plan contract version beside it (provisional names such as `analysis_v1`/`document_plan_v1`), route by the persisted value, and gate new tasks behind a development feature flag.
- Keep legacy read tools and existing writers. Add adapters that resolve typed plans into existing native structures; switch validated tasks gradually.
- Never reinterpret an in-progress legacy journal as a new-schema task. Resume through its original path or create a new task after reconciling the draft.
- Keep feature flags and fallback decisions explicit in test evidence. Never silently retry a failed new-path write through the old writer.
- Source refresh invalidates affected facts and plans. A user edit creates a new expected state; stale plans cannot overwrite it.
- Pilot on development builds from CI artifacts. Public Scribble stays frozen at 2.0.91: only a new explicit owner release request can change GitHub Latest or `continuous`, and CI has no release job.
- Roll back by disabling the new path for new tasks; reconcile in-flight operations with the executor that wrote them. Retain evidence and existing drafts.
- Once the new path covers a capability and passes its gates, remove obsolete prompt exceptions, redundant reviewer loops and the matching §2 review filters for that capability. Preserve regression coverage while removing the workaround.

## 10. Budget and progress discipline

No paid inference is needed to write this plan or complete most offline contract/renderer work. Before a live batch, inspect the current authorized cap and remaining balance, estimate the batch, and keep the existing reserve. The stress harness enforces a $30 total cap on the dedicated no-reset OpenRouter key (`TestLabStressBudget.MaximumTotalUsd`), checks the key before each model request and keeps a $0.05 reserve. About $5.62 remained at 17:51 UTC on 22 September. At current volumes that covers roughly 11–16 XA01/PP01 runs, fewer than the live ladder needs (§8). This plan does not raise the cap or authorize a different paid provider. Run as much authorized validation as fits; if funding prevents the remaining required checks, report the exact incomplete gate rather than claiming success.

Every experimental batch records: hypothesis, code commit, configuration fingerprint, test IDs, expected change, actual outputs, defect class, calls/tokens/cost/time, and next decision. After two unsuccessful attempts at the same architectural hypothesis, stop that experiment and isolate the failing component. Do not continue the same full workflow with one more prose prohibition.

Progress is measured by defects eliminated, unseen cases completed, reduced execution cost, and verified recovery—not commit count, build version, or accumulated guardrails.

## 11. Risks and decision points

| Risk | Mitigation / decision |
| --- | --- |
| Typed registry gives false confidence in incorrectly mapped data | Retain source mappings; test semantic associations independently; expose ambiguity rather than certifying extraction automatically |
| Smaller schemas restrict useful authoring | Add composable blocks and supported expression operations as demand is demonstrated; keep unsupported requests explicit |
| Layout library still looks poor | Design and approve renderer fixtures before more model tuning; renderer owns layout improvements |
| Native Office differs by version/fonts/locale | Record environment; use native readbacks plus tolerant visual comparisons; validate intended deployment environment |
| Structured-output support differs by runtime | Capability probe and bounded fallback; a tool schema in a prompt is not proof of enforced decoding |
| Open workbook state changes mid-task | Relevant fingerprints and bound targets; explicit conflict handling; no blind replay |
| Refactor competes with active patch work | Reconcile baseline; allocate file ownership; use isolated implementation worktrees and one native test owner |
| Good pilot results fail to generalize | Held-out tasks, failure taxonomy, broader suite, and separate claims per tested model/runtime |
| Plan depends on unmerged draft PR #20 | Land PR #20 in `codex/development` first, or get explicit owner approval to stack; reconcile before each phase |
| Remaining live budget cannot cover the evaluation ladder | Prove efficiency offline and on smoke runs first; then request explicit authorization or evaluate on a local runtime; report incomplete gates |
| Hosted OpenRouter evaluation does not represent the local deployment | Choose the acceptance configuration in phase 0; make claims per evaluated configuration |
| Guardrail scan rejects new model-facing tools | Prefer typed arguments on existing tools; update `scripts/Test-Guardrails.ps1` deliberately, with review, in the same PR |

## 12. Coding-agent handoff

> Implement the Scribble reliability plan incrementally. First read `CLAUDE.md` and `docs/release-channels.md`. Confirm that PR #20 (`codex/stress-suite-200`) has landed in `codex/development`, or that the owner approved stacking on it. Reconcile the current branch against planning baseline `fad5cdd859713561af704e08674936d719b2d5ca` and the later commits (`ebaad7b`, `b873435` at review time). Preserve unrelated work and coordinate with the existing implementation task before editing overlapping files or operating Office. Do not modify `src/Scribble/Office/LegacySamsung/`.
>
> Start with phase 0 and phase 1. Archive the local traces. Restate the existing stale-brief, ranking, ISO-date and native-folio regressions as mechanism-level assertions, and add offline regressions for context growth, geometry-repair routing and Excel write-then-validate ordering. Extend existing source/provenance and calculation infrastructure into immutable typed snapshots, facts, tables, and analysis artifacts, starting at the workbook extractors and COM reads. Keep user requirements and narrative briefs separate from factual authority.
>
> Prove one analysis can produce a live-formula Excel report and an editable four-slide deck through the existing writers. The model must reference facts rather than retype values and citations. AI reviewers cannot overrule verified arithmetic. Route every repair to the layer capable of performing it, and use targeted patches under a shared repair budget. Change model-facing tool names only with a deliberate, reviewed update to `scripts/Test-Guardrails.ps1`.
>
> Preserve source documents and existing unsaved work. Use bound, staged, receipted execution; do not promise atomic COM transactions. Keep the old and new paths versioned during migration through the existing workflow-version mechanism. Do not expand to all apps or all benchmark cases before the architectural pilot passes.
>
> Complete offline and native component checks before paid end-to-end runs. Use the user's existing authorized budget and a pinned configuration; check that the planned batch fits the remaining balance first. Report exact evidence and incomplete gates. Do not report completion based on CI, model self-assessment, or a pending visual evaluation. The plan's acceptance matrix defines completion.

## References

- Repository and native traces at the baseline identified above; file names in this plan are relative to that repository. The traces are local to the workstation until phase 0 archives them.
- [Anthropic: Writing effective tools for agents](https://www.anthropic.com/engineering/writing-tools-for-agents): reducing intermediate context and consolidating useful operations.
- [Anthropic: Building effective agents](https://www.anthropic.com/engineering/building-effective-agents): choosing simple workflows and using evaluator loops only when feedback produces measurable improvement.
- [vLLM: Tool calling and strict mode](https://docs.vllm.ai/en/latest/features/tool_calling/#strict-mode): runtime-dependent schema enforcement. This is a capability to verify on the chosen server, not evidence that Scribble currently enables it.

The implementation phases, interfaces, targets, and rollout decisions above are proposed engineering decisions grounded in the inspected code. They have not yet been implemented or validated.
