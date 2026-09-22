# Scribble reliability: architecture and implementation plan

Prepared 22 September 2026. Planning baseline: `fad5cdd859713561af704e08674936d719b2d5ca` on `codex/stress-suite-200`, repository `C:\Users\keeoh\Documents\ChatGPT\Scribble`.

The existing implementation task is active. This is a design and coding-agent handoff, not an instruction to overwrite its current work. Reconcile changes made after this baseline before implementation. Creating this plan does not change Scribble, stop another task, or authorize additional API spending.

## 1. Outcome and engineering decision

Make routine Office work reliable with the configured local model: create an accurate, editable Excel analysis; use the same analysis in a well-designed PowerPoint deck; and make bounded edits without corrupting source work or entering repetitive review loops.

The core change is to make Scribble own facts, calculations, document identity, layout mechanics, and execution. The model owns interpretation, narrative, and choices among supported operations. Extend the existing implementation rather than replacing the Office integration or building another renderer.

The first proof is **one verified analysis → Excel report and four-slide deck**, followed by **repairing a six-slide deck using an authoritative workbook**. These are development targets, not the boundaries of the eventual product. Generalization must be demonstrated with unseen inputs.

Success means correct native outputs, preserved sources, acceptable rendered quality, bounded execution, and honest completion status. A new build, a model saying “done,” or a single lucky run does not establish success.

## 2. Evidence and root causes

These observations were verified in source code and native-run traces. The plan addresses mechanisms, not individual benchmark values.

| Observed failure | Architectural cause | Required correction |
| --- | --- | --- |
| Reviewer demands 55.74%; host calculation requires 55.76% | Model-authored brief competes with verified data as an authority | Separate requirements, narrative plans, and facts; calculations have a single authoritative result |
| Reviewer says 22,044 exceeds 22,675 and rejects the correct ranking | Probabilistic reviewer can overrule deterministic comparisons | Rankings and arithmetic belong to code; review objections must identify a checkable source or binding error |
| Date suffixes such as `-10` and `-16` become unsupported numbers | Numeric scanning cannot distinguish dates from quantities | Typed dates, identifiers, quantities, and text segments |
| Slide six is reviewed against footer one | Native identity, logical identity, and display order are conflated | Explicit slide/page identities and typed page-number metadata |
| Model is asked to fix fixed-position titles and chart-highlight boxes | Repair authority does not match defect ownership | Geometry repairs belong to the renderer; model repairs change supported semantic fields |
| Six-slide run makes 87 model calls and processes 3,387,583 prompt tokens, including repeated/cached context | Duplicated evidence, nested reviews, large payloads, and regeneration | Scoped context, fact references, bounded stage calls, targeted patches |
| Ordinary Excel write can discover formula errors after some cells were written | Validation and mutation are interleaved | Stage, validate, commit, read back, and reconcile against an explicit document target |

Trace baselines: `artifacts/qa-388-pp01-failed/`, `artifacts/qa-385-xa01-run1/`, and `artifacts/qa-391-pp01-blocked/`. The PP01 v388 run used 13 hosted providers for the same Qwen model. Some author/reviewer calls disabled reasoning. This is not a controlled benchmark of the best local runtime configuration.

### Preserve these existing assets

- `TaskSources`: captured source text and provenance spans.
- `WorkbookGroupedTotals`: deterministic decimal aggregation.
- `SamsungSlideDesign` and `PresentationDraftWriter.Samsung`: native editable layouts and rendering.
- `SamsungGenerationJournal`: write receipts, slide identities, fingerprints, and resume checks.
- `DurableExcelTransform`: staged capture/write/readback machinery for specialized transformations.
- Native artifact readbacks, stress fixtures, Test Lab, and existing source-preservation checks.

The problem is how these components exchange information and authority. Their presence makes an incremental migration practical.

One concrete generalization issue belongs in the first data phase: `TaskSources.CompleteWorkbookTotals` recognizes specific `RowID/Period/Group/RevenueEUR/CostEUR` headers, while generic read capture flattens string leaves. Replace that special case with typed table capture and reusable aggregate operations. Do not add more benchmark-specific header recognizers. Preserve distinct provenance when two documents contain identical text.

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

Names below are proposed types; adapt to repository conventions. Keep the current .NET Framework 4.8/C# 7.3 compatibility unless a separate change is justified.

### SourceSnapshot

Store `snapshot_id`, source instance identity, source type, capture revision, capture time, content hash, supported coverage, calculation state, and locators. Excel locators identify workbook instance, worksheet identity, range, formula/value, raw cell type, number format, cached value, and relevant table headers. PowerPoint locators distinguish presentation identity, stable SlideID, shape ID, and displayed ordinal. A stale formula cache is not a verified recalculated value.

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

`WorkbookPlan` uses semantic columns, rows, formulas, and charts. The compiler allocates cell addresses and generates native formulas. Supported formula operations must cover the actual pilot, including aggregate rates and references to source data. Advanced formulas may use an explicit validated expression path; unsupported operations return a capability error rather than silently becoming values.

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

Patches operate on a specific field or block. Do not regenerate immutable data and citations when shortening a title. Page metadata must include an integer `expected_page_number`, separate from native slide ID and logical slide ID.

## 6. Runtime, review, and recovery

### Model interaction

- Expose only the tools and schema needed at the current stage.
- Return relevant typed facts and summaries; fetch additional evidence by ID when needed.
- Keep full evidence in the existing protected evidence store, not repeatedly inside prompts.
- Support capability-probed strict structured output where the selected server supports it. Keep a bounded parse/validation fallback for servers that do not. JSON validity does not establish factual correctness.
- Record model, provider/runtime build, tokenizer/chat template, quantization, context limit, reasoning settings, sampling, and output limits for comparisons.
- For local evaluation, pin one configuration. For hosted diagnostics, pin a supported provider where possible. Do not mix configurations and attribute the difference solely to code.
- Prompt corrections must describe one structured defect. Avoid growing the universal system prompt after each failure.

Discover the actual local endpoint, model, hardware, vision support, context capacity, and schema-enforcement behavior in phase 0. The existing PowerPoint path requires vision; if the chosen local model lacks it, the architectural pilot uses native assertions plus an identified human visual reviewer. Any separate local vision model must be an explicit, measured configuration. Do not silently route visual review to a hosted service. Fully automated visual review remains unproven until its configured path passes calibration.

### Review boundaries

Deterministic checks own numeric values, dates, periods, totals, rankings, required counts, data bindings, object identity, and source-preservation comparisons. Model review handles narrative and visual judgments, using only relevant evidence and rendered pages.

A reviewer may flag a wrong label-to-column mapping or unsupported conclusion. It must identify the affected binding; it cannot replace a verified value from memory. Aesthetic observations route to supported layout alternatives or renderer operations. Unresolvable defects remain explicit failed criteria; the system must not conceal them by deleting requested content.

Proposed global repair policy: at most one normal corrective model patch per affected block, with an explicit task-level repair budget. Schema retries, page reviews, and final review all consume that budget. There is no nested reset that grants three more cycles inside another three-cycle loop. Renderer fitting can try a finite ordered list of valid variants without an LLM call.

### Native execution

Use a durable state machine:

`Captured → Planned → Validated → Staged → Applying → ReadbackVerified → Complete`

Exceptional outcomes: `Conflict`, `Unsupported`, `NeedsInspection`, or `Failed`. “Applying” after a crash is uncertain until reconciled; it is not evidence that an operation succeeded or failed.

Each operation records an operation ID, target binding, expected pre-state, intended patch, result fingerprint, and verification receipt. Before mutation, confirm the relevant source, destination, and formula-dependency fingerprints. Retain before-images of every touched property needed for supported recovery: formulas, values, formats, dimensions, and owned objects as applicable. Formula staging must preserve workbook reference semantics; an isolated scratch cell is insufficient for formulas that depend on sheets, names, tables, or workbook settings.

COM does not provide a general atomic transaction. Use staged preparation and compensating restoration where proven safe. On interruption, compare actual state with before/after expectations; resume or acknowledge a verified operation once. If concurrent user edits prevent safe restoration, preserve the document and explain the exact conflict. Do not claim rollback succeeded without readback evidence.

Declare capability handling for merged/protected ranges, spill arrays, external links, volatile functions, and UDFs. Preserve or reject unsupported cases explicitly; do not pretend that an arbitrary formula has been validated. Keep “applied in memory,” “verified,” and “saved” as distinct states.

For six-slide repair, work in a marked draft copy of the original deck, preserve unchanged notes/artwork/hyperlinks, and patch only requested defects. Unsupported objects remain preserved or are explicitly identified as unsupported. Do not approximate them silently while reconstructing a deck.

## 7. Implementation sequence and gates

Each phase ends with a reviewable commit/PR and a recorded gate. Avoid full installation/model runs for pure contract or arithmetic changes; use native Office only for checks that require native behavior.

| Phase | Work and existing integration points | Exit gate |
| --- | --- | --- |
| 0. Freeze evidence and establish baseline | Archive the named traces and exact commit/configuration. Discover local runtime/hardware/capabilities. Add sanitized offline regressions for contradictory percentages/rankings, date parsing, wrong page numbers, and review context growth. Update checkpoint format. | All five failure mechanisms reproducible without paid inference; runtime capabilities and baseline metrics recorded, not inferred |
| 1. Canonical analysis contracts | Extend `TaskSources`, `WorkbookGroupedTotals`, and task persistence. Add snapshot/fact/dataset/calculation/analysis types and versioning. | Deterministic extraction/calculation, source identity, invalidation, and serialization tests pass |
| 2. First complete path | Add semantic workbook/deck plan adapters beside `WorkbookToolCatalog`, `CrossAppToolCatalog`, and `SamsungWorkflowSchema`. Resolve fact references into existing writers. Use isolated new drafts and a development feature flag. | One hand-authored analysis produces native Excel plus four slides without duplicated author-supplied numbers or citations; both outputs match the independent oracle and existing isolation/recovery requirements |
| 3. Correct review and repair ownership | Refactor `SamsungAuthoringPolicy`, `SamsungEvidence`, `DocumentDraftHost.PowerPoint`, and `DocumentDraftHost.SlideRepair`. Add typed findings/page metadata and a shared repair budget. | Wrong AI arithmetic cannot override data; contradictory review exits precisely; geometric defects are handled by renderer operations; requests stay within context/call limits |
| 4. Certify common slide layouts and six-slide repair | Reuse `SamsungSlideDesign`, `PresentationDraftWriter.Samsung`, `PresentationRevision`, and `SamsungGenerationJournal`. Add tested layout variants and source-preserving patch plans. | Native editable chart/table/scorecard/text layouts pass approved render fixtures; six-slide repair preserves unrelated content and completes on the new contract |
| 5. Generalize safe Excel execution | Extend concepts in `DurableExcelTransform`/`ExcelTransformTarget`; migrate ordinary draft and bounded edit paths in `WorkbookDraftWriter` and `DocumentDraftHost`. | Formula/address generation, native recalculation, target binding, failure injection, reconciliation, and supported restoration tests pass |
| 6. Controlled model evaluation | Extend `ModelContractProbe`, request serialization, Test Lab, and benchmark reporting. Freeze development and held-out splits. | Target cases and unseen cases pass the acceptance matrix below on a pinned configuration, with measured costs and latency |
| 7. Migration and release | Complete compatibility routing, session version handling, release evidence, user-facing completion status, and fallback rules. | Tested release candidate; no pending visual review mislabeled as passed; release rollback and native smoke checks documented |
| 8. Remaining app adapters | Migrate Word report content and Outlook-origin handoffs onto the shared analysis and staged output contracts. Keep mail reading/drafting permissions unchanged. | Excel/PowerPoint/Word agree on values and source revision regardless of the initiating pane; Word native layout/source checks pass; no mail is sent |
| 9. Capability expansion | Certify the remaining layouts, broader Excel operations, source types, and the existing 200-case corpus in budgeted batches. | A published internal capability matrix distinguishes supported, tested, and unsupported combinations; broader reliability claims have matching evidence |

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
- Correct arithmetic paired with a stale/wrong model brief; correct ranking paired with an incorrect AI objection.
- Snapshot changes, source conflicts, incomplete capture, and unsupported extraction.
- Formula dependency cycles, relative/absolute references, locale-independent formula emission, recalculation errors, and requested live links.
- Unsupported prose does not become verified merely by referring to a valid fact ID.

### Renderer fixtures

Begin with six common families: scorecard, comparison chart/table, grouped-data chart, evidence cards, dense table, and cover/closing. Exercise short/long titles, varied digit widths, negative/zero/missing values, long labels, and realistic density boundaries. Preserve other layouts behind their existing route until certified.

Start with three density variants per family: normal, long labels/titles, and maximum supported content, yielding 18 reference fixtures. Calibrate the reviewer on these plus at least 18 deliberately defective variants. Include false arithmetic objections and wrong expected-page metadata as reviewer-contract tests, not aesthetic judgments.

Freeze calibration thresholds before running it: zero acceptance of seeded critical defects, zero permitted arithmetic/native-identity overrides, at least 90% recall of noncritical visual blockers, and no more than 10% false blockers on approved fixtures. Report each category separately. These small-set thresholds qualify the reviewer for the pilot, not universal visual judgment. Until calibration passes, an identified human must review the exact candidate outputs. Pilot release also retains that independent visual attestation; automated review alone does not replace it.

Create native rendered reference outputs using deliberately designed specifications. An identified reviewer approves the baseline against the user's design direction. Store both native geometry/readback and images. Use tolerances for font/Office-version rendering differences; pixel equality alone is not the acceptance criterion. Never automatically bless a new screenshot after a regression.

A visual pass requires readable text, no unintended overlap or clipping, correct emphasis and chart annotations, balanced composition, consistent typography, and useful content retained. Automated checks plus an AI “approved” verdict are insufficient until the reviewer's quality has itself been evaluated against known good/bad examples.

### Recovery tests

Inject failure before mutation, after partial native writes, after native success but before receipt persistence, during readback, and during resume. Change the active sheet/window and edit a target concurrently. Verify no duplicate writes, no wrong-document edits, preserved sources, explicit uncertainty where applicable, and restoration only when it can be verified.

### Live evaluation ladder

Before selecting held-out tasks, freeze the supported capability matrix: Excel operations/formula families, protected or merged ranges, spill arrays, external links/UDFs, presentation objects, layout density, source types, and runtime configuration. The positive held-out set exercises declared-supported capabilities. `Unsupported` on such a case is a non-completion. Expected rejections for unsupported inputs belong to a separate negative suite. Do not narrow support after observing failures and retain the old pass-rate claim.

1. Use a small development set with the existing XA01/PP01 cases and varied Excel/deck inputs.
2. Run XA01 and PP01 three times each on a fixed candidate/configuration. This is a smoke/repeatability gate, not proof of broad reliability.
3. Hold out 12 distinct tasks, four each for Excel, PowerPoint, and cross-app work. Include new amounts, schemas, periods, long labels, missing data, and user edits. Run each twice: 24 held-out executions. Keep expected answers outside model context.
4. Release gate: all primary smoke runs pass; at least 23/24 held-out runs complete correctly and meet visual criteria; zero wrong-source writes, silent corruption, false success, or incorrect numeric outputs presented as verified. Any non-completion must be safe, explicit, and diagnosed. Do not tune on the held-out set and continue calling it held-out—replace exposed cases.
5. Run the existing broader Office/Outlook/browser regression suites. Claim only the capabilities and model/runtime configurations actually evaluated. Expand the 200-case live suite in later batches when the pilot and budget support it.

### Proposed efficiency targets for the bounded pilot

These are initial engineering gates, not measured achievements or promises for arbitrary documents.

| Metric | XA01: Excel + four slides | PP01: six-slide repair |
| --- | --- | --- |
| Total model requests, including internal reviewers/retries | ≤12 | ≤18 |
| Aggregate prompt tokens, counting repeated/cached inputs | ≤150,000 | ≤250,000 |
| Individual review input | ≤12,000 tokens, plus explicitly bounded images | ≤12,000 tokens, plus explicitly bounded images |
| Full-slide regeneration for a wording/geometry defect | 0 | 0 |
| Silent data loss or false completion | 0 | 0 |

Measure actual endpoint token usage and image accounting. If a review exceeds its budget, select a smaller relevant evidence slice or record an explicit budget failure; never silently truncate necessary facts. Compare latency on the same fixed runtime/hardware, report median and p95, and aim for at least 50% median reduction against a reproducible baseline. If a baseline cannot be reproduced, report absolute timings without claiming the percentage improvement.

### Test commands and evidence

Use the repository's current scripts rather than inventing a parallel test harness. Relevant entry points are `scripts/Test-Guardrails.ps1`, `tests/NativeAcceptance/Test-PowerPointWorkflow.ps1`, `tests/NativeAcceptance/Test-OfficeReadiness.ps1`, `tests/benchmarks/evaluator/evaluate.py`, and the Test Lab runner described in `tests/benchmarks/stress/README.md`. Inspect current parameters at implementation time. The CI build uses MSBuild for the solution and separate browser-extension checks.

Report deterministic status, native execution status, visual status, source-preservation status, and overall acceptance separately. Preserve an auditable final attestation that binds manual visual review to the exact output hashes, reviewer, and rubric version. Do not leave a JSON result saying pending while a prose checkpoint says fully passed.

## 9. Migration, compatibility, and release

- Introduce `analysis_v1`/`document_plan_v1` schema versions and a development feature flag. The names are provisional.
- Keep legacy read tools and existing writers. Add adapters that resolve typed plans into existing native structures; switch validated tasks gradually.
- Never reinterpret an in-progress legacy journal as a new-schema task. Resume through its original path or create a new task after reconciling the draft.
- Keep feature flags and fallback decisions explicit in test evidence. Never silently retry a failed new-path write through the old writer.
- Source refresh invalidates affected facts and plans. A user edit creates a new expected state; stale plans cannot overwrite it.
- Pilot on a development candidate. A successful build does not update the public stable release automatically.
- Roll back by disabling the new path for new tasks; reconcile in-flight operations with the executor that wrote them. Retain evidence and existing drafts.
- Once the new path covers a capability and passes its gates, remove obsolete prompt exceptions and redundant reviewer loops for that capability. Preserve regression coverage while removing the workaround.

## 10. Budget and progress discipline

No paid inference is needed to write this plan or complete most offline contract/renderer work. Before a live batch, inspect the current authorized cap and remaining balance, estimate the batch, and keep the existing reserve. This plan does not raise the cap or authorize a different paid provider. Run as much authorized validation as fits; if funding prevents the remaining required checks, report the exact incomplete gate rather than claiming success.

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

## 12. Coding-agent handoff

> Implement the Scribble reliability plan incrementally. First read repository instructions and reconcile the current branch against planning baseline `fad5cdd859713561af704e08674936d719b2d5ca`. Preserve unrelated work and coordinate with the existing implementation task before editing overlapping files or operating Office.
>
> Start with phase 0 and phase 1. Capture the known contradictory-review, date-token, native-numbering, and context-growth failures as offline regressions. Extend existing source/provenance and calculation infrastructure into immutable typed snapshots, facts, tables, and analysis artifacts. Keep user requirements and narrative briefs separate from factual authority.
>
> Prove one analysis can produce a live-formula Excel report and an editable four-slide deck through the existing writers. The model must reference facts rather than retype values and citations. AI reviewers cannot overrule verified arithmetic. Route every repair to the layer capable of performing it, and use targeted patches under a shared repair budget.
>
> Preserve source documents and existing unsaved work. Use bound, staged, receipted execution; do not promise atomic COM transactions. Keep the old and new paths versioned during migration. Do not expand to all apps or all benchmark cases before the architectural pilot passes.
>
> Complete offline and native component checks before paid end-to-end runs. Use the user's existing authorized budget and a pinned configuration. Report exact evidence and incomplete gates. Do not report completion based on CI, model self-assessment, or a pending visual evaluation. The plan's acceptance matrix defines completion.

## References

- Repository and native traces at the baseline identified above; file names in this plan are relative to that repository.
- [Anthropic: Writing effective tools for agents](https://www.anthropic.com/engineering/writing-tools-for-agents): reducing intermediate context and consolidating useful operations.
- [Anthropic: Building effective agents](https://www.anthropic.com/engineering/building-effective-agents): choosing simple workflows and using evaluator loops only when feedback produces measurable improvement.
- [vLLM: Tool calling and strict mode](https://docs.vllm.ai/en/latest/features/tool_calling/#strict-mode): runtime-dependent schema enforcement. This is a capability to verify on the chosen server, not evidence that Scribble currently enables it.

The implementation phases, interfaces, targets, and rollout decisions above are proposed engineering decisions grounded in the inspected code. They have not yet been implemented or validated.
