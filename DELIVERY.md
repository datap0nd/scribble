# Scribble 2.0 — master delivery plan

Updated 26 September 2026. [PR #41](https://github.com/datap0nd/scribble/pull/41)
merged this plan into `codex/development` on 25 September 2026. This remains
the authoritative implementation and acceptance plan for the owner's
instruction to redirect Scribble and finish delivery. Execution progress is
tracked in [DELIVERY_STATUS.md](DELIVERY_STATUS.md).

Navigation: [authority and completion](#1-authority-integration-and-definition-of-completion)
· [current evidence](#3-verified-starting-position)
· [scope](#4-supported-scope-and-explicit-boundaries)
· [architecture](#5-architecture-and-ownership-of-decisions)
· [work packages](#6-ordered-implementation-work-packages)
· [acceptance](#7-acceptance-stop-rules-and-evidence-format)
· [runtime and budget](#8-runtime-inference-budget-and-efficiency-gates)
· [commands](#9-reproducible-execution-entry-points)
· [risks](#10-risks-and-decision-rules)
· [handoff](#11-operating-rules-and-immediate-handoff).

## 1. Authority, integration, and definition of completion

Continue from `codex/development` through reviewable follow-up changes. PR #41
merged the integration plan, and PRs #20–#40 were consolidated and closed with
their source branches preserved. Do not reopen the phase stack or create a
competing integration line. Preserve unrelated local changes. One
implementation owner coordinates changes and one test runner owns disposable
Office sessions at a time.

This document supersedes the execution order, branch prerequisites, dated
budget figures, ownership instructions, and acceptance-policy conflicts in
[RELIABILITY_PLAN.md](RELIABILITY_PLAN.md),
[RELIABILITY_PLAN_REVIEW.md](RELIABILITY_PLAN_REVIEW.md), and historical phase
checkpoints. Those documents remain design rationale and dated evidence.
Their statements about what is currently implemented are not a current status
report. The detailed contract rationale remains useful; the decisions and
measurable gates below govern new work.

The product outcome is accurate, useful, editable Office output that preserves
source work, finishes predictably, and reports exactly what was verified. The
version number is not the outcome: public Scribble already uses 2.0.91.

| Milestone | Required outcome | What it does not certify |
| --- | --- | --- |
| Integrated core development candidate | Verified Excel analysis and four-slide output; workbook-backed repair of a six-slide source; safe bounded Excel edits; repeatability and held-out evaluation on a declared configuration | Every Office operation, arbitrary decks, every model, or public availability |
| Deployment-qualified core | The accepted core works with the intended local runtime and actual Office/deployment environment | Hosted success alone cannot establish this |
| Broader suite qualification | Word and Outlook-origin adapters reuse the same verified analysis; existing Outlook/browser behavior remains regression-tested; additional capabilities have their own evidence | A core pilot does not automatically qualify the entire suite |
| Public promotion | Exact accepted artifacts, recovery instructions, and a new explicit owner release request | A green PR or development build cannot lift the public freeze |

No stage is complete because code exists, a model says it succeeded, or a
component test passed. Completion requires the applicable current-candidate
receipts. Unsupported inputs return a precise boundary; they do not count as
successful supported tasks.

## 2. Why delivery stalled and the engineering decisions

| Failure mechanism | Decision | Required proof |
| --- | --- | --- |
| Model-authored numbers, AI reviewers, and code compete as authorities | Code owns source bindings, arithmetic, dates, units, identity and execution; the model selects and explains evidence | Wrong AI arithmetic cannot change a verified result |
| Blank cells, formulas, source identity or period meaning are lost during capture | Preserve typed cells, coordinates, source instances and calculation state before planning | Changed schemas, missing values, dates and duplicate source text retain their meanings |
| Nested reviewer/repair loops multiply work | Persist one task-wide allowance; patch the defective field; keep facts immutable | Transport/retry accounting, restart tests, and measured request/token ceilings |
| Native PowerPoint chart operations can terminate Office | Separate chartless revision from workbook-backed chart reconstruction; fail closed outside certified scope | Current-binary native checks with no process exit and preserved sources |
| A helper works but the user's task never finishes | Exercise request construction, public tools, journal and terminal completion together | End-to-end native receipt, including the actual task completion event |
| Repair code depends on PP01 slide positions and text | Establish one working route, then replace fixture assumptions with explicit semantic bindings | Reordered slides, changed labels/data and sealed unseen tasks |
| Mechanically valid decks remain visually weak | Approve exact rendered baselines; distinguish preservation from redesign; repair layout in the renderer | Identified visual review and calibrated reviewer behavior |
| Many branches and checkpoints obscure the candidate | One integration PR and one candidate-bound acceptance record | Commit, binary, configuration, fixtures and artifacts agree |
| Hosted evidence is treated as local readiness | Qualify runtime and deployment separately | Real tool/vision/context tests alongside Office on the target machine |

Retain the existing native writers, analysis contracts, source registry,
journals, deterministic calculations and independent graders. A wholesale
rewrite or replacement renderer is not justified by the evidence. Stop adding
prompt exceptions to compensate for a broken source or execution contract.

## 3. Verified starting position

This section records the position before native execution. The current
candidate's measured progress and remaining gates are in
[DELIVERY_STATUS.md](DELIVERY_STATUS.md); do not treat the historical pending
rows below as the latest status.

Status at implementation commit `05c02212a89b2798a185e6412679869aad6949ea`:

| Item | Evidence and current status |
| --- | --- |
| Integration | PR #41 was the single integration PR; inherited implementation and evidence are preserved in `codex/development` |
| Build and offline checks | Local Release build, all 218 guardrails, static capability scan and delivery-script parsing passed |
| Windows CI | [Run 36157065102](https://github.com/datap0nd/scribble/actions/runs/36157065102) passed on `05c0221`, including browser fixtures, build, guardrails, native-harness upload, ordinary/pilot installer checks and release protection |
| Typed analysis and output compilers | Implemented with offline coverage; earlier native component evidence exists, but acceptance is not transferable to a new binary |
| Bound Excel writes and recovery | Implemented; historical native evidence is in the [Phase 5 checkpoint](tests/benchmarks/stress/PHASE5_CHECKPOINT.md); renew on the candidate |
| Chartless native capability gate | Implemented, with scope-isolation tests; current-binary native qualification is pending |
| Full production PP01 route harness | Implemented; its real-HTTP fake-endpoint protocol passes offline; actual Office execution is pending |
| Request budget | New PowerPoint revision pilot tasks reserve at most 18 OpenAI-compatible requests at the transport boundary; the counter survives restart. This is not proof of universal budgeting for every host/provider or older task |
| Existing visual references | All 18 full-size images inspected and hashes checked by the agent: 16 visual approvals, two rejections (D2P2/D2P3, undersized monthly-series legends). [Recorded review](tests/benchmarks/stress/evidence/phase4-native-candidate/visual-review/agent-review-2026-09-25.json) |
| Current-candidate native, visual and model acceptance | Open. Existing reference inspection does not certify current output, facts, or native editability |
| Hosted pilot, fixed-candidate repeatability, unseen cases | Not completed for this delivery candidate |
| Local runtime and protected deployment | Not certified; historical low-memory probes and native protection constraints remain unresolved |
| Public release | Frozen at 2.0.91; no promotion is part of this PR |

The local offline candidate report is generated under
`tests/benchmarks/generated/delivery-candidate-final/20260925T154650743Z-0bbb6dc497f8494a807ebb3a0dc36b13/`.
Generated local files are not implicitly repository evidence. Archive sanitized
reports with their provenance when establishing a milestone.

The current CI harness is artifact `Phase2NativeHarness` from run 36157065102.
Its Scribble.dll SHA-256 is
`409662a2961e7e033551d7e15874eb776645680aa4c1a8f17de08bb0925aad39`.
Use the actual binary hash for qualification, not the branch name or file
version. CI stamps versions, so the local build has a different hash.

## 4. Supported scope and explicit boundaries

The first two end-to-end proofs remain XA01 (one analysis to a native Excel
report and four slides) and PP01 (repair a six-slide source using its workbook).
They are development cases, not the complete capability definition.

| Capability | Immediate scope | Work required before broader support |
| --- | --- | --- |
| Workbook analysis | Typed, explicitly bound tabular data; host calculations; missing/error values preserved | Remove literal `Period`/YYYY-MM/header dependencies through validated column/period mappings; certify formula-derived source values separately |
| Excel output | New marked unsaved output, requested live formulas, tables and supported native charts | Verify recalculation and source-reference semantics; never substitute constants for requested formulas |
| PowerPoint authoring | Existing native editable layouts using the same analysis artifact | Qualify supported density, long labels, number widths and narrative organization |
| Deck repair | Saved source to a task-owned unsaved draft; bounded chartless patches; separately reconstructed workbook-backed chart | Replace six-slide, chart-on-slide-2, table-on-slide-3, replacement-on-slide-4 and literal-text assumptions before generalizing |
| Native chart operations | Only separately evidenced reconstruction paths | Arbitrary chart copy/edit, extra chart arrangements and unsupported objects remain unavailable until certified |
| Bounded Excel edits | Request-bound worksheet, existing 200-row/30-column/500-character cell limits, preflight and receipted recovery | Renew native proof for identity/focus changes, formulas, interruption and concurrent edits; certify any expanded formula family separately |
| Word and Outlook-origin analysis | Existing behavior stays available under its existing contract | Migrate shared-analysis adapters after the core passes; preserve mail authorization boundaries |
| Browser and other existing surfaces | Maintain existing regression coverage | No blanket new reliability claim from the Office pilot |

Merged/protected ranges, spills, external links, UDFs, ambiguous sources,
unsupported formula families, protected artwork/actions/animations and other
unsupported structures must be preserved or rejected explicitly. Do not
silently approximate them or declare them supported because a simpler fixture
passed. Freeze the positive capability matrix before held-out evaluation.

## 5. Architecture and ownership of decisions

The execution flow is:

`request + authorization → bound source snapshots → verified analysis → semantic plan → staged native operations → readback → scoped review/repair → terminal receipt`

| Layer | Owns | Must not do | Existing integration points |
| --- | --- | --- | --- |
| Request/target binding | User intent, requested outputs, source precedence, destination identity | Infer a different destination from the newly active window | `DocumentChatRequestFactory`, `TaskContextManager`, `TaskRecoveryInput`, `ExcelDraftBinding` |
| Source capture | Instance identity, revision, typed cells/text, coverage and provenance | Treat a cached formula as freshly recalculated or collapse distinct sources | `TaskSources`, `OpenXmlWorkbookSnapshotReader`, `SourceSnapshot`, `SourceLocator` |
| Analysis | Units, periods, calculations, ranking, null policy, precision and contradictions | Accept a reviewer-supplied replacement number as truth | `AnalysisContracts`, `AnalysisCalculator`, `AnalysisTableArtifactBuilder` |
| Semantic planning | Narrative selection, order, emphasis and supported operation choices | Rewrite immutable facts, source precedence or user authorization | `AnalysisSlidePlanContract`, `AnalysisWorkbookPlanBuilder`, `AnalysisDocumentPilot` |
| Compilation/rendering | Cell addresses, formula emission, native layout, bounded geometry variants | Ask the model to repair geometry it cannot control | `AnalysisDocumentCompiler`, native workbook and presentation writers |
| Execution/recovery | Stage, apply, verify, journal and reconcile against captured bindings | Claim a COM transaction or blindly replay an uncertain write | `DocumentDraftHost`, `PresentationRevision`, `PresentationDraftCopy`, `ExcelGridWriteRecovery` |
| Review/repair | Structured content/visual findings with exact target and evidence | Override arithmetic, identities, authorization or broaden repair scope | `AnalysisReviewContract`, `AnalysisRepairBudget`, `AnalysisDocumentRepair` |
| Completion | Verified output identities, checks, unresolved limits and saved/unsaved state | Mark complete while a write, review or recovery outcome is uncertain | Task journal, diagnostics, `CanComplete`/`CompleteTask` |

### Contract requirements

1. **Source snapshots:** immutable instance/revision bindings, content hash,
   capture coverage, formula/calculation state and locators. Preserve raw and
   displayed cell values separately. Equal contents do not imply equal source
   identity. Refresh creates a new revision and invalidates affected plans.
2. **Facts and datasets:** host-issued IDs; type, metric, unit/currency, period,
   dimensions, precision, source locators and verification status. Distinguish
   zero, blank, error, unknown and not applicable. A model cannot register an
   invented value as verified merely by supplying an ID.
3. **Calculations:** declared inputs and host-computed results. Cover sums,
   grouped/filtered totals, count/null policies, differences, ratios, growth,
   margin, ranking and comparisons. Validate units, period alignment, duplicate
   rows, ties, zero denominators and cycles. Round only for display.
4. **Plans:** Excel and PowerPoint reference the same analysis revision. Emit
   invariant-culture native formulas with validated dependencies. Resolve fact
   text and citations in code. Keep logical slide ID, native SlideID, shape ID,
   display order and page number distinct.
5. **Findings:** code, owner, target, evidence/measurement, severity and supported
   action. A suspected binding error returns to analysis. A layout defect goes
   to a renderer operation. Unsupported prose receives a targeted content patch.
6. **Receipts:** task/operation identity, pre-state, intended operation,
   post-state, verification and recovery status. Persist reservations before
   mutation. Version contracts beside the existing workflow version; never
   reinterpret an old journal as a new-schema task.

Stay on .NET Framework 4.8/C# 7.3 and the existing serializer/writers. Register
new source/test files in the classic project files. Keep existing tool names
and draft markers; any necessary tool-surface change must update the guardrail
allow-list deliberately in the same reviewed change. Preserve `LegacySamsung`.

## 6. Ordered implementation work packages

The implementation owner is accountable for every package below. The owner is
asked only for missing intent, explicitly required native-automation permission,
additional spending authorization if needed, or eventual public promotion.
No additional approval loop is required for routine code, tests or delegated
product/design decisions. These are dependency gates, not new phase PRs.

Critical path: D0 → D1 → D2 → D3 → D4 → D5 → renewed candidate gates →
D6 → D7 → D8. D9 broadens the qualified core; D10 remains a separate owner
release decision. Offline preparation may proceed while native permission is
pending, but later acceptance cannot skip an unresolved prerequisite.

### D0 — establish and maintain one candidate (code baseline complete)

- Keep PR #41 as the integration point; reconcile any new upstream change
  before merging. Update this plan's status and evidence in the same PR.
- Record commit, source tree, binary/installer hashes, build run, feature flags,
  fixture seal, Office build/bitness/fonts/locale and runtime configuration.
- Preserve original source documents, user-open Office sessions and existing
  unsaved work. Never terminate Office by process name as cleanup.
- Maintain a single capability/gate ledger. A historical component pass is
  recorded as historical, not copied into the current candidate's pass fields.

**Exit:** reproducible candidate identity, clean intended diff, green applicable
CI and a truthful ledger. The code baseline above meets the offline/CI part;
subsequent code changes require the affected gates again.

### D1 — qualify bounded native execution

- The owner granted current-task computer-use permission on 26 September 2026.
  Keep Office runs on disposable inputs and preserve existing sessions.
- Run chartless revision acceptance on the exact candidate: text and font edits,
  table edits, annotation/notes, move/insert/replace/delete within authorization,
  preservation, revert, injected failure and concurrent-edit behavior.
- Accept only a matching binary/policy `chartless-v1` receipt. It must not claim
  `all_operations_passed` or enable chart operations. Reject chart-bearing
  operations before staging or permission consumption.
- Renew native Excel binding/recovery checks on the same candidate, including
  focus changes, live-formula recalculation and partially written ranges.
- If Office exits, retain the stage, input/output identities and crash evidence.
  Isolate the offending operation. Do not repeatedly run the entire workflow
  or disable the guard to obtain a passing receipt.

**Exit:** current-binary native reports establish every enabled operation's
preservation and recovery behavior; no wrong-target writes, unreported
uncertainty or Office process exit. Unsupported chart paths remain unavailable.

### D2 — prove the actual production routes without paid inference (open)

- Drive the real request factory and exposed schema through actual HTTP fake
  responses, inspection tools, public draft host, write journal and terminal
  task completion. `PilotRouteNativeAcceptance` is the PP01 entry point.
- Verify the same task-owned six-slide unsaved draft is returned; every required
  slide was inspected; requested notes/content/owner-date pairs remain; the
  source deck and workbook are unchanged; the rebuilt chart uses bound facts.
- Verify every write is resolved and the terminal receipt occurs within the
  task budget. A correct PPTX with a stalled task is a failure.
- Extend or verify the existing native analysis harness covers the equivalent
  complete XA01 route through the public host and terminal completion, rather
  than stopping at successful compiler/helper calls.
- Exercise a failed tool, malformed proposal, rejected review, cancellation,
  transport retry, exhausted allowance and restart. Preserve any uncertain
  draft, stop blind replay and report a precise non-success state.
- Run the independent artifact grader on saved **test captures**. Test harness
  capture is not permission for the product to save the user's source.

**Exit:** both fake-endpoint routes reach verified completion with correct native
outputs and preserved sources; negative paths do not report success. Fake
reviewer approval establishes plumbing only, not model or visual quality.

### D3 — settle the visual standard and current-output quality (open)

- Correct the essential legends on reference D2P2/D2P3; render them again and
  record new hashes and reasons. Retain the rejected images as prior evidence.
- Cover six families at normal/long/maximum supported density: scorecard,
  comparison chart/table, grouped chart, evidence cards, dense table, and
  cover/closing. Use existing layout IDs where possible.
- Judge full-size rendered pages for legibility, hierarchy, contrast, clipping,
  collisions, chart meaning and retained content. Distinguish a dense reference
  page from a presentation summary; distinguish requested redesign from repair.
- Independently verify native chart/table editability and source values. PNG
  inspection cannot establish either. Review the actual D2 output, not only
  layout specimens, including narrative order and useful emphasis.
- Record reviewer identity and rubric version against exact output hashes.
  Delegated agent judgment is explicitly labeled `agent`; it must never become
  a fabricated human verdict. The historical human-verdict file stays intact.
- Prepare automated-review calibration against approved examples and at least
  18 seeded defective variants. Run model-based calibration under D4's runtime
  and spending preflight before relying on automated approval. Require zero
  accepted critical defects, zero
  arithmetic/identity overrides, at least 90% recall of noncritical blockers
  and at most 10% false blockers on approved examples.

**Exit:** all required baseline/output pages have explicit accepted verdicts
and independent native/factual checks, with identified agent review under the
owner's delegation where automation is not yet calibrated. Automated review is
not certified until calibration passes; pending or rejected pages cannot be
reported as approved. This permits D4 calibration without a circular gate.

### D4 — isolate real-model feasibility on one configuration (open)

Prerequisites: D1–D3 pass for the same source/binary and runtime/budget preflight
in section 8 is recorded.

- Run one controlled XA01 and one PP01 development attempt on the declared
  configuration. Include the user's actual request path and terminal result.
- Capture each author, reviewer, compaction and retry request; record actual
  provider/model, tokens, image payload, latency, cost and stage outcome.
- Probe tool arguments, structured-output enforcement or its bounded fallback,
  representative context and actual image understanding. Record runtime build,
  tokenizer/chat template where available, quantization, context limit,
  reasoning/sampling settings and output limits. Model-name heuristics are not
  capability certification. Complete the D3 calibration before treating an
  automated visual verdict as acceptance; retain identified independent review
  of pilot outputs in either case.
- Compare native output to an independent oracle and review the exact renders.
- On failure, classify the first faulty stage: capture/binding, interpretation,
  plan/schema, calculation, renderer, native execution, review or completion.
  Reproduce it with the smallest offline/native case before another paid run.
- After two unsuccessful attempts at the same architectural hypothesis, stop
  that experiment and change the mechanism under test. Do not add another
  universal prompt prohibition or secretly change providers.

**Exit:** both workflows complete correctly within the efficiency budgets, or
a specific reproducible blocker is recorded. One pass enables continued work;
it does not satisfy repeatability or generalization.

### D5 — remove fixture dependence and close recovery/migration gaps (open)

- Replace page-number/literal-text assumptions in `DocumentDraftHost.PilotCopy`
  and `PresentationDraftCopy.Pp01NativeStyleOperations` with inspected stable
  IDs, explicit source-to-object bindings and supported semantic roles.
- Bind chart categories/series to the workbook analysis, not a hardcoded number
  of categories. Support only the chart types and transformations actually
  qualified. Reject an ambiguous role mapping before writing.
- Replace hardcoded period/header relationships in `AnalysisTableArtifactBuilder`
  with validated mappings. Certify changed dates/labels/order independently;
  support formula-derived inputs only after proving fresh calculation semantics.
- Route migrated capabilities through typed findings and budgets. Retire the
  corresponding stale-brief/numeric/page exception filters after replacement
  regressions pass. Do not delete legacy protections wholesale.
- Verify transport-wide allowances cover the entire claimed workflow, including
  internal review, compaction, retries and resume. The current 18-request
  PowerPoint/OpenAI-compatible gate is narrower than universal coverage.
- Exercise failure before mutation, during partial writes, after native success
  before receipt persistence, during readback and during recovery. Change the
  active window and edit the destination concurrently. Verify no duplicates,
  no redirected writes, preserved user changes and explicit uncertainty.
- Keep contract versions and feature-flag selection durable. Never retry a
  failed new-path write through the old writer as a silent fallback.

**Exit:** the declared useful capability family survives changed amounts,
headers, periods, labels, slide order and supported density; interruption and
user edits are safe. Development variation is separate from the sealed set.
After these changes, freeze a new candidate and renew D1–D4 evidence as affected.

### D6 — demonstrate repeatability on the frozen candidate (open)

- Require **three consecutive unassisted passes each for XA01 and PP01**.
- Keep binary, fixtures, feature flags, Office environment and runtime/model
  configuration fixed. Restore clean disposable inputs for each run.
- Every pass includes correct native outputs, source preservation, visual
  approval, terminal completion and cost/request compliance. No manual artifact
  repair or selective exclusion of failed attempts.
- Failure resets the affected sequence. A relevant code/configuration change
  invalidates the corresponding candidate evidence; do not combine old passes
  into the new sequence. Preserve failed attempts in the ledger.

**Exit:** six qualifying runs on one candidate/configuration. This is a smoke
and repeatability gate, not a statistical guarantee of broad reliability.

### D7 — evaluate sealed unseen work (open)

- Freeze the supported capability matrix and evaluator before selecting cases.
- Use **12 newly authored tasks: four Excel, four PowerPoint, four Excel↔PowerPoint
  handoffs**, each run twice, giving **24 executions**. Vary amounts, schemas,
  periods, labels, missing data, density and supported user-edit scenarios.
- Keep prompts/oracles outside implementation tuning and model context. Seal
  the set with a manifest hash. Existing stress templates that implementation
  has repeatedly seen are development regressions, not unseen evidence.
- Separate expected unsupported/conflict rejections into a negative suite.
  `Unsupported` on a declared-supported positive case is non-completion.
- Require **at least 23/24 correct, visually accepted completions** and **zero
  wrong-source writes, silent corruption, false success or incorrect numbers
  presented as verified**. The remaining non-completion must be safe and
  diagnosed; no unsupported capability may be silently removed after scoring.
- If a case is exposed for tuning, replace it and reseal before making a fresh
  unseen claim. Record who prepared/reviewed cases; do not describe a same-agent
  tuned dataset as independent evaluation.

**Exit:** the fixed acceptance matrix passes, with all attempts and safety
outcomes reported. The threshold is a product acceptance rule, not a claimed
population reliability percentage.

### D8 — qualify deployment and development-candidate readiness (open)

- Validate the intended local model/runtime with actual tool use, structured
  arguments, vision images, representative context and Office memory load.
  Measure memory headroom, latency and failures; a model name is not a probe.
- Do not silently swap to hosted inference or a separate vision model. Record
  any explicitly selected alternative as its own evaluated configuration.
- Re-run native smoke and recovery on the actual target Office build/bitness,
  fonts and locale. Repeat D6–D7 on the local configuration before making the
  same local reliability claim; hosted results are not transferable.
- Validate protected-workstation behavior separately; local saved-package
  evidence does not prove protected-file access. Preserve the no-save/source
  boundaries and document which readback mechanism is valid there.
- Verify ordinary installer dependencies and opt-in pilot renderer packaging,
  COM registration, upgrade/uninstall behavior and feature-flag rollback.
- Verify user-facing completion reports describe unsaved drafts, checks and
  limitations accurately. Show an actionable conflict/unsupported reason,
  without exposing implementation details unnecessarily in ordinary flows.
- Assemble the final evidence index and capability matrix. Keep the PR draft
  until the core gates and integration review are satisfied; merge into
  `codex/development` only as a tested development candidate.

**Exit:** a reproducible, deployment-qualified core candidate with documented
limitations and recovery. Hosted-only results must retain a hosted-only claim
until local qualification passes. Public promotion remains separate.

### D9 — migrate the remaining adapters and qualify broader scope (later)

- Reuse the same analysis ID and source revision in Word reports and
  Outlook-origin handoffs; do not recompute facts merely because the initiating
  pane changes. Retain reading/drafting permissions; never add mail sending.
- Verify native Word structure/layout/citations and cross-app agreement on
  values, source provenance, revision and completion. Preserve old task routing.
- Run the existing Outlook/browser/Office regression suites. Expand the 200-case
  live corpus in budgeted batches after the core gates pass.
- Publish an internal capability matrix distinguishing implemented, tested,
  supported and unsupported combinations by source, operation and runtime.

**Exit:** matching evidence for broader suite claims. Additions cannot dilute
the accepted core's source-preservation, budget or completion guarantees.

### D10 — public promotion (not authorized by this plan)

Follow [release-channels.md](docs/release-channels.md). Public Latest,
`continuous` and the stable source remain at the original 2.0.91 artifact.
A new explicit owner release request is required to lift that freeze. Prepare
exact artifact hashes, successful acceptance/build records, recovery copy and
manifest for that decision; do not rebuild or relabel the old stable artifact.

## 7. Acceptance, stop rules, and evidence format

The native lifecycle is `Captured → Planned → Validated → Staged → Applying →
ReadbackVerified → Complete`. Exceptional outcomes are `Conflict`,
`Unsupported`, `NeedsInspection` or `Failed`. A crash during `Applying` is
uncertain until native state is reconciled. COM is not a general atomic
transaction; compensating restoration is claimed only after verified readback.

| Dimension | Required evidence | Blocks acceptance when |
| --- | --- | --- |
| Deterministic correctness | Independent expected values, types, periods, ranking, formulas and source bindings | Any wrong value is presented as verified, or a required check is absent |
| Preservation | Source/destination identity and pre/post relevant fingerprints; unsupported objects retained | Wrong-target write, silent loss or unexplained source change |
| Native output | Editable objects, formula recalculation and independent artifact readback | A raster substitute/constant violates a requested native object/formula |
| Recovery | Before/after receipt, failure injection and restart/concurrency results | Blind replay, duplicated output, overwritten user change or false rollback claim |
| Visual/content | Exact rendered artifact hashes, rubric, identified reviewer and reasons | Required page is pending/rejected, evidence is unreadable or requested content is dropped |
| Execution bounds | Whole-task request/token/cost accounting, including retries/review | Limit bypass, unbounded loop or omitted provider usage |
| Terminal behavior | Verified write journal and actual completion event | Artifact exists but task stalls, or success hides unresolved writes/reviews |
| Generalization | Frozen scope, sealed cases, full denominator and negative tests | Fixture tuning is counted as unseen success or unsupported cases are relabeled |

The acceptance record must contain candidate commit/binary/installer hashes;
CI run; contract and flag versions; fixture manifest/input hashes; Office and
runtime configuration; case/run/attempt IDs; stage statuses; source-preservation
and output hashes; deterministic/native/visual/model/generalization results;
request/token/image/cost/time measurements; reviewer identity; failure stage;
and the next action. Retain sanitized artifacts plus a protected trace location.
Never commit API keys or private user document content.

`Test-DeliveryCandidate.ps1` currently records separate offline, native scope,
production route, paid-model, visual and generalization fields. It creates a
new attempt directory, rejects existing PowerPoint sessions and restores the
previous native receipt after testing. Extend the evidence index as later gates
close; do not manually flip its full-acceptance flag to bypass missing evidence.

Code that changes the assembly invalidates its native acceptance receipt.
Changes to rendering require affected visual checks; changes to source mapping,
execution or runtime require affected deterministic/native/model checks. A
relevant change during final repeatability creates a new candidate sequence.
Documentation-only changes can reference the previously tested code commit,
but must not mislabel a different binary as the tested artifact.

## 8. Runtime, inference budget, and efficiency gates

Use the existing hosted `qwen/qwen3.8-27b` configuration for comparable diagnostic
runs only after D1–D3; this does not certify local deployment. Record actual
provider routing. Prefer one supported provider without fallbacks for a matched
comparison; if fallback is necessary, retain actual providers and do not
attribute their latency differences solely to code.

Before a paid batch, inspect the existing authorization and fresh remaining
balance, estimate the batch using measured runs, and preserve the enforced
reserve. Current code in `TestLabStressBudget` caps total key usage at **$40**;
that is the existing ceiling, not a new allowance or a claim that $40 remains.
Do not use dated $30/$5.62 planning figures as current limits or print secrets.
Changing provider, raising the cap, or buying more credit is not authorized by
this document. If funding is insufficient, record the exact unrun gates.

| Efficiency measure | XA01 | PP01 |
| --- | ---: | ---: |
| Whole-task model requests, including internal review, compaction and retries | ≤12 | ≤18 |
| Aggregate prompt tokens, including repeated/cached input | ≤150,000 | ≤250,000 |
| Individual review input | ≤12,000 tokens plus bounded images | ≤12,000 tokens plus bounded images |
| Full-slide regeneration for a wording/geometry-only defect | 0 | 0 |
| Silent data loss or false completion | 0 | 0 |

These are acceptance targets, not measured achievements. The existing review
budget also enforces 36,000 prompt characters and 8,192 response tokens in its
covered path; characters are not equivalent to tokens. Preserve both measures.
The whole-task request counter is currently specific to new PowerPoint revision
pilot tasks over the OpenAI-compatible transport. D5 must prove or extend
coverage for each claimed route; do not infer Gemini/XA01/legacy coverage.

Use a bounded relevant evidence slice rather than repeating full documents.
Do not truncate required facts silently. Persist allowances before sending a
request or applying a patch. A budget error stops the workflow with its draft
and journal retained; starting nested loops or resuming cannot reset it.

Record absolute latency, median and p95 with sample sizes on a fixed runtime.
The previous goal of at least 50% median improvement is a comparison target
only when the baseline is reproducible. Three runs do not justify a confident
p95 claim. If provider tokens/cost are unavailable, mark them unavailable;
do not convert that absence into a pass on a token/cost gate.

## 9. Reproducible execution entry points

Run from the PR #41 checkout. Inspect current script parameters before use.
The native commands below remain subject to the current task's explicit Office
permission; these instructions do not launch or authorize Office by themselves.

Build and offline checks:

```powershell
msbuild Scribble.sln /t:Restore /p:Configuration=Release
msbuild Scribble.sln /p:Configuration=Release
& scripts/Test-Guardrails.ps1
& tests/NativeAcceptance/Test-DeliveryCandidate.ps1 `
  -TestExecutable tests/GuardrailTests/bin/Release/GuardrailTests.exe `
  -OutputDirectory tests/benchmarks/generated/delivery-offline
```

On the current workstation, a non-admin SDK exists at
`%LOCALAPPDATA%/ScribbleDeliveryTools/dotnet/dotnet.exe`. Its `msbuild` invocation
uses `TargetFrameworkRootPath` set to the cached
`Microsoft.NETFramework.ReferenceAssemblies.net48/1.0.3/build/` directory.
This is a local convenience, not a portable path assumption or CI substitute.

The current downloaded candidate harness is under
`tests/benchmarks/generated/delivery-ci-36218395090/harness/`. Its
CI-stamped Scribble.dll was replaced locally with the exact candidate DLL
recorded in `DELIVERY_STATUS.md` before the native run; verify both executable
and DLL hashes before reusing it. The current sealed
stress-corpus manifest is
`64f73305c8cf0cc2efdc5b87c539b833fd442c82c67c0ba2f62f06e0ff689d2e`.
Verify input hashes against that manifest before a native attempt.

```powershell
& tests/NativeAcceptance/Test-DeliveryCandidate.ps1 `
  -TestExecutable tests/benchmarks/generated/delivery-ci-36218395090/harness/GuardrailTests.exe `
  -OutputDirectory tests/benchmarks/generated/delivery-native `
  -SourcePresentation tests/benchmarks/generated/stress-corpus/inputs/powerpoint/PPT01.pptx `
  -SourceWorkbook tests/benchmarks/generated/stress-corpus/inputs/excel/WB01.xlsx `
  -SkipOffline -RunNative
```

`-SkipOffline` is appropriate only with the separately recorded passing offline
checks for that candidate. The script prints the fresh attempt directory; use
its exact candidate path for grading and review. Never mix reports from earlier
attempts into the same pass.

Additional existing harness modes are `--native-xa01-route`,
`--native-xa01-failed-workbook`, `--native-xa01-malformed-workbook`,
`--native-xa01-cancelled`, `--native-xa01-transport-retry`,
`--native-xa01-rejected-review`,
`--native-analysis-pilot`,
`--native-phase5-excel-binding`, `--native-saved-chart-fingerprint`,
`--native-phase4-reference` and `--native-phase4-defects`, each with a report path.
The independent PP01 grader is
`tests/benchmarks/stress/Test-StressNativeGrading.ps1` with `-ScribbleAssembly`,
`-CorpusRoot`, `-ExpectedManifestHash`, `-CaseId PP01`, and `-CandidatePptx`.
Run native checks serially in disposable test-owned sessions.

Paid/live orchestration uses the existing Test Lab and
`tests/benchmarks/stress/Run-StressSuite.ps1`. Specify `-CaseIds` for the approved
batch; omitting it selects the full corpus. Its default host is the installed
build, so verify installed payload hash and flags match the accepted candidate
before running. Do not substitute the runner executable for a correctly
qualified installed Office payload. Summarize usage with the existing
`summarize_usage.py`; never expose credentials in commands or reports.

## 10. Risks and decision rules

| Risk | Observable trigger | Required response |
| --- | --- | --- |
| Native chart crash | Office exits, RPC disconnects or a failed crash report | Stop the affected operation, retain evidence, isolate a smaller native reproduction; do not certify chart support through a chartless pass |
| Incorrect semantic binding | Header, period, unit or object role is ambiguous or changes | Reject before mutation or resolve the precise mapping; do not invent a benchmark-specific synonym and call it generalization |
| The real model still fails with correct inputs | Repeated failure at the same planning/review stage | Reproduce that stage, check schema/context and measured capabilities, then make a controlled configuration comparison if needed |
| Local model leaves insufficient Office headroom | OOM, paging/latency collapse, failed vision/context test | Leave local deployment unqualified; measure a suitable configuration rather than claiming hosted evidence transfers |
| Evaluation exceeds remaining authorization | Estimated batch plus reserve exceeds fresh remaining balance | Run only a justified affordable subset, retain incomplete gates, and seek additional authorization only if necessary |
| Reviewer approves defects or blocks correct facts | Calibration misses/false blockers exceed the fixed rubric | Do not use its approval as acceptance; repair the review contract and retain independent identified output review |
| User changes a bound document during work | Source/destination fingerprint no longer matches | Preserve the user's change, mark conflict/uncertainty and reconcile; never force replay against the active window |
| Different artifacts are discussed as one candidate | Commit, assembly, flags, fixture seal or output hash differs | Split the evidence records and renew affected gates; do not merge incompatible passes |
| Held-out set becomes tuning material | Implementation changes are informed by its answers | Retire exposed cases from the unseen claim and reseal replacement cases |

These are operational decisions, not reasons to add routine permission prompts.
The existing permission, spending and release boundaries remain the only
approval requirements unless the user introduces another explicit constraint.

## 11. Operating rules and immediate handoff

Work on the first unresolved dependency. The scoped D1 native receipts and the
positive PP01 and XA01 D2 routes are recorded in [DELIVERY_STATUS.md](DELIVERY_STATUS.md).
The failed-workbook, malformed-workbook and rejected-review routes have D2
negative receipts. Pre-write cancellation and one recovered HTTP 503 inference
retry are also checkpointed. The rejected-review deck remains uncertain.
Cancellation during a native write, nonrecoverable transport, exhausted
allowance and restart remain open. D3 calibration and D4–D9 remain open.

Each change must name its defect, owning layer, intended behavior, relevant
regression and acceptance evidence. After a failure, preserve the smallest
reproduction and repair that mechanism. Avoid new layouts, new product scope,
provider switching and repeated full runs before the current route is closed.

At each milestone update this document and the PR with: what passed on which
candidate, what remains unproved, the first blocking gate and the next concrete
action. Keep old failures and their context. Do not report progress as commit
count, accumulated test count or a build-number increase.

There is no defensible calendar completion promise before native and real-model
feasibility are established. Sequence and exit conditions are fixed here; dates
can be estimated from measured work after D4. The owner should approve a concrete
missing permission or final release artifact, not repeatedly reapprove routine
implementation choices.
