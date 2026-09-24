# Phase 3 checkpoint: typed review and bounded repairs

Recorded 23 September 2026. Phase 3 is **in progress**; its exit gate has not
passed, and the architectural-feasibility pilot is not authorized yet.

## Verified so far

- The development-only `AnalysisReviewContract` binds findings to host-issued
  fact IDs, field targets, logical slide IDs, native slide IDs, page order, and
  rendered fingerprints. It refuses value/citation overrides, unsupported
  route combinations, and contradictory verdicts. Host geometry measurements
  become blockers even when a model omits them.
- `AnalysisRepairBudget` persists model-call and patch-target counts across
  retries: 12 calls for cross-app analysis, 18 for PP01, 8 distinct patch
  targets, bounded prompt size and response tokens.
- A narrative repair changes one plan field, recompiles, and verifies that
  facts, citations, native chart/table bindings, and page count remain stable.
  Renderer-owned operations fix an exact native folio, text overflow within a
  minimum font size, or a shape outside the canvas in a disposable deck.
- [CI run 35837118421](https://github.com/datap0nd/scribble/actions/runs/35837118421)
  built the native harness. On the Windows Office workstation, that harness
  created one synthetic Excel draft and four editable PowerPoint slides,
  verified native formulas and chart values, preserved the source ledger,
  recovered with an isolated new draft after a deliberate mismatch, then
  deliberately damaged and repaired a native folio and an out-of-bounds shape.
  The harness reported `typed_review_contract_passed=true`,
  `renderer_repair_passed=true`, and `full_acceptance_passed=false`.
- The isolated [OfficeIMO evaluation](OPEN_SOURCE_UNBLOCKERS.md) is promising
  on the synthetic deck. It is not a PP01 or production-dependency gate.
- A new reliability defect surfaced while testing the shared native patch
  budget: the first `f4b8413` native run refused a repair because PowerPoint's
  second PNG export had a different byte fingerprint; the next run passed.
  `e7756bf` keeps the PNG hash as exact review evidence but authorizes repair
  against a separate hash of native slide ID, shape IDs, geometry, text, and
  font sizes. The CI-built harness from
  [run 35838835743](https://github.com/datap0nd/scribble/actions/runs/35838835743)
  passed three fresh, sequential disposable Office trials (3/3) with budgeted
  folio and out-of-bounds repair. This is a narrow repeatability check, not
  a full repair/recovery qualification.
- The CI-built `a0a4071` harness from
  [run 35840351830](https://github.com/datap0nd/scribble/actions/runs/35840351830)
  passed the disposable Office trial after deliberately overlapping the
  native chart and table on slide 2. The host issued a `COLLISION`
  measurement bound to both shape IDs. The renderer translated the chart to
  a 6 pt gap, then PowerPoint readback found no remaining measured defect;
  chart-cache values and the source ledger still matched the independent
  oracle. The saved slide was rendered and visually inspected. This detector
  covers central content shapes in the current pilot layouts, not arbitrary
  authored-deck geometry.
- [CI run 35842213293](https://github.com/datap0nd/scribble/actions/runs/35842213293)
  passed after task-level review calls were reserved in the checkpoint before
  inference and corrupt empty receipts were rejected. The subsequent
  `68682ed` patch also reserves a native repair against an exact measurement
  and native-state fingerprint before mutation. Resume leaves an interrupted
  patch pending until a saved deck is reopened and the defect is measured
  clear. [CI run 35842974961](https://github.com/datap0nd/scribble/actions/runs/35842974961)
  passed for the initial reservation implementation. The CI-built harness artifact from
  [run 35843324667](https://github.com/datap0nd/scribble/actions/runs/35843324667)
  passed one fresh disposable Office trial: native formulas, four slides,
  source preservation, isolated retry, typed review and renderer repair all
  reported true with no failure. `full_acceptance_passed=false` remains the
  correct status because this was hand-authored, without model or visual
  attestation. The full Windows CI job for that commit passed.
- The four-slide native output is an accuracy/structure fixture, not an
  approved visual baseline. Visual inspection found repeated data on slides
  2 and 3, a small chart beside an oversized table, and a fourth page with
  little decision content. Source preservation and editable shapes do not
  establish the design quality requested for Scribble 2.0.
- `2f9f148` joins native page capture, task-checkpointed review-call
  reservation, and typed verdict parsing in the development pilot. The
  checkpoint is reloaded before the verdict. Native review completion now
  refuses an approval after a page edit; its state fingerprint includes table
  cell text and native chart series data as well as shape text and geometry.
  The CI-built harness from
  [run 35845406941](https://github.com/datap0nd/scribble/actions/runs/35845406941)
  passed one disposable Office trial, including deliberate native table-cell
  and folio edits that invalidated an earlier approval. Full CI was still
  running when the native result was recorded. This is not yet the active
  Samsung model-facing path or a general authored-deck edit guarantee.
- The CI-built `cc0b760` harness from
  [run 35845849719](https://github.com/datap0nd/scribble/actions/runs/35845849719)
  passed a second disposable Office trial after adding a native chart-series
  edit and restore to the review-freshness test. All six structural result
  fields remained true, with `full_acceptance_passed=false` as expected.
- `33c01f6` closes one source-to-analysis gap in the development pilot.
  `AnalysisTableArtifactBuilder` binds a complete typed period table to
  source-cell facts; it rejects unresolved formula caches, blank metrics,
  duplicate keys, ambiguous headers, partial coverage, and mismatched
  metric/currency labels. This automatic binding currently supports a
  `Period` column with `YYYY-MM` text and exact source metric headers; broader
  semantic aliases need a separately trusted mapping. The CI-built harness
  from [run 35847770074](https://github.com/datap0nd/scribble/actions/runs/35847770074)
  read actual disposable Excel cells into that artifact, then passed the
  live-formula report, four native slides, source-preservation, isolated
  retry, typed review and renderer-repair checks. Full Windows CI for that
  commit passed. No model generated the plan, and
  `full_acceptance_passed=false`.
- `e329b79` connects the same typed table binding to the pilot's actual
  `read_cells` tool route. With `SCRIBBLE_ANALYSIS_PILOT=1`, a complete
  single-page range of at most 500 cells can specify exact period, metric,
  currency and dimension headers. The Office host captures the typed cells,
  rejects unverified values, issues source-cell fact IDs, and passes the full
  artifact to the task evidence store through a host-only payload. The model
  sees fact IDs and source cells; it cannot supply numeric facts. The normal
  read schema and behavior remain unchanged with the pilot flag off. The
  Windows solution build in
  [CI run 35849873136](https://github.com/datap0nd/scribble/actions/runs/35849873136)
  passed. Its CI-built harness passed one fresh disposable Office run,
  including task-checkpointed retention of four live-cell facts, native Excel
  formulas, four slides, source preservation, isolated retry, typed review,
  and renderer repair. `full_acceptance_passed=false` remains correct. The
  full Windows CI run passed, including guardrails, installer construction,
  and installer smoke checks.
- `56bc6a0` removes model-authored workbook formulas from the structural
  pilot. `AnalysisWorkbookPlanBuilder` derives report rows, source worksheet
  ranges, `SUMIF` formulas and expected fact IDs from the verified typed
  table. It refuses dimensioned or otherwise unsupported fact sets rather
  than guessing an aggregation. The CI-built harness from
  [run 35851458827](https://github.com/datap0nd/scribble/actions/runs/35851458827)
  passed one disposable Excel/PowerPoint run with native formula readback,
  four slides, source preservation, isolated retry, typed review and renderer
  repair. `full_acceptance_passed=false` remains correct. The full Windows
  CI run passed.
- `9e39f0f` adds a strict slide-plan intake boundary. It accepts a bounded
  fact-referenced narrative/layout plan, rejects unsupported fields and
  model-supplied formula values, then supplies the workbook rows from the
  host-generated formula builder. The first parser harness run at `b05f3bf`
  rejected null formula fields emitted by the C# fixture serializer before
  any Office write; the corrected parser accepts nulls while rejecting an
  injected `=1` formula value. The CI-built harness from
  [run 35852563754](https://github.com/datap0nd/scribble/actions/runs/35852563754)
  passed one fresh disposable Office run through this boundary, with all six
  structural result fields true and `full_acceptance_passed=false`. Full
  Windows CI passed. The slides remain hand-authored fixture content, not a
  model-generated plan.
- `aca7f91` adds the pilot's first active write route. After an Excel task
  captures a typed analysis, the request swaps `write_draft_sheet` to an
  `analysis_id` and title contract. `DocumentDraftHost` checks the task's
  draft authorization and the exact source identity, re-reads the bound
  workbook range, generates host-owned formulas, and verifies native results.
  A changed source cell is rejected before permission consumption or sheet
  creation. The CI-built harness from
  [run 35855905340](https://github.com/datap0nd/scribble/actions/runs/35855905340)
  passed one disposable Office run through that tool call; the report,
  source-preservation, isolated retry, four slides, typed review and renderer
  repair result fields were true. `full_acceptance_passed=false` remains
  correct: the slide content is still hand-authored fixture content, and the
  active PowerPoint tool has not migrated. The full Windows CI run passed,
  including guardrails and installer smoke checks.
- `bc76b95` adds a fact-referenced schema under the existing
  `send_to_powerpoint` tool name. Each native review image is captured in the
  same export used for its page fingerprint. The CI-built harness from
  [run 35857648842](https://github.com/datap0nd/scribble/actions/runs/35857648842)
  passed in disposable local Office: verified workbook formulas, four editable
  slides, source preservation, isolated retry, typed review, and renderer
  repair. The new image/hash assertion passed. The full Windows CI run passed.
- `231b3f5` routes a bound Excel analysis through the typed deck plan and
  task-tagged new PowerPoint destination. It reserves typed native review and
  task-level geometry repairs, then withholds completion on a rejected verdict.
  `7fde9c1` scopes the task's write receipt per analysis destination, allowing
  one Excel draft and one deck in the same authorized task while blocking a
  duplicate output. [CI run 35858756158](https://github.com/datap0nd/scribble/actions/runs/35858756158)
  passed. These commits are development-flagged; a CI compile and guardrail
  pass alone do not establish native end-to-end acceptance.
- `3e645e6` and `88ee554` add an offline native test of the active deck
  handoff, with a loopback fake reviewer that checks all four image hashes and
  a stale-source preflight case. The full Windows
  [CI run 35860033826](https://github.com/datap0nd/scribble/actions/runs/35860033826)
  passed, including guardrails and installer smoke checks. Its CI-built
  harness was attempted on the local Office workstation, but PowerPoint's COM
  server was unavailable (`0x800706BA`) at the test's presentation-count read
  before the typed handoff. The report says `typed_deck_handoff_passed=false`
  and `full_acceptance_passed=false`; it cannot validate this route. The
  interrupted disposable run left an empty automation Excel process, which
  was inspected and closed. A clean native rerun and failure-path review are
  required before the Phase 3 gate can pass.

## PR #22–#25 review corrections (23 September 2026)

- Phase 1 now resolves Excel's real `DBNull.Value` response for mixed
  `NumberFormat` pages by reading uniform column formats, then individual
  formats only in mixed columns. The work is capped at 500 cells. Unresolved
  formats are marked incomplete and cannot bind verified facts. A regression
  test uses `DBNull.Value` rather than a prebuilt per-cell format matrix; the
  native harness also creates a disposable workbook with a real Excel date
  column and records `native_date_column_passed`.
- The XLSX attachment extractor drops rows containing only styled empty
  cells. A fixture with twelve ledger rows and a styled empty row below them
  verifies that the default `CompleteWorkbookTotals` path still returns host
  totals. Typed `read_cells` fields are emitted only for an explicit
  `analysis_binding`; ordinary reads retain their original payload.
- Windows Application Error event 1000 at 16:31:45 on 23 September identifies
  `POWERPNT.EXE` 16.0.20326.20158 crashing in `chart.dll` with exception
  `0xc0000005`; the same signature occurred twice earlier that day. The
  harness now labels RPC-disconnection failures `POWERPOINT_EXITED` and omits
  its own chart-slide PNG exports. The Phase 3 review capture also omits
  chart exports and withholds model review and a completion receipt when a
  visual page is unavailable. The draft remains pending for visual
  inspection. This is an explicit incomplete gate, not visual approval.
- The OfficeIMO source probe is parked for Phase 4 and its CI job is removed.
  No paid model run was made for these corrections.
- A disposable local run of the CI-built harness from commit `39f2cf8`
  reached the typed handoff. Its report recorded
  `native_date_column_passed=true`, `powerpoint_exited=false`, and PNGs for
  the three non-chart slides. A subsequent `POWERPNT` call returned
  `RPC_E_CALL_REJECTED` after the route result, so this run cannot qualify the
  handoff. Commit `952f086` makes the harness report the route result before
  making that extra COM call. No new PowerPoint Application Error event 1000
  accompanied the run. A further native check requires owner permission
  under the current workspace AGENTS.md instructions.
- After the owner approved a native rerun, the CI-built harness from
  `096f653` again passed the real date-column check but returned
  `0x800706BA` inside the typed deck tool result. Windows Application Error
  event 1000 at 11:16:57 on 24 September again identified the same
  `POWERPNT.EXE`/`chart.dll` access violation and fault offset. This showed
  that the writer's own chart preview export remained a crash path even
  after the harness and review capture stopped exporting charts. Commit
  `3184600` suppresses that writer export only in the development pilot,
  including journal resume, and maps the RPC failure to an explicit
  `POWERPOINT_EXITED` tool error. The CI-built `3184600` native rerun still
  faulted in `chart.dll` at the same offset (Application Error event 1000,
  11:22:26 on 24 September). Its report correctly recorded
  `native_date_column_passed=true`, `powerpoint_exited=true`,
  `full_acceptance_passed=false`, and an unapproved pending draft. Known
  pilot preview exports are suppressed, but the cause of the remaining
  crash was not yet isolated. The review path therefore failed closed;
  no model approval or Phase 3 exit claim followed.

## Native chart review recovery (24 September 2026)

- Closing the first chart deck (`c48d2f5`) did not prevent the same
  `chart.dll` crash. A trace from `445caf1` located it after chart creation
  and before the second page's journal receipt completed. The crash was in
  the journal's live chart fingerprint path, not the chart writer or the
  presence of a second open deck.
- `908987c` fingerprints a chart page using PowerPoint `SaveCopyAs` and
  hashes the slide package with its related chart, embedded workbook and
  layout parts. The journal and review freshness checks use this package
  state instead of reopening the fragile live chart COM object. Edits to
  table cells and chart series invalidate the previous receipt. A restored
  edit requires a fresh review because PowerPoint can retain a different
  native package state even when displayed text matches again.
- The development pilot exports the unsaved native deck as PDF using
  PowerPoint format 32, then renders each bounded page with PDFtoImage
  5.4.0. The unsaved deck's name, `Saved` state and native fingerprints are
  checked before accepting any page image. Every PNG is hashed into its
  review page record; missing or oversized pages fail closed before a model
  request. A disposable unsaved chart deck kept its name and unsaved state
  through the PDF export. No chart-slide `Slide.Export` is attempted.
- CI-built `035b6f4` passed three sequential disposable native runs on
  this Office build. Each reported `native_date_column_passed=true`,
  `typed_deck_handoff_passed=true`, `typed_review_contract_passed=true`,
  `renderer_repair_passed=true`, `four_slides_passed=true`,
  `source_preserved=true`, `isolated_retry_passed=true`, and
  `powerpoint_exited=false`. The chart review page was saved as
  `analysis-slide-02.png`; its native chart, values and table were visually
  inspected. These are offline hand-authored runs with a fake reviewer.
  Their `full_acceptance_passed=false` field remains correct: no hosted
  model, calibrated or identified human visual attestation, broad recovery
  qualification, or release gate has passed.
- CI-built `2232692` passed a further disposable native run after moving
  unresolved native geometry rejection ahead of the task's model-review
  reservation. Its regression damages the folio and verifies that no review
  call is consumed. CI-built `2d4f048` and `922288b` each passed a fresh
  disposable native run. These add explicit unsupported-geometry rejection
  for a chart overlapping the title and body text overlapping the source
  footer, respectively. The active tool returns
  `ANALYSIS_DECK_GEOMETRY_UNSUPPORTED` for this unowned geometry instead of
  allowing model approval. Both runs preserved the date column, workbook,
  four slides, typed handoff, and chart review image; neither PowerPoint run
  exited. The full CI run for `922288b` is still pending at this checkpoint.

## Still required for the Phase 3 exit gate

- Finish the development-only analysis-bound PowerPoint handoff with
  offline content-finding failure injection. Its typed native review path
  passes the hand-authored harness, but content findings still need targeted
  corrective operations and recovery receipts. Migrate the
  remaining supported Samsung capabilities in `DocumentDraftHost.PowerPoint`
  and `DocumentDraftHost.SlideRepair`, then retire their old brief/number/native
  regex finding filters as each capability moves.
- The pilot now explicitly rejects semantic collisions outside its bounded
  content canvas. Certify the remaining layout families before widening
  renderer ownership, and continue testing source-bound facts, page
  continuations, and retry receipts against native outputs and context
  limits.
- Freeze Phase 3 code after the active-route and geometry gaps close, then
  complete the offline gate. Only after Phases 0–3 pass offline may the pinned hosted
  OpenRouter `qwen/qwen3.8-27b` run the small paid architectural pilot.

No paid model call was made for this checkpoint. The last independently
verified balance remains $14.47 of the $40 key cap at 19:35 UTC on
22 September 2026; it must be refreshed immediately before any paid run.
