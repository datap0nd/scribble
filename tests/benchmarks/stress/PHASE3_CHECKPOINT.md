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

## Still required for the Phase 3 exit gate

- Migrate the active Samsung review and repair route in
  `DocumentDraftHost.PowerPoint` and `DocumentDraftHost.SlideRepair` to the
  typed analysis contract for supported capabilities, then retire the old
  brief/number/native regex finding filters on that route.
- Extend collision ownership beyond the pilot's central content canvas or
  return an explicit unsupported capability for other geometry. Continue
  testing source-bound facts, page continuations, and retry receipts against
  native outputs and context limits.
- Freeze Phase 3 code after the active-route and geometry gaps close, then
  complete the offline gate. Only after Phases 0–3 pass offline may the pinned hosted
  OpenRouter `qwen/qwen3.8-27b` run the small paid architectural pilot.

No paid model call was made for this checkpoint. The last independently
verified balance remains $14.47 of the $40 key cap at 19:35 UTC on
22 September 2026; it must be refreshed immediately before any paid run.
