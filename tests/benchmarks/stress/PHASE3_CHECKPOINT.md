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

## Still required for the Phase 3 exit gate

- Migrate the active Samsung review and repair route in
  `DocumentDraftHost.PowerPoint` and `DocumentDraftHost.SlideRepair` to the
  typed analysis contract for supported capabilities, then retire the old
  brief/number/native regex finding filters on that route.
- Extend collision ownership beyond the pilot's central content canvas or
  return an explicit unsupported capability for other geometry. Continue
  testing source-bound facts, page continuations, and retry receipts against
  native outputs and context limits.
- Finish and record the current full CI result, then freeze Phase 3 code for
  the offline gate. Only after Phases 0–3 pass offline may the pinned hosted
  OpenRouter `qwen/qwen3.8-27b` run the small paid architectural pilot.

No paid model call was made for this checkpoint. The last independently
verified balance remains $14.47 of the $40 key cap at 19:35 UTC on
22 September 2026; it must be refreshed immediately before any paid run.
