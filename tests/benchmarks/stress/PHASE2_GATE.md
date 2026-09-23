# Phase 2 gate: first complete verified Office path

Recorded 23 September 2026. This is a development-only, hand-authored structural pilot. The model-facing tool-name allow-list is unchanged, and no paid model run was made.

## Candidate and independent checks

- Candidate code: `a5b0702854f9acabeed55677d265eaede3a5d43d`, stacked on `codex/reliability-phase1`. The native harness was built by [CI run 35832192069](https://github.com/datap0nd/scribble/actions/runs/35832192069); its `Scribble.dll` SHA-256 is `5f71cf09975862d7eabfb4f883a8d49224f0093ddab8953f2f87afb4b8a30d0c`.
- Full CI passed: Windows build, `GuardrailTests.exe`, static capability scan, browser fixtures, installer checks, release-evidence gate, and public-release freeze.
- One host-issued `AnalysisArtifact` compiled into a new numbered Excel draft and four native PowerPoint slides. The hand-authored plan used fact IDs for values and source locators for citations; it did not duplicate the oracle's business values or citations.
- Independent synthetic ledger oracle: May revenue 85,519 and cost 36,702; June revenue 82,992 and cost 36,714. Native Excel formula readback matched all four values. PPTX package readback matched the visible slide values and the native chart's two revenue points. Source ledger cells were unchanged.
- A deliberately wrong expected fact made the adapter reject a completed Excel formula write on `Scribble Draft 2`. A corrected retry used `Scribble Draft 3`; the source and original `Scribble Draft` remained unchanged. Neither workbook nor deck was saved as a user document.
- The development path requires `SCRIBBLE_ANALYSIS_PILOT=1`. It uses a bound workbook target and the existing Office writers. There is no automatic fallback to the legacy writer after a failed verification.

## Review artifacts and limits

The synthetic [native evidence](evidence/phase2-native-candidate/native-evidence.json), [Excel PDF](evidence/phase2-native-candidate/analysis-workbook.pdf), [editable PowerPoint copy](evidence/phase2-native-candidate/analysis-deck.pptx), [Excel render](evidence/phase2-native-candidate/analysis-workbook-page1.png), and four [slide renders](evidence/phase2-native-candidate/analysis-slide-01.png) through [slide 4](evidence/phase2-native-candidate/analysis-slide-04.png) are candidate evidence, not approved visual references. Native output correctness and pilot isolation passed. General recovery, ordinary in-place edits, model reliability, and visual certification remain open in later phases.

Visual inspection found no clipped numbers after the pilot formatting fix, but the scorecard uses too much empty space, the comparison table dominates its chart, and the deck repeats the same few facts. These are Phase 4 design blockers. The JSON intentionally records `full_acceptance_passed: false` and `visual_status: candidate_not_approved`.

**Gate decision:** Phase 2's hand-authored, disposable first path passed its structural and native oracle checks. Continue with Phase 3 offline review ownership. Do not start paid Qwen runs until Phases 0–3 pass offline. The public release remains frozen.
