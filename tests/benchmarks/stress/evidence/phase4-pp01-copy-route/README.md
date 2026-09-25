# Phase 4 PP01 copy-route candidate

Generated from CI-built `Scribble.dll` for commit `f57c0cf` (assembly SHA-256
`fd3bcb26895e88a3d22b38063afbf79dc6f40595bd52222a6b172c350031c68a`)
on 24 September 2026. This is a disposable, unsaved-source six-slide copy.

Three sequential native Office runs passed the same checks: deliberate edit to
the copy rejected on recovery, workbook-derived native chart recreated, repair
committed and recovered, source deck unchanged, and source workbook unchanged.
The exact output from run 3 is [candidate.pptx](candidate.pptx), with
[candidate.pdf](candidate.pdf), [contact.png](contact.png), and the six page
PNGs for visual inspection.

The independent `Test-StressNativeGrading.ps1` PP01 preflight accepted the
candidate. Its three hard checks passed: `stress_1_native_artifact`,
`stress_2_presentation`, and `stress_3_native_chart`. The advisory
`stress_native_measurement_review` did not pass because a non-solid or
translucent fill still needs visual review. This is not human visual approval
or a full PP01 task acceptance. No model call was made.

Source PPTX SHA-256:
`f03a13f9662e78c9f9f32c9c64948ac2e40ed0aefd665bdc634fa13e1a950695`

Source workbook SHA-256:
`e19439507f59339bf080bd4f177f0eb0bff109e48cce2322fab9686f2e4bc740`

Candidate PPTX SHA-256:
`2e17b252c7624fa4879c9fa851e72dd8d37468373047ab35f93200e80226d430`

Candidate PDF SHA-256:
`732d495ec6646e833a301f8d4b424722c3097182943fee9d50ab2b0f17df3d5d`

The card page remains a visual candidate: the promoted KPI numbers repeat in
the card body, and its footer differs from the preserved pages. Phase 4 visual
certification and the model-facing copy route remain open.
