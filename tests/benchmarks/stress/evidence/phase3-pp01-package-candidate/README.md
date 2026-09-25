# PP01 owned-package chart fingerprint preflight

This is disposable native Office evidence from the resealed PP01 stress
fixture, not a hosted-model output or a completed pilot verdict.

- `candidate.pptx` is a six-slide editable, unsaved-copy output. Its PDF and
  `slide-1.png` through `slide-6.png` are review renders.
- `native-acceptance.json` records source and workbook preservation,
  chart reconstruction, draft-conflict rejection, pre-chart recovery,
  and `powerpoint_exited=false` for assembly SHA-256
  `9dc041bfca7b2b4eb61b740419ef9fef484ff0d452d301aee7c9355fd9f8f20d`.
- `native-grading.json` accepts the independent hard artifact,
  presentation, and chart checks against sealed manifest SHA-256
  `64f73305c8cf0cc2efdc5b87c539b833fd442c82c67c0ba2f62f06e0ff689d2e`.
  A translucent chart fill still needs human visual review.
- The three `*failure.json` files preserve earlier attempts where chart COM
  fingerprinting terminated PowerPoint. The passing run uses an owned,
  unsaved draft package to fingerprint the chart slide and embedded data.

This acceptance exercises `PresentationDraftCopy` through a native harness.
The exact model-facing `revise_slides` wrapper and general PowerPoint
revision receipt remain separate open checks.
