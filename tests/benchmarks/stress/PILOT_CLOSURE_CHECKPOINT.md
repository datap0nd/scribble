# Scribble 2.0 pilot closure checkpoint

25 September 2026, branch `codex/reliability-pilot-closure` on top of PR #40.
This is development evidence. Public release remains frozen at 2.0.91.

## What passed

- The offline native Excel analysis route produced a new workbook and four
  slides with source preservation, typed review, recovery, date-column, and
  saved-chart fingerprint checks passing (`analysis-acceptance.json` from a
  local native run using the harness built by CI run 36125034712). The fake
  reviewer received five calls, 17,916 prompt-text characters, and 987,152
  request-wire characters.
- A saved chart deck is fingerprinted by reading its existing PPTX package;
  it is never sent through `SaveCopyAs` for edit detection. The native check
  found zero temporary package copies and detected a saved chart edit.
- The ordinary installer excludes PDFtoImage, SkiaSharp, pdfium, and native
  libSkiaSharp. The opt-in pilot installer includes the renderer payload.
- The 18 Phase 4 reference pages are rendered at full size and catalogued
  with SHA-256 values in `evidence/phase4-native-candidate/visual-review/`.
  Every verdict remains pending human review.
- The saved defective PP01 fixture passed native chartless copy, bounded
  repair, workbook-derived native chart recreation, recovery, source hashes,
  and the independent PP01 grader's three hard checks. See
  `evidence/phase4-pp01-saved-chartless-candidate/`.
- The complete 200-case stress kit was regenerated and sealed with defective
  PP01. Manifest SHA-256:
  `64f73305c8cf0cc2efdc5b87c539b833fd442c82c67c0ba2f62f06e0ff689d2e`.

## Native PP01 crash path and repair

Whole-slide clipboard copy and `Slides.InsertFromFile` each made
`POWERPNT.EXE` 16.0.20326.20158 exit in `chart.dll` 16.0.20326.20158,
exception `0xc0000005`, offset `0x43f399`. An early harness error also opened
the source with `Untitled=true`, removing its saved path and selecting the
unsafe chart inspection path. The final native route opens a read-only saved
source, copies the chart slide's non-chart shapes and notes, applies text/table
repairs while that draft slide is chartless, then recreates and fingerprints
the native chart from WB01. The source and workbook remain unchanged. The
final native report records `powerpoint_exited=false`. The harness reports
PowerPoint exits explicitly for any future crash. Other saved-chart layouts
still fail closed as unsupported.

## Renderer packaging choice

1. **Keep in every installer:** simplest layout, but brings the pilot-only
   native renderers into every Office process and install.
2. **Ship only with the pilot (current recommendation and implementation):**
   keeps ordinary Office installs lean while the pilot is evaluated.
3. **Render in an existing helper executable:** isolates native PDF/image
   libraries from Office, at the cost of an IPC contract and a separate
   failure/recovery path. Consider this if the pilot renderer itself proves
   unstable; it does not solve PowerPoint's chart.dll copy crash.

## Open gate

Production PP01 copy-and-patch is not yet wired to the model-facing route.
Phase 3 remains open for that route and failure injection. Phase 4 has no
approved reference pages. The paid Qwen pilot has not run and no new paid
calls have been made. Do not spend the $3 pilot allowance until the offline
gate passes. The future merge path remains PR #20 into `codex/development`,
then a few consolidated PRs from this branch, with no more stacked PRs.
