# Scribble 2.0 pilot closure checkpoint

25 September 2026, branch `codex/reliability-pilot-closure` on top of PR #40.
This is development evidence. Public release remains frozen at 2.0.91.

## What passed

- The offline native Excel analysis route produced a new workbook and four
  slides with source preservation, typed review, recovery, date-column, and
  saved-chart fingerprint checks passing (`analysis-acceptance.json` from a
  local native run using the harness built by CI run 36125034712). The fake reviewer received five
  calls, 17,916 prompt-text characters, and 987,152 request-wire characters.
- A saved chart deck is fingerprinted by reading its existing PPTX package;
  it is never sent through `SaveCopyAs` for edit detection. The native check
  found zero temporary package copies and detected a saved chart edit.
- The ordinary installer excludes PDFtoImage, SkiaSharp, pdfium, and native
  libSkiaSharp. The opt-in pilot installer includes the renderer payload.
- The 18 Phase 4 reference pages are rendered at full size and catalogued
  with SHA-256 values in `evidence/phase4-native-candidate/visual-review/`.
  Every verdict remains pending human review.

## Native PP01 blocker

The new saved defective PP01 fixture is unchanged after both failed native
copy approaches. The workbook is also unchanged. Clipboard copy and
`Slides.InsertFromFile` each caused `POWERPNT.EXE` 16.0.20326.20158 to exit
in `chart.dll` 16.0.20326.20158, exception `0xc0000005`, offset `0x43f399`.
The harness now reports `powerpoint_exited` explicitly. It recorded the
  second failure at `create_draft_copy` / `inspect_source_slide_2` in a local
  native run using CI run 36127062057's harness. The copy pilot now rejects a saved deck with a
native chart before creating a draft (`REVISION_COPY_NATIVE_CHART_UNSUPPORTED`).
This is an architectural failure for saved PP01 on this Office build. The
earlier unsaved-source PP01 candidate passed native checks, but it does not
close the saved-source stress case.

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

Production PP01 copy-and-patch is not wired to the model-facing route because
the saved-chart native acceptance failed. Phase 3 is open; Phase 4 has no
approved reference pages; the paid Qwen pilot has not run and no new paid
calls have been made. Do not spend the $3 pilot allowance until the offline
gate passes. The future merge path remains PR #20 into `codex/development`,
then a few consolidated PRs from this branch, with no more stacked PRs.
