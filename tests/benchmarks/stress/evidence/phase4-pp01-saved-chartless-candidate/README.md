# Saved PP01 chartless-copy candidate

Generated on 25 September 2026 by the CI-built native harness from commit
`6a75922` with `SCRIBBLE_ANALYSIS_PILOT=1`. The source deck was opened read-only
as a saved PPTX. Its chart slide was copied as ordinary shapes, bounded text
and table repairs were applied to the unsaved working deck, then a native
editable chart was recreated from WB01. No model call was made.

`native-acceptance.json` records six-slide patch acceptance, workbook-derived
chart recreation, deliberate draft-tamper rejection, source preservation, and
`powerpoint_exited=false`. The source PPTX SHA-256 stayed
`d6aeedb79018a8808181488b1814e3a83b8efc482b70e95b078d7a3b410d77a4`;
the workbook stayed
`aadb514adfd8270d5cee2bf6d17428a662d459dd1de54ef75071a2819afff978`.

The independent native PP01 grader accepted the candidate against sealed kit
manifest SHA-256
`64f73305c8cf0cc2efdc5b87c539b833fd442c82c67c0ba2f62f06e0ff689d2e`.
Its three hard checks passed: native PPTX readback, six-slide presentation
content/geometry, and exact native chart categories/series. The advisory
measurement review flags a non-solid chart fill for human visual review.
`native-grading-preflight.json` records the exact result.

| Artifact | SHA-256 |
| --- | --- |
| `candidate.pptx` | `7af7ab262bf830f367733ecbaa7cedc0aab29a6b107887805cce0a05b05b5cb8` |
| `candidate.pdf` | `25a91dc31f0e0e34c87db1bdd6cfb2703aa75f9c7b420192593db2437d3e022e` |
| `contact.png` | `1f03f7cc349ed12097942a0ff6ea227ecbd159e6fe2f1b4fe463aac763838d5e` |

The six `page-*.png` files are full-size renders. Visual approval and the
model-facing route are still open; this is a native offline candidate.
