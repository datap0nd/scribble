# Phase 4 six-slide repair scope

Status: design decision for the development pilot; native PP01 acceptance is pending.

PP01 requests one editable six-slide output in the Samsung MD visual language, a recreated monthly native chart, repaired period/category labels, readable text, and preserved useful tables and source notes. The original presentation and workbook remain source documents. A repair runs in a separate unsaved draft copy, verifies the copy against the source before writing, and exports a six-slide candidate only after native readback. An interrupted copy is reconciled by its own journal; it is never treated as the original deck.

| Requested change | Allowed operation on the draft copy | Required evidence |
| --- | --- | --- |
| Correct a unique stale period, unit, or sentence | `replace_text` on one identified shape and literal span | Source fact or a user-requested wording change; exact before text, source fingerprint, and readback |
| Correct one native table value | `table_cell` on the identified row and column | Verified workbook fact; table labels, other cells, and notes preserved |
| Correct chart values when categories, series count/order, type, units, and source workbook are already right | `chart_point` on the identified series/category | Verified fact and embedded chart workbook readback |
| Fix the stale monthly categories, wrong series, nonzero axis, or a broken/uneditable chart | Recreate the chart slide with `replace_slide` only when that slide's remaining useful content is explicitly carried into the replacement | Six `YYYY-MM` categories, required series and order, zero axis, native chart and embedded workbook readback, source notes, and preservation checks for protected objects |
| Increase undersized text or clear overlap/overflow | Bounded native renderer operation on the identified shape, or `replace_slide` when the composition cannot be repaired without deleting useful content | Native geometry/readback and an identified visual review |
| Restore citations | Preserve the existing notes and append only verified new source references | Exact source IDs and workbook/sheet/cell lineage |

No operation changes an unrelated slide, chart datum, table cell, note, action, hyperlink, animation, artwork, or background. The existing `PresentationRevision` restrictions on recomposition remain in force. If a protected object prevents the requested repair, the draft stops with a scoped failure rather than replacing that object silently.

The PP01 grader must evaluate the separate six-slide output and verify the source files are unchanged. Native editability, chart series/categories, table coverage, notes, geometry, and independent visual approval remain separate gates. This document does not claim that the copy-and-patch route has passed them.
