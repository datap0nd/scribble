# Native acceptance checkpoint — 2026-09-22

Acceptance requires **three consecutive unassisted live passes each** for XA01 (Excel to four-slide deck) and PP01 (repair existing six-slide deck). Each pass must verify factual accuracy, preserved source work, and rendered slide quality. CI and draft creation alone do not count.

| Case | Last live failure | Fix / next hypothesis | Consecutive passes |
| --- | --- | --- | --- |
| XA01 | 2.0.368 clean deck initially hit a false-negative hierarchy gate; sealed evidence regraded under 2.0.370 passes all 10 hard gates and visual review. Independent 2.0.370 run at 13:26–13:41 UTC passes all deterministic gates, source preservation, and four-slide visual review. | The earlier oversized incidental `33` and grader false negative are fixed. One more independent unassisted live pass is required. | 2/3 |
| PP01 | 2.0.368 appended six repairs to six source slides (12 total); 2.0.369 was first interrupted by a closed PowerPoint window, then stalled before write over unlabelled workbook totals. | Full-deck repair now routes to a separate unsaved draft and host-calculated totals include period, metric labels, and units. 2.0.371 live run started 13:42 UTC to test exactly six output slides, complete native chart, and unchanged sources. | 0/3 |

At 12:23 UTC, OpenRouter confirmed the user's **$10 top-up** (account balance **$10.28**) and the active no-reset inference key's total cap was increased from **$20 to $30**, with prior key usage **$19.722619137**. The stress harness refuses any key cap above $30, checks before each model request, and retains a $0.05 reserve near the cap. After the 13:41 XA01 pass, key usage was **$20.688245519 / $30**, leaving **$9.311754481**. PP01 is running on 2.0.371; refresh budget afterward. Do not infer passes from CI or a partial native deck. Preserve Office originals and uncertain drafts for inspection; never replay a changed write without receipt reconciliation.
