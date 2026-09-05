# Samsung renderer 1.0

The supplied `SAMSUNG_SLIDE_DESIGN_SPEC.md` is retained unchanged as reference data. Its statements about source decks and measured fidelity are supplied claims, not independently verified against the private decks.

The implementation owns the theme, geometry, font roles and minimum sizes. Content schemas expose layout names, text, editable table/chart data, semantic highlight rows and evidence, never arbitrary styling or execution instructions. New presentations use 960 x 540 points. Existing presentations with other dimensions are rejected without resizing their slides.

## Reconciled measurements

- The cover's negative top accent coordinate is omitted. The measured bottom accent remains.
- Shared action-title guidance at 16.1% takes precedence over 21.2% observations. Captions and units occupy the next band.
- Content that extended beyond 100% or collided with footers is constrained to the content area. Recipe B's overlapping title/takeaway and lower panes are separated. Comparison and chart layouts preserve distinct title, action, evidence and conclusion bands.
- Page numbers use a 6% wide right-aligned region ending at the right edge, accommodating multi-digit numbers that cannot fit the supplied 3.9% box at 10.5 pt.
- Action-list rows may use double-height slots for wrapped content. Long tables repeat their headers on continuation slides instead of discarding rows. Native text metrics are checked again after deterministic font measurement.
- Missing Samsung Sharp Sans falls back explicitly to Arial. Korean text uses Malgun Gothic. Final appearance depends on the installed work-PC fonts and requires comparison with the original decks.

## Generation and review

The drafting prompt establishes missing brief details and a source-backed storyline before generation. Independent source review rejects unverified excerpts and unsupported numbers. The host preflights the layout before creating any slides, uses native editable PowerPoint shapes/tables/charts, renders each created slide to PNG, and reviews it with the user's configured vision-capable model. Automatic typography repair is limited to owned, unchanged draft shapes and design minimum sizes. Unresolved defects leave an explicit incomplete draft; no presentation is saved.

Source references and supporting excerpts are retained in notes. Long references use a visible pointer to notes. Input batch limits report an error rather than silently truncating slides, rows, cells or chart values. The first batch freezes ordered storyline IDs; completion requires a reviewed receipt for every planned ID. Shared task checkpoints retain rendered evidence and prevent repeating uncertain writes.

Automated checks cover layout bounds, row preservation across pagination, source-number validation, font mapping and explicit overflow rejection. These checks do not establish a pixel-perfect match to the original Samsung slides. Live PowerPoint rendering and font/connector behavior remain work-PC acceptance checks.

Native API behavior follows Microsoft's documentation for [text bounds](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textrange.boundheight), [slide image export](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.export), [chart source ranges](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.chart.setsourcedata), and [ownership tags](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.tags). Image export does not save the presentation.
