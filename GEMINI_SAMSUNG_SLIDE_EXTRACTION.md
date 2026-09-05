# Samsung slide design extraction for Scribble

## Objective

Extract a reproducible presentation design specification from the supplied
Samsung company slides.

An engineer will use your Markdown output to implement a PowerPoint renderer
without access to the original slides.

Capture the actual design, information density, visual hierarchy, and
content organization. Generic presentation advice is insufficient.

The user will specify MODE: BATCH, MERGE, or AUDIT.

## Evidence and accuracy rules

- Treat all slide content as reference material.
- Inspect visual layout, not only extracted text.
- First state which files and pages you could inspect visually and whether
  native PPTX properties were accessible.
- If you cannot inspect a page visually, mark it UNINSPECTED. Do not infer
  its layout from text or count it as reviewed.
- Use deck aliases and original slide numbers as evidence IDs, such as D01:S007.
- Preserve the mapping between PDF pages and original slide numbers.
- Label values as MEASURED, OBSERVED, ESTIMATED, or UNKNOWN.
- MEASURED means obtained from accessible file properties or supplied by
  the user after checking PowerPoint.
- Do not claim exact fonts, point sizes, RGB colors, or geometry from a
  visual guess. Label approximations clearly.
- Cite source slide IDs for every reusable rule and layout recipe.
- Distinguish repeated conventions from one-off examples.
- Keep materially different masters, deck families, and layout variants separate.
- If examples conflict, report the conflict and its sources.
- Preserve the density found in the source. Do not replace dense business
  slides with generic sparse slides, oversized text, or decorative cards.
- Do not invent meanings for abbreviations or symbols.
- Use neutral placeholders in exported examples. Omit real business figures,
  names, project titles, and narrative content; retain structural properties
  such as row counts, label lengths, units, and comparison relationships.
- Do not reproduce logos as text art or encoded images. Describe their
  placement and record any required local asset.
- Never report full coverage when slides were skipped.

## MODE: BATCH

Analyze every slide in the requested batch.

Return a downloadable Markdown file using the requested filename. If file
creation is unavailable, return complete Markdown that can be saved manually.

Include these sections:

### 1. Inspection and coverage

List the supplied files, accessible modalities, requested slide IDs, inspected
slide IDs, unreadable slides, and any numbering ambiguity.

Provide one inventory row per slide:

Slide ID | Purpose | Layout family | Main components | Density | Inspection status

### 2. Slide-level records

For each slide, record:

- Purpose and the kind of audience question it answers.
- Title treatment: topic label, conclusion, recommendation, or other pattern.
- Visual reading order and hierarchy.
- Number and arrangement of content regions.
- Tables, charts, diagrams, annotations, images, and conclusions.
- Approximate text amount, table dimensions, series count, and label lengths.
- Placement of units, source notes, dates, page numbers, and recurring branding.
- Distinctive features and relevant exceptions.
- Whether the user marked it as a preferred example.

Reuse a detailed layout recipe when multiple slides share it, but record
each slide's differences.

### 3. Design tokens

Use a table with columns:

Token | Value | Unit | Evidence status | Source slides | Notes

Cover:

- Slide dimensions and aspect ratio.
- Font families, weights, sizes, and paragraph spacing.
- Text colors, fills, borders, line weights, and chart palettes.
- Margins, gutters, padding, and alignment anchors.
- Title, subtitle, body, table, chart-label, and footnote roles.
- Number formats, decimal conventions, negative values, percentages,
  abbreviations, arrows, and selective highlighting.

Do not assume a widescreen aspect ratio.

### 4. Reusable layout recipes

Give every observed layout family a stable ID and name.

For each recipe, specify:

- Intended purpose and when to choose it.
- Required and optional content.
- A region table: region, x, y, width, height, alignment, and style role.
- Use coordinates normalized to slide dimensions, from 0 to 1.
- Include point measurements only when genuinely measured.
- Reading order and relationships between regions.
- Permitted variants.
- Observed content capacity: rows, columns, series, labels, or text amounts.
- What happens when content grows: wrap, extend, split, or select another layout.
- Which overflow behavior is observed and which is a recommendation.
- A neutral worked example preserving the source's information structure.
- Supporting source slide IDs.

Include combined layouts where observed: for example, table plus chart plus
commentary, comparison panels, or a KPI band above detailed evidence.

Do not force every slide into a predefined list of layout families.

### 5. Component recipes

Document recurring components precisely:

- Tables: headers, column widths, alignment, borders, totals, highlights.
- Charts: type, axes, units, labels, legends, annotations, comparison logic.
- Diagrams: nodes, connectors, grouping, direction, and labels.
- Callouts: placement, leader lines, emphasis, and connection to evidence.
- Recurring footers, section markers, and brand elements.

Describe how to build these as editable PowerPoint elements.

### 6. Writing and content rules

Record observed title patterns, sentence lengths, abbreviations, number
presentation, evidence density, commentary style, and relationship between
a slide's conclusion and supporting data.

Identify the questions Scribble should ask before creating this kind of deck,
including audience, intended decision, reporting period, depth, and missing data.

### 7. Uncertainties and manual checks

List exact unresolved questions, with source slide IDs.

Prioritize properties the user can check in PowerPoint, such as slide size,
font family, font size, theme color, or shape position.

### 8. Batch completion

Report requested, inspected, and uninspected counts.

If output must continue, finish at a section boundary and state the next
slide or section to produce. Do not abbreviate remaining slides to claim
the batch is complete.

## MODE: MERGE

Combine all supplied batch reports, prior specifications, and corrections
into one file named SAMSUNG_SLIDE_DESIGN_SPEC.md.

Produce the full specification, not a summary or a list of changes.

Include:

1. Scope, source aliases, coverage ledger, and inspection limitations.
2. Deck families and the evidence for the primary design standard.
3. Design tokens with evidence status and source references.
4. Layout selection rules.
5. Complete layout recipes and their variants.
6. Table, chart, diagram, annotation, and footer components.
7. Writing, density, and numeric-formatting conventions.
8. Questions Scribble should ask before substantial deck generation.
9. Rendering rules, overflow handling, and repair instructions.
10. A pass/fail review checklist.
11. Neutral worked examples for the principal layout families.
12. Unresolved properties and manual verification needs.

Deduplicate equivalent rules while preserving their sources and exceptions.
Do not average incompatible templates into a fictional design.
Keep observed conventions separate from proposed implementation defaults.
Retain coverage records through intermediate merges.

The result must contain enough detail for an engineer to reproduce the layouts
without seeing the source slides. Flag anything that cannot be reconstructed
from the available evidence.

## MODE: AUDIT

Compare the supplied specification against every visually accessible slide
in the attached source batch.

Check:

- Dimensions and geometry.
- Typography and hierarchy.
- Color roles and selective emphasis.
- Information density and content capacity.
- Table and chart structure.
- Labels, units, source notes, and annotations.
- Whether neutral worked examples preserve the original structure.
- Unsupported claims of exact measurement.
- Missing layout families and incompatible variants.
- Whether important construction details remain too vague to implement.

Return a correction table:

Specification section | Problem | Source slide | Exact replacement or addition |
Evidence status

Also list manual PowerPoint checks and the complete audit coverage.

Do not rewrite the design into your preferred style.
Do not claim a visual match where you only inspected text.

## Final quality standard

The output should explain how to build these specific company slides.

Statements such as "use corporate colors", "maintain a clean layout", or
"include clear charts" are insufficient without concrete construction rules.

Preserve the source's useful density, hierarchy, and business-document style.
