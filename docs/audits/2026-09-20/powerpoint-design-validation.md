# PowerPoint design validation — Scribble 2.0.291

## Verdict

The four-slide XA01 deck is factually accurate and natively editable, but it is not acceptable as a professionally designed presentation. It must not be used as evidence that Scribble's slide creation quality is complete.

## Visual scorecard

| Dimension | Score | Finding |
| --- | ---: | --- |
| Factual accuracy | 10/10 | Values, periods, formulas, chart series and citations match the source. |
| Native editability | 9/10 | Text, tables and the chart remain editable PowerPoint objects. |
| Readability | 7/10 | Core text is readable, but citations are extremely small and the chart/table styling is weak. |
| Information hierarchy | 3/10 | Titles, subtitles, body copy and takeaway banners repeat the same messages without a clear visual focal point. |
| Composition | 2/10 | Slides 1 and 4 are report text placed at the top of a mostly empty canvas. Slide 3 is a small table floating in unused space. |
| Visual storytelling | 2/10 | Only slide 2 turns evidence into a visual comparison. The other slides do not use scale, imagery, diagrams or strong metric typography to explain the story. |
| Design variety | 2/10 | The deck repeats one title/body/footer skeleton and looks generated from a rigid report template. |
| Overall presentation quality | **38/100** | Accurate output, unacceptable presentation design. |

## Slide findings

1. **Headline:** fail. A numeric executive headline should use prominent KPI typography or a deliberate visual composition. This slide is a multiline data dump with excessive empty space.
2. **Period comparison:** partial pass. The chart is correct and editable, but the default chart styling, tiny adjacent table and repeated blue conclusion banner feel mechanical rather than presentation-ready.
3. **Group analysis:** fail. A small native table occupies a fraction of the canvas. The best and weakest groups should be communicated through a ranked chart, selective emphasis or a stronger table composition.
4. **Data quality:** fail. Six long text lines reproduce an audit note. The evidence should be structured as a small set of visual assurances and exceptions, with detail moved to notes.

## Root cause

Scribble's acceptance logic verified facts, editability, fit, palette and native chart structure. It did not require meaningful visual hierarchy. The renderer allowed analytical slides containing several numbers to use the generic `bullets` recipe, and its vision review prompt treated readable geometry as sufficient.

## Remediation implemented

- Added a `scorecard` composition with two to four prominent KPI values.
- Raised default title and body minimum sizes.
- Rejected numeric analytical slides that use a plain bullets layout.
- Required at least two thirds of a multi-slide executive deck to use charts, tables, source images, scorecards or structured cards.
- Added a deterministic native-output guardrail that fails numeric body-text dumps unless they contain native evidence or at least two prominent metrics.
- Strengthened slide and deck review prompts to reject Word-page compositions, accidental whitespace, weak hierarchy and repetitive layouts.

## External engine assessment

`ui-ux-pro-max` is design guidance for software interfaces, not a slide renderer. PptxGenJS is a capable OOXML construction library, but replacing PowerPoint COM with another low-level writer would not add design judgment.

Presenton is the strongest integration candidate because it provides an Apache-2.0 template engine, structured JSON input, OpenAI-compatible providers including OpenRouter, reusable PPTX-derived layouts and editable PPTX export. A safe Scribble integration should keep Scribble's evidence and arithmetic layer authoritative, send only validated structured slide content to a pinned local Presenton service, disable web search, import the returned slides into an unsaved deck, and run Scribble's native fact and visual checks again before completion.

DeepSlides is worth tracking for a later template-free mode. Its design-first workflow is promising, but its experimental Python pipeline and model requirements make it a higher-risk production dependency than Presenton.
