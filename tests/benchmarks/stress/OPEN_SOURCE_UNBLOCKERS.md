# Open-source candidates for the Office reliability path

Researched 23 September 2026. This is a dependency evaluation, not a change to
the supported Scribble runtime or the Phase 3 gate. No paid model call is needed
for the trial. Keep the public 2.0.91 release frozen.

## Best candidate: OfficeIMO.PowerPoint

[OfficeIMO.PowerPoint](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.PowerPoint/README.md)
is a managed .NET library that loads, edits, and saves existing presentations.
Its documented API covers editable charts, tables, notes, semantic slide
families, feature inspection, and image export. This overlaps work in
`PresentationDraftWriter`, `SamsungSlideDesign`, and `PresentationRevision`.
Its feature report separates editable, partially editable, preserved, and
unsupported package content. It could be especially useful as a preflight for
PP01's source-preserving edits. The library's own
[visual proof and designer APIs](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.PowerPoint/README.md#designer-composition)
also expose semantic deck plans, rhythm inspection, corporate-template
inventory, and an opt-in PowerPoint Desktop reference render. These could
replace some custom layout bookkeeping if their output meets Scribble's native
quality bar. The library's
[image-export matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.image-export-capability-matrix.md)
describes incomplete representative authored-deck baselines and rendering
fallbacks, so its image output cannot replace PowerPoint's native render oracle
without comparison. The source deck may contain objects that a library can
preserve but not safely edit.

An isolated `tests/OfficeImoProbe` project and optional PR CI job perform the
first synthetic no-op and one-text-edit package round trip. They do not add
OfficeIMO to the add-in or certify visual fidelity. The broader trial before
adding a production dependency is to pin the verified
[NuGet 3.4.2 package](https://www.nuget.org/packages/OfficeIMO.PowerPoint/3.4.2)
in an isolated test project, check that it loads under Scribble's .NET
Framework 4.8 target (NuGet lists a 4.7.2 asset), and run the existing
synthetic and a PP01-like disposable deck through
`InspectFeatures()`. Make one local text edit and one chart-data edit in copies.
Compare slide and shape identities, notes, charts, embedded workbook data,
relationships, media, artwork, and untouched package parts before and after.
Open and render both copies in PowerPoint; compare native geometry and human
reviewed images. Accept it for a narrow operation only if it passes the same
readback and preservation checks as the current COM writer. Record any
unsupported feature as a capability refusal, not a silent conversion.

**First trial result.** [CI run 35837501877](https://github.com/datap0nd/scribble/actions/runs/35837501877)
built and ran OfficeIMO.PowerPoint 3.4.2 on `net48` against the four-slide
synthetic deck. Both the no-op save and one-title edit retained all 97 package
parts. OfficeIMO reserialized 40 parts on the no-op save, including slide,
chart, note and relationship XML; that is why package-byte equality is not
the preservation oracle. PowerPoint opened all three files read-only. Its
native 1920×1080 PNG exports for the no-op copy were byte-identical to the
original on all four slides. The edit changed only slide 2's native render.
Extracted slide text, note text, and chart cache values were identical between
the original and no-op copies; the edit changed only slide 2's text sequence.
The enhanced feature report in
[CI run 35837988465](https://github.com/datap0nd/scribble/actions/runs/35837988465)
classified 39 text boxes, two tables, and four speaker notes as editable, but
the one imported chart as **partially editable**. Chart-data mutation therefore
needs a separate operation-specific test before adoption.
This is a promising **synthetic** result. It does not cover a dense PP01-like
source deck, animations, embedded objects, or an in-place edit of an unsaved
Office session. The next trial needs those cases and the feature report's
individual findings, not just part counts.

## Alternatives and ideas worth borrowing

| Candidate | Useful part | Limit for Scribble | Decision |
| --- | --- | --- | --- |
| [ShapeCrawler](https://github.com/ShapeCrawler/ShapeCrawler) | C# API for existing PowerPoint shapes, tables, and charts | Rendering and preservation breadth still need the same native trial | Reserve if OfficeIMO fails a bounded edit |
| [office-kit/pptx](https://github.com/office-kit/pptx) | Reads and writes decks, native charts, notes, previews, and unknown-part retention | Adds a JavaScript runtime boundary to the .NET Framework add-in; pre-1.0 API | Keep as a second engine candidate |
| [hands-on-deck](https://github.com/EveryInc/hands-on-deck) | Atomic JSON patches, inspect/diff/lint/render workflow, layout measurement | Python command-line runtime; native chart creation is outside its stated scope | Borrow its patch and verification workflow, not the whole runtime |
| [pptx-automizer](https://github.com/singerla/pptx-automizer) | Existing-template composition and modifications | Its [limitations](https://singerla.github.io/pptx-automizer/limitations) say an edit to one element requires including all other slides in the output process; poor fit for preserving unrelated PP01 content | Lower priority |
| [SlideForge](https://github.com/UIUC-MONET/SLIDEFORGE) | Localized deck-state graph and visual diff ideas | Heavy vision/runtime and inference cost relative to this pilot's remaining budget | Borrow inspection ideas only |

The [Open XML SDK](https://learn.microsoft.com/en-us/office/open-xml/about-the-open-xml-sdk)
is a useful independent package validator. Package validity alone does not
establish visual fidelity, correct business values, or safe edits to an open
unsaved Office session. Scribble's typed source contract, task-level receipts,
and native readback remain necessary whichever slide library wins.

## Measurable unblock decision

The immediate bottleneck is review and repair ownership. A model can identify
a suspected issue, but it cannot be allowed to replace verified numbers,
rewrite an entire slide for a folio defect, or claim that a native write is
correct without readback. Phase 3 routes facts to deterministic calculation,
page/geometry fixes to the renderer, and bounded wording changes to a
field-scoped patch. A slide library can reduce the amount of custom OOXML/COM
work after it passes the preservation trial; it does not replace that contract.

Record the trial outcome in a later gate with the pinned version, input/output
hashes, feature report, native readback, and images. Until then, do not add the
dependency or claim that it solves PP01.
