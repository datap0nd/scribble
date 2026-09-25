# Phase 4 native layout candidate

This is an **unapproved candidate**, produced on the Windows workstation from the CI-built `cc92bce` assembly. It does not certify Phase 4 or PP01. The exact assembly, fixture, source-deck, and output hashes are in `manifest.json` and the two raw native reports. The raw reports retain the workstation's original temporary file paths; the corresponding files are archived here under `reference/` and `defect/`.

The native harness passed 18 structurally checked reference slides and 18 separately rendered seeded defects. Each set contains three editable six-slide PowerPoint decks and matching PDFs. A before/after package copy accompanies each PDF export; the harness accepted only metadata and at most 2 EMU of native table-frame rounding. The defect report binds every mutated slide to the SHA-256 of its unchanged reference deck. The bundled `fixtures/` files are byte-for-byte copies of the CI harness inputs; the repository checkout may have different line endings.

The 18 reference pages cover scorecards, comparison charts and tables, grouped charts, evidence cards, dense tables, and cover/closing layouts at normal, long-label, and maximum supported densities. Each reference and defect page has a full-size PNG. The contact sheets help locate a page, but visual review should use the full-size PNG and editable PPTX/PDF. The defect pages are intentionally broken and must not be approved as output.

`visual_approved=false` and `reviewer=null` are deliberate. The exact references still need identified human approval against the design rubric. Reviewer calibration, six-slide copy-and-patch PP01 acceptance, and the live repeatability ladder remain open. No paid model run was used to create this evidence.
