# Samsung presentation workflow v2

Implementation status: code and automated validation implemented; native acceptance pending. The reference is
`SAMSUNG_SLIDE_DESIGN_SPEC.md`. Real Samsung reference decks have not been evaluated;
do not describe the candidate as having verified fidelity to those decks.

## Authoring and generation

`SamsungAuthoringPolicy` is the versioned source for authoring, evidence and review
instructions. The PowerPoint pane and cross-application generation share the tool
catalog. Existing generation tool names and ID-list payloads remain accepted.
New tasks use workflow 2; persisted tasks without this version retain the legacy
review path. `Office/LegacySamsung` and `DocumentDraftHost.LegacySamsung.cs` freeze
the previous renderer, typography, authoring evidence checks and review orchestration.
New tasks use the v2 implementation.

New requests should supply slide briefs with purpose, message, evidence references,
recipe and required content. The outline is reviewed before native writes. Exact
requested counts include covers and appendices. Detailed evidence stays visible;
density conflicts produce a blocker rather than silently dropping rows.

The policy separates subject titles, analytical action subtitles and optional
takeaways. Cover, divider, agenda and explanatory exceptions are explicit. Claims
carry evidence associations, labels, units and periods. Host arithmetic validates
cited operands, supported formulas, result units and rounding. Independent factual
review checks semantic associations; numeric matching alone is insufficient.
Sample authorization is scoped to marked sample slides.

Recipes use a 960 x 540 canvas, 24 pt titles, 14 pt action titles, 18 pt body text
with a 14 pt floor, and tables starting at 10 pt with a 7.5 pt floor. Existing
16:9 canvases are scaled. Other aspect ratios support targeted edits only;
Samsung recomposition requires a separate 16:9 presentation. Native tables and
charts, missing chart values, source images and bilingual text are supported.
Annotations address primary/secondary tables or charts and specific cells or
series. Source notes are appended, preserving existing notes.

## Review and completion

Vision capability is checked before writes. Factual review, native geometry,
individual previews and whole-deck review are separate checks. Generated slides
get an initial visual review and at most three defect-directed repair attempts.
Repair must preserve table/chart evidence; repeated findings or unchanged proposals
stop the loop. Repair budgets persist across interruptions. Continuation pages are
reflowed and reviewed together, preserving the complete source data and allocated
page count. A repair that cannot fit explains the exact count conflict.
The final review includes native preview contact sheets and narrative/coverage
checks. Whole-deck findings route affected generated slides back through fact,
layout and rendered review, including slides from earlier generation batches.
Previously approved native content is fingerprint-checked before reuse.
A final receipt is required for task completion. Review caches include
policy, renderer, model, endpoint, content and evidence inputs.

Private PNGs support verification. No presentation Save/SaveAs or user-facing
PPTX/PDF export capability is added. Internal findings do not belong in slides or
speaker notes.

## Inspection and revision

`inspect_slide` returns paginated native content, stable presentation/slide/shape
identifiers, fingerprints, geometry, tables, chart data, notes, groups, styling,
unsupported-object reports and a private preview.

`revise_slides` is limited to 24 operations. Supported operation names are
`replace_text`, `table_cell`, `chart_point`, `annotate`, `notes_append`,
`replace_slide`, `insert`, `move` and `delete`. Structural actions require an
explicit request. Target IDs and original fingerprints are mandatory. The host
reviews an unsaved native working copy before applying a journaled batch.
Failed staged visual, geometry or deck checks receive at most three repair cycles.
Repairs preserve operation count/order, targets, original fingerprints, requested
numeric changes and evidence; a correction needing broader authority stops with
a specific scope conflict.
Recomposition is rejected when it would lose protected artwork, actions,
animations, hyperlinks or a non-solid background. Deletion is rejected when
custom slide shows or incoming slide hyperlinks cannot be preserved. Native chart
point edits validate a direct local SERIES range and update its embedded constant
source cell, then verify the formula and displayed value. External, computed,
named, disjoint and multidimensional ranges are explicitly unsupported.

`revert_scribble_changes` restores the latest batch from native originals in an
unsaved recovery presentation. Both apply and revert check concurrent changes.
Restoration is verified before reporting success. Uncertain recovery is reported
explicitly and exposes the surviving recovery presentation. These are session
recovery facilities, not recovery after PowerPoint closes. Restoring a deleted
slide can create a new slide ID; inspect again before another edit.

Generation persists the original payload and each rendered native slide receipt.
On resume, the same generation tool and unchanged payload reconcile surviving IDs,
order and fingerprints before rendering missing pages. An unrecorded partial native
write or later user change stops automatic mutation and preserves the draft.

Revision journals persist operation state and references to tagged unsaved recovery
presentations. A surviving completed batch is verified and reviewed without replay.
A partial batch is rolled back only when live state and recovery originals match
recorded fingerprints. Recovery conflicts remain explicit blockers. Original
payloads are available through paginated task evidence for continuation. Existing
Office source/session binding checks still apply; closing Office ends unsaved
recovery guarantees.

Revision tools remain absent from the model catalog until native acceptance
passes for the exact Scribble assembly SHA-256 and policy version. The receipt
is `%LOCALAPPDATA%/Scribble/PowerPointAcceptance.json`. Rebuilding invalidates it.
Do not create the receipt by hand.

## Validation and release

Build the solution in Release and run `tests/GuardrailTests/bin/Release/GuardrailTests.exe`.
Run `scripts/Test-Guardrails.ps1` and
`tests/ReleaseGate/Test-ReleaseEvidence.Tests.ps1` in PowerShell 7.

On a machine with working desktop PowerPoint, run:

```powershell
./tests/NativeAcceptance/Test-PowerPointWorkflow.ps1 -OutputPath ./tmp/PowerPointWorkflow.json
```

The harness creates unsaved test presentations, exercises revision operations,
concurrent-edit rejection, rollback and revert, and checks native preservation.
It never terminates the user's PowerPoint process. After inspecting a passing
report, `-EnableRevision` installs the matching receipt. This is an engineering
capability gate, not evidence that every model/UI benchmark has passed.

`tests/NativeAcceptance/SamsungBenchmarks.json` defines the five benchmark cases
and measurements. Run them from PowerPoint, Outlook, Excel, Word and Chrome with
identical inputs/model configuration against baseline and candidate. Record real
results; synthetic validator fixtures are not acceptance evidence. Workflow-2
release candidates require the additional native scenario evidence, exact content
coverage, editability, no unintended changes and no presentation saving/export.

The implementation session passed the Release build and full regression suite,
including deterministic interrupted-generation and revision-recovery tests, repair
scope/value checks, chart source mapping and the frozen legacy renderer. Static
save/export guardrails and release-evidence validator fixtures also passed.
The exact final test count is recorded in the implementation task.

## Native acceptance pending

Local desktop PowerPoint is not properly installed, as confirmed by the user.
Native PowerPoint activation previously failed with
`0x80080005 (CO_E_SERVER_EXEC_FAILURE)`. No native acceptance receipt is installed,
and native revision tools remain gated. This is a deployment/acceptance condition,
not an unfinished implementation item.

Complete the native/model/UI benchmarks and representative real-deck review on a
machine with working PowerPoint before enabling revisions or claiming native
acceptance. The native harness covers embedded chart source associations, original
preservation, all supported revision operations, snapshot reconciliation, rollback
and revert. Broader media, animation, master/theme and bilingual fixtures remain
part of the benchmark acceptance checklist. No performance or human Samsung-fidelity
scores have been measured locally; synthetic tests never count as native evidence.
