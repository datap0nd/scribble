# Phase 5 Excel execution checkpoint

Status: native component checks passed on the development branch; the Phase 5 exit gate remains open until the final CI head and supported-scope review are recorded. This checkpoint does not authorize paid inference or a public release.

## Supported scope under test

| Surface | Current bounded behavior | Boundary |
| --- | --- | --- |
| `write_draft_sheet` | Appends a marked `Scribble Draft` sheet to the workbook captured when the user submitted the request, or creates a new workbook when none was captured. A focus change cannot redirect it. | A failed draft-sheet build may leave a visibly incomplete new sheet. It does not edit an existing sheet or save the workbook. |
| `write_cells` | Writes at most 200 rows, 30 columns per row, and 500 characters per cell into the request-bound worksheet. Every destination is checked for bounds, protection, and merging before permission is consumed and immediately before mutation. | A renamed, saved-as, closed, protected, or merged target fails closed. The host does not claim a native COM transaction. |
| Formula cells | Allowed local formulas are written as native formulas and checked after recalculation. Syntax failures and native Excel error results trigger restoration. | The existing `DraftFormulaPolicy` rejects external or unsafe functions; unsupported formulas are not certified as live calculations. |
| In-process failure | Captures each original value, formula, and number format. On failure, restores only cells that still match Scribble's observed write. | Concurrent or unreadable cells remain uncertain and block the task from another document write. |
| Process interruption | An encrypted receipt contains workbook and sheet identity, call ID, each typed before-image, and planned text. Resume compares all cells: all original means rolled back; all planned means applied; a verified mixture is restored; divergence remains uncertain. | The original workbook must still be open under the same identity. A renamed or unavailable workbook is not guessed from the active window. Unsaved work is never silently saved. |

## Evidence

- CI-built `32d46bc` from [run 36011129279](https://github.com/datap0nd/scribble/actions/runs/36011129279) passed the disposable native Excel harness. The report records target binding, merged-cell preflight, in-process rollback, formula-error rollback, a recalculated absolute/relative formula, and restart reconciliation as true. Its `full_acceptance_passed` remains false while the final Phase 5 gate is open. Binary and report hashes are in `evidence/phase5-bound-target-candidate/manifest.json`.
- The native restart checks use a numeric Excel date serial with a date format, move focus to another workbook, and exercise untouched, partially written, fully written, and user-edited cells. Partial writes restore the original date and text. User edits remain untouched and the write stays pending.
- Offline tests inject a later-cell write failure, restart readback failure, restart restoration failure, workbook identity change, and an uncertain task receipt. The CI-built `3439954` binary passed all 213 guardrail tests locally; the repository static guardrail scan passed. The final head's CI result is pending.

## Remaining gate

Record a green full CI run for the final Phase 5 head and its native report. Review the supported scope above, then update the manifest and PR status consistently. Phase 3 and Phase 4 have separate open gates; this checkpoint makes no claim about the hosted model pilot, visual approval, or broader workbook capabilities.
