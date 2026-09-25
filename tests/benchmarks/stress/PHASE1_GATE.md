# Reliability phase 1 offline gate

Phase 1 source commit: `0256e8e9b93d704821edebb3624775ddacdfe855` on
`codex/reliability-phase1` ([draft PR #23](https://github.com/datap0nd/scribble/pull/23)).
The full Windows CI run
[`35827271350`](https://github.com/datap0nd/scribble/actions/runs/35827271350)
completed successfully, including the solution build, guardrail executable,
static capability scan, browser fixtures, and installer checks.

The offline guardrails exercise distinct identities for identical source
instances, changed-revision invalidation, source-span persistence, typed
attached and live workbook capture, sparse `.xlsx` and `.xlsb` column
alignment, formulas, formats, dates, blanks, malformed values, deterministic
calculations, and analysis serialization. Cached formula results remain
unresolved until native recalculation and are excluded from authoritative
grouped totals.

This is the phase 1 **offline** exit gate. Native Office behavior, generated
output quality, and the acceptance run ladder remain unproved. No paid model
inference or Office write was used for this gate. Phase 2 must prove that a
single analysis creates new, editable Excel and PowerPoint drafts with native
readback against an independent oracle.
