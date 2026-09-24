# Phase 3 native layout reflow evidence

The image is the exact first slide exported by the disposable native analysis
acceptance harness built from commit `59deb88` (Windows CI run `36060850857`).
SHA-256: `a1c794a9d9ad1d44e8738ab22fab1576075df3468cd88f41e92a158def29150d`.

The offline fake reviewer first requested a title correction, then a bounded
`scorecard` to `cards` reflow. The active route applied both to the same
unsaved, task-owned deck and completed native readback and task receipt
reconciliation. A second disposable native test rejected a stale layout
reservation, blocked review after a pending write was reloaded, rejected an
incorrect layout readback, then reconciled the correct native state. The native
run reported `layout_recovery_passed=true`, `source_preserved=true`,
`native_date_column_passed=true`, and `powerpoint_exited=false`.

This image is evidence of the functional route, **not an approved visual
baseline**. It has excessive empty space, small KPI emphasis, and a repeated
revenue value. The native report still records `full_acceptance_passed=false`.
No paid model call was made.
