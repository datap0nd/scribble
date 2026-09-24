# Paired KPI native render

`typed-kpi-pair.png` is the exact first slide exported by the disposable
native analysis harness built from commit `c5d0afd` (Windows CI run
`36062241184`). SHA-256:
`3a346e2bc26afae755f48c5392defe68ad116337c337c24a33ccc46385072ad6`.

The active offline fake reviewer requested a title correction followed by a
bounded `scorecard` to `cards` reflow. Both verified KPI values remain native
editable text, and the exact source footer remains visible. The same run
reported `layout_recovery_passed=true`, `source_preserved=true`,
`native_date_column_passed=true`, and `powerpoint_exited=false`.

This is a visual candidate, not an approved baseline. The native report keeps
`full_acceptance_passed=false`; an identified human has not yet attested the
visual output, and the broader Phase 4 and release gates remain open. No paid
model call was made.
