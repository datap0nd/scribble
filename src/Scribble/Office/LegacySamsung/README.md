# Frozen Samsung workflow 1

These five source files were copied from commit
`57bda70` before workflow 2, with only their namespace changed to
`Scribble.Office.LegacySamsung` and a provenance comment added. The paired
`DocumentDraftHost.LegacySamsung.cs` preserves the original orchestration and
review behavior, with method names and type references isolated from v2.

Only persisted tasks whose SamsungWorkflowVersion is below 2 dispatch here.
New tasks use the current renderer, policy, review and recovery implementation.
Do not silently update these geometry/typography recipes alongside v2 changes.
The user-visible generation tool names remain compatible.
