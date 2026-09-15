# Scribble synthetic stress corpus

This suite contains 200 distinct tasks against 20 fictional Excel workbooks,
30 fictional PowerPoint presentations and 500 fictional Outlook messages. The
case catalog is generated from independent source calculations. Generating a
corpus is not a model run and does not establish model quality.

The four groups are 60 Excel cases (monthly formulas and charts, group audits,
isolated what-if calculations), 70 Outlook searches (500-message pagination,
folder/date/unread boundaries, exact sender and project identities, current
decisions and long-body reads), 50 PowerPoint cases (30 repairs and 20 new
reviews), and 20 output handoffs between Office applications.

## Build and seal

Use the bundled Python and Node runtimes. First run `generate_office.py` with
the runtime paths described by its help. Complete native Office-file authoring
and verification; `--data-only` is useful for development but cannot produce a
runnable kit. Then run:

```powershell
python tests/benchmarks/stress/generate_mail.py
python tests/benchmarks/stress/build_catalog.py --manifest
python tests/benchmarks/stress/test_catalog.py
```

Generated sources, evaluator references and output are ignored by Git under
`tests/benchmarks/generated/stress-corpus`. The final command reports the exact
`manifest.json` SHA256. The seal refuses missing case inputs, missing reference
themes or an incomplete 500/20/30 corpus. Rebuilding native files invalidates
the seal: regenerate mail attachments and seal again before starting a run.

## Run through the normal Test Lab

Use the installed BrowserHost operator entry point with the exact generated
manifest hash and a new absolute result path:

```powershell
$corpus = (Resolve-Path tests/benchmarks/generated/stress-corpus).Path
$manifestHash = (Get-FileHash -LiteralPath (Join-Path $corpus 'manifest.json') -Algorithm SHA256).Hash.ToLowerInvariant()
& "$env:LOCALAPPDATA\Programs\Scribble\ScribbleBrowserHost.exe" --test-lab-run `
  --kit $corpus --kit-sha256 $manifestHash --result-json 'C:\ScribbleRuns\stress-attempt-001.json'
```

The entry point opens the same visible operator window, acquires the normal
suite/session ownership, and uses the configured model through Scribble's real
Office panes. The selected kit is verified and copied into the results folder
once. Each case gets an isolated projection containing only its source inputs,
attachment closure, operator metadata and private oracle. The native synthetic
Outlook store is bound to the verified full snapshot; it does not search the
user's default mailbox. Missing installation, Outlook profile, native store,
model configuration or budget authorization produces an explicit blocked case.

The stress budget gate checks the dedicated provider key before each case and
after each terminal case. Native provider limits enforce the authorized cap.
An interrupted or budget-stopped attempt retains its results. Continue by
selecting the remaining IDs in a new attempt, for example:

```powershell
& "$env:LOCALAPPDATA\Programs\Scribble\ScribbleBrowserHost.exe" --test-lab-run `
  --kit $corpus --kit-sha256 $manifestHash --cases 'EX04,OL01,PP02,XA20' `
  --result-json 'C:\ScribbleRuns\stress-attempt-002.json'
```

`--case EX01` selects one case. `--cases` preserves explicit order and refuses
duplicates or absent IDs. Neither option automatically replays completed work.
Omitting `--kit` continues to run the existing bundled 16-case suite.

For a single launch, run `Run-StressSuite.ps1`. It pins the corpus manifest and
starts all 200 cases. `-OpenReport` opens the final PDF, including an explicit
blocked report if no cases could run. `-ResumeFrom <previous result.json>` selects only cases
marked `not_run` in the same corpus; earlier failures remain recorded. Supply
`-ExpectedOfficeBuild` and `-ExpectedOfficePlatform` once the work PC versions
are known. Without both, environment metadata explicitly records Office parity
as unverified. `summarize_usage.py <suite.json>` extracts provider token counts
and reported costs from recorded inference responses without making API calls.

The result JSON publishes the current folder and stop token. Writing that token
to its `stop_file` requests the usual bounded stop and evidence finalization.
A completed harness requires every requested case to reach a terminal result
and a validated PDF. Exit code 3 means the harness completed with deterministic
correctness failures; 2 means blocked/incomplete; 4 means stopped. A completed
harness remains subject to native and visual review.

## Evaluation boundary

Only files beneath `inputs/` may enter the model's document context. Each
operator case points to a private `evaluator-only/cases/<ID>.json` oracle.
Oracle values, expected mail IDs, corrected references and formula probes never
enter prompts or mail attachments. Output cell addresses, requested formats,
scenario input changes and source identifiers are legitimate task instructions.

The native evaluator checks recalculated formula values in named output cells,
actual chart categories and series, preserved sources, native draft mail
headers and labeled metric values, exact returned mail ID sets and complete
search cursors, and native slide count, table fills, chart colours and text fit.
Incomplete rates must remain explicit text rather than invented percentages.
It must reject unsupported checks rather than
silently count them as passed. The mail prompt requires `Total matches: N` to
make zero matches and incomplete result sets unambiguous.

`Test-SyntheticMailbox.ps1` imports the verified 500-message corpus and exercises
all 70 query scopes through the production Outlook adapter without a model.
`Test-StressNativeGrading.ps1 -ScribbleAssembly <absolute Scribble.dll path>`
checks the presentation grader against a clean native reference and isolated
colour/layout defects. These are native adapter and grader preflights; neither
is a Qwen evaluation or a substitute for running the 200 tasks.

Formula probes retain independently calculated expected results after a source
input change. Running those probes on throwaway copies and examining visual
hierarchy, semantic missing-data handling, attribution and proposed actions
remain explicit review items. A formula string or an answer-number match alone
does not establish dependency correctness. Keep the full native evidence,
individual evaluations and PDF alongside any review screenshots.
