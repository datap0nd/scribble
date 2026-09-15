# Scribble synthetic testing kit

Click **Test Lab**, then **Start**. The default scope runs **16 Excel, PowerPoint
and Outlook tests** against your configured model. Keep `qwen3.8-27b-fast`
selected to measure Qwen. No other AI judges or changes its answers during a run.

The fictional Atlas Office Supplies dataset has known results: June revenue
EUR 120000, budget EUR 130000, gross profit EUR 46000, weighted margin 38.33%,
and delivery 94% against a 97% target. The runner submits authored prompts and
only authored clarification answers. Expected answers are used after capture.

## One-button workflow

1. Install a tested development installer from the repository's Actions
   artifacts over the existing installation and restart Office. Public updates
   are frozen at 2.0.91; they do not deliver the development Test Lab fixes.
   Configure and test the model connection in Settings.
2. Open **Scribble Test Lab** from Start or the **Test Lab** button in a pane.
   Opening it makes no model request. Leave Case ID blank for the 16-case run.
3. Click **Start**. The runner verifies the kit embedded in that exact build,
   opens fresh read-only fixtures, runs the normal Scribble panes, captures
   native results and evaluates them. Word and Chrome are optional historical
   cases, available only by explicitly entering their case IDs.
4. Click **View final PDF**. The ten-page PDF reserves two case cards per page,
   including expected facts, observed answers/formulas, checks and available
   native slide previews. Screenshot those pages for review; full outputs,
   all slide images and complete traces remain in the adjacent HTML/evidence ZIPs.

The catalog contains 19 cases, of which 16 belong to the default Office scope.
The extra Office cases test weighted margins (EX05), chart scales and units
(PP04), and a grounded unsent email (OL02). Follow-ups retain their prerequisites.
RC01 still observes the workbook boundary before stopping and continuing.

Office preparation is native C# and retains all participating applications for
the suite lifetime. Outlook inputs are verified local MSG fixtures read by the
normal message and attachment readers; no PST or mailbox import is required.
Native Office exports use short paths under **Documents/Scribble Testcases/Native**,
because Office can be denied saves inside LocalAppData even when the runner can
write there. If native saves fail, live readback remains distinguishable from a
saved editable file; the report identifies missing or incomplete evidence.

**Stop** prevents further submissions, requests cancellation and preserves a
partial report. After an interrupted run, **Start** verifies that the prior
request stopped, reconnects its recorder and preserves incomplete evidence
before starting fresh. A stale or dead pipe cannot prevent recovery. If a live
host cannot acknowledge cancellation, its identity is shown and later cases
remain unsubmitted. No unrelated application or user document is force-closed.

Results are under **Documents/Scribble Testcases/suite-<timestamp>-<id>/**, with
a writable LocalAppData fallback if Documents is unavailable:

- `report.pdf`: concise review of all 16 cases, up to ten pages.
- `report.html`, `diagnostics.txt`, `summary.txt`: complete copyable diagnostics.
- `suite.log`, `suite.json`: progress, failures and structured outcomes.
- `cases/<id>/`: verified fixtures, native evidence ZIP and deterministic checks.
- `test-kit.zip`: the kit from the installed build, with its checksum recorded.

A completed run is not automatically a correctness pass. Deterministic checks
reject missing outputs, native formula errors and missing expected facts, and
compare preserved source content. Chat-only checks inspect the final assistant
answer, not numbers already present in input events. Formula semantics, charts,
grounding and slide layout still require review. Tests never send email.

If Office loses its connection during a case, the runner preserves that failure
and checks the exact submission receipt and Office process identity. A completed
request or confirmed process exit releases the following case; a live request
without confirmation stays blocked. A separate stop signal lets the pane cancel
even when COM status calls fail; it targets that exact submission so RC01 can
resume. Reopened apps receive fresh connections. Outlook can also attach through
its registered automation class when its running instance is absent from ROT;
capture and cleanup reuse the retained application instead of rediscovering it.

Completed answers pass through the recorder in both pane types. Blocked questions
and rejected PowerPoint arguments appear in the report. Source capture records
the original worksheet names and slide IDs, so draft labels already present in
the starter cannot turn an untouched fixture into a generated result.

## Verification boundary

### Run the configured model from a local terminal

The installed browser host also has an explicit operator command. It opens the
same visible Test Lab window and invokes the same Start workflow, including stale
capture recovery, Office preparation, normal Scribble permissions and final PDF:

```powershell
$hostExe = Join-Path $env:LOCALAPPDATA 'Programs\Scribble\ScribbleBrowserHost.exe'
$result = Join-Path $env:TEMP ('scribble-run-' + [guid]::NewGuid().ToString('N') + '.json')
& $hostExe --test-lab-run --result-json $result
Get-Content -LiteralPath $result -Raw | ConvertFrom-Json
```

Add `--case EX01` to run one of the 16 Office cases. The result path must be a new
absolute `.json` file. The command refuses an existing Test Lab window or live
suite; it never closes or replaces another operator's window. Configure the model
normally in Scribble Settings before running; this command does not select a
different model.

The result JSON first reports `starting`, then `running` with the exact suite
folder. Read `suite.log` there for progress. The command exits only after the
current report is finalized and publishes the PDF path and all case statuses.
Exit `0` means all requested cases completed and the PDF is valid, with model
quality still needing review; `3` means the run completed but deterministic
correctness checks failed; `2` means a blocked/incomplete run or report failure;
`4` means stopped; `1` means the operator launch or arguments failed.

To stop from a second local terminal, read the result JSON and write its exact
`stop_token` to its `stop_file`. This requests the same Stop action as the visible
button and waits for evidence/PDF finalization. Do not terminate the host process
to stop a test.

```powershell
$state = Get-Content -LiteralPath $result -Raw | ConvertFrom-Json
[IO.File]::WriteAllText($state.stop_file, $state.stop_token)
```

`Test-TestLabReliability.ps1` reproduces dead-recorder recovery and restart,
checks run-isolated message IDs, rejects correct-input/wrong-answer evidence,
and generates a complete 16-case PDF without Office or a model. The optional
`Test-TestLabNative.ps1 -RunOffice` opens installed Office apps and manufactures
known outputs to check the native capture boundary. These smoke outputs are
explicitly not Qwen outputs. A successful real Qwen run on the work machine is
required before claiming model quality or full work-environment acceptance.

For a local recovery check, add `-InjectOfficeExit`: after capturing PP01 and
closing its synthetic documents, the smoke terminates only the empty PowerPoint
process it created, then runs PP03 with the same environment. `-ReuseOutlook`
allows an existing Outlook session for local MSG inputs and an unsent draft;
the smoke never quits Outlook. `Test-TestLabSuite.ps1` also checks cancellation
without COM, live-versus-exited executor ownership, and both real pane transcript
entry points without contacting a model.

## Manual operator scripts

The ZIP and the `operator` scripts remain available for individual-case diagnosis and evidence exports. `Enable-ScribbleTestLab.ps1` loads the installed DLL, preferring `%LOCALAPPDATA%/Programs/Scribble/Scribble.dll`. `Prepare-ScribbleTestCase.ps1 -CaseId PP01 -PlanOnly` prints the setup plan without opening apps. The suite supplies `-Suite` to prepare without manual next-step instructions. `Export-ScribbleTestRun.ps1` exports an existing finished run. Do not run a separate manual capture during a suite.

If a script reports `Unable to find type [Scribble.Testing.TestLab]`, install a tested [development Actions artifact](https://github.com/datap0nd/scribble/actions/workflows/build.yml), restart Office, and use its **Test Lab** button. The manual scripts accept `-AssemblyPath` for custom installations. See [release channels](../../docs/release-channels.md) for the stable release freeze.

## Evidence and evaluation

The HTML starts with case/run/build/model identity, UTC timestamps, capture completeness, missing output types and initial error highlights. It includes prompts, expected results, video markers, full error events, model responses, receipts, available slide images, native text previews and the complete recorded timeline. Spreadsheet previews show all recorded rows with formulas and cached values; these are not recalculated or proof of native correctness.

Use **View final PDF** to inspect or share the displayed terminal run. Errors and event bodies are retained, and native Office PDF pages or PNG derivatives are appended when capture succeeded. Original outputs remain locally available for deeper inspection; deterministic checks do not replace native visual review.

Full permitted request and response bodies, tool calls/results, returned provider reasoning, source hashes, handoffs and timestamps are captured locally in encrypted event files. No API key or Authorization header is intentionally captured. The trace contains exactly the context sent through Scribble, which may already be bounded by normal extraction limits. Provider reasoning is available only if the endpoint returned it.

Enable the lab only in a dedicated synthetic session. Guards check case-specific file hashes, selected mail identity/body, saved Office sources and the loopback browser URL; newly created Office drafts carry the run ID. The operator must also start a clean chat and avoid other Scribble activity during capture. This is not an isolation boundary against a compromised local account.

The suite runs deterministic checks automatically against final run-owned outputs and writes `evaluation.json` per case. Sources and prerequisite captures cannot satisfy a final-output rule, and a correct chat claim cannot mask a wrong workbook or deck. The kit's Python evaluator remains available for independent review and comparison. Complete native recalculation, source preservation, unsent state and visual review before marking an attempt passed. Artifact presence alone never means pass. Videos and evidence are private outputs and should not be committed to the repository.

The evaluator-only reference files are for human comparison. Never add them to Qwen context. The fixture server never exposes them. Run each priority case three times and preserve all failures. After a fix, compare the same prompts/build/model settings, then use fresh synthetic variants before claiming general improvement. This kit does not train model weights or certify the broader release gate.

## Rebuild in the repository

Use the bundled Python and Node runtime with `python-docx`, `reportlab`, and `@oai/artifact-tool`. Link the generator's `node_modules` to the bundled package directory. Run `generators/build_kit.py`, then `generators/build_office.mjs`, native/render QA, `generators/build_kit.py --pack`, and `evaluator/test_kit.py`. The committed ZIP is independently downloadable under `tests/benchmarks/releases/` on main.

Expected outputs include formula-based reference analysis, a six-slide editable deck, a memo, a draft email, operations PDF, and slide images. All amounts exclude tax. Margin changes are percentage points. Exact expected facts and the scoring rubric live under `evaluator-only/`.

## Recover an unfinished capture

Use **Start** to preserve a stopped interrupted capture and begin a new run,
or **Stop** to preserve it without starting tests. Recovery uses a fresh recorder
and direct native cancellation acknowledgement. If an older loaded add-in lacks
that acknowledgement, stop its request or close that specific app, then retry.
The runner never replays a model submission during recovery.
