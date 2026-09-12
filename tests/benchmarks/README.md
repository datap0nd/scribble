# Scribble synthetic testing kit

The kit contains fictional Atlas Office Supplies data, 16 exact prompts and evaluator-only answers. Start with **EX01**, **PP01**, and **XA01**. Their expected June revenue is EUR 120000, budget gap EUR 10000 and margin 38.33%.

## Run the whole suite with one button

Install the latest Scribble installer, restart Office, and make sure the Chrome extension is current. Configure the model you want to measure. Start your screen recording, then click **Test Lab** in any Scribble pane. The button is available without an activation script.

That click opens or focuses one idle window; it performs no model inference and opens no fixtures. Click **Start** to run all 16 cases. Test Lab downloads the latest kit from an exact main commit, verifies its ZIP checksum and all manifest hashes, and opens a fresh case copy in the visible origin/output apps. Supporting CSV, DOCX and PDF files use Scribble's normal context readers instead of becoming unrelated add-in prerequisites. Follow-up cases execute their prerequisites first. XA04 saves the generated deck before requesting its email attachment. RC01 observes a new workbook, stops, then submits the continuation; a missed stop boundary is recorded as blocked.

A live window shows UTC progress and captured errors, including PowerShell stdout/stderr when preparation fails before writing its JSON report. The runner answers questions that explicitly name audience, period, currency or format using only the kit’s predefined operator answers. Missing applications, disabled add-ins, other model questions and timeouts are recorded as blockers. The runner stops that request before moving on. If it cannot confirm stopping, it aborts the remaining cases, retains the synthetic source guard and exports an incomplete evidence snapshot. Stop the request in its app before starting another suite.

Results go to **Documents/Scribble Testcases/suite-<UTC timestamp>-<id>/**:

- `report.pdf`: validated final report with summary, deterministic findings, retained traces, and usable native PDF/PNG output pages.
- `report.html` and `summary.txt`: retained diagnostic derivatives used to build and troubleshoot the final PDF.
- `diagnostics.txt`: complete report text, also available as numbered copyable parts in the HTML.
- `suite.log` and `suite.json`: live log and structured case outcomes.
- `cases/<case>/`: preparation logs, case report, evidence ZIP with original collected Office/email outputs, and that case's isolated fixtures.
- `test-kit.zip`: the verified download; no manual ZIP management is needed.

The window has exactly **Start**, **Stop**, and **View final PDF**. Stop prevents the next submission, requests cancellation, and preserves available evidence while mandatory finalization continues. The PDF renderer is local and does not require Chrome, Office, Python, or a model. A model response or an existing output file is never automatically declared correct. Cases are marked for review, deterministic failure, incomplete, blocked, stopped or not run.

The suite leaves opened apps visible for recording and inspection. Emails are unsent drafts. Existing unrelated files are not saved or closed. Synthetic source files open read-only; the collector can save generated changes from the suite-owned copies, as well as new run-tagged documents. Answer keys remain evaluator-only and are added to the report after execution. Full native Office/Qwen acceptance still needs a run on a machine with the apps, connected add-ins and configured model; infrastructure checks do not certify model output quality.

## Manual operator scripts

The ZIP and the `operator` scripts remain available for individual-case diagnosis and evidence exports. `Enable-ScribbleTestLab.ps1` loads the installed DLL, preferring `%LOCALAPPDATA%/Programs/Scribble/Scribble.dll`. `Prepare-ScribbleTestCase.ps1 -CaseId PP01 -PlanOnly` prints the setup plan without opening apps. The suite supplies `-Suite` to prepare without manual next-step instructions. `Export-ScribbleTestRun.ps1` exports an existing finished run. Do not run a separate manual capture during a suite.

If a script reports `Unable to find type [Scribble.Testing.TestLab]`, install the [current Scribble installer](https://github.com/datap0nd/scribble/releases/latest/download/ScribbleSetup.exe), restart Office, and use its **Test Lab** button. The manual scripts accept `-AssemblyPath` for custom installations.

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

Test Lab opens or focuses one visible standalone window. If the previous run was interrupted, **Start** stays disabled and **Stop** performs an idempotent recheck. Once no model host can still mutate the run, Stop preserves the old capture in a validated PDF and enables a fresh Start. No unrelated application is closed automatically. The report header records the runner build separately from the downloaded kit commit.
