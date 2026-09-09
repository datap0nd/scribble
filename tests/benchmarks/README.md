# Scribble synthetic testing kit

The kit contains fictional Atlas Office Supplies data, 16 exact prompts and evaluator-only answers. Start with **EX01**, **PP01**, and **XA01**. Their expected June revenue is EUR 120000, budget gap EUR 10000 and margin 38.33%.

## Run the whole suite with one button

Install the latest Scribble installer, restart Office, and make sure the Chrome extension is current. Configure the model you want to measure. Start your screen recording, then click **Test Lab** in any Scribble pane. The button is available without an activation script.

That click starts all 16 cases. Test Lab downloads the latest kit from an exact main commit, verifies its ZIP checksum and all manifest hashes, and opens a fresh case copy in the visible apps. It selects the synthetic Outlook messages, opens the Chrome fixture site and a visible extension controller, loads the source context, and submits the predefined prompts through the normal Scribble chat flow. Follow-up cases execute their prerequisites first. XA04 saves the generated deck before requesting its email attachment. RC01 observes a new workbook, stops, then submits the continuation; a missed stop boundary is recorded as blocked.

A live window shows UTC progress and captured errors, including PowerShell stdout/stderr when preparation fails before writing its JSON report. The runner answers questions that explicitly name audience, period, currency or format using only the kit’s predefined operator answers. Missing applications, disabled add-ins, other model questions and timeouts are recorded as blockers. The runner stops that request before moving on. If it cannot confirm stopping, it aborts the remaining cases, retains the synthetic source guard and exports an incomplete evidence snapshot. Stop the request in its app before starting another suite.

Results go to **Documents/Scribble Testcases/suite-<UTC timestamp>-<id>/**:

- `report.html`: final summary, full errors, timestamps, per-case traces and available output previews.
- `summary.txt`: use **Copy summary** to paste this into chat; the HTML summary is also suitable for a screenshot.
- `diagnostics.txt`: complete report text, also available as numbered copyable parts in the HTML.
- `suite.log` and `suite.json`: live log and structured case outcomes.
- `cases/<case>/`: preparation logs, case report, evidence ZIP with original collected Office/email outputs, and that case's isolated fixtures.
- `test-kit.zip`: the verified download; no manual ZIP management is needed.

Use **Stop suite** to stop and export what is available. The final HTML is generated without a PDF renderer. A model response or an existing output file is never automatically declared correct. Cases are marked for review, incomplete, blocked, stopped or not run.

The suite leaves opened apps visible for recording and inspection. Emails are unsent drafts. Existing unrelated files are not saved or closed. Synthetic source files open read-only; the collector can save generated changes from the suite-owned copies, as well as new run-tagged documents. Answer keys remain evaluator-only and are added to the report after execution. Full native Office/Qwen acceptance still needs a run on a machine with the apps, connected add-ins and configured model; infrastructure checks do not certify model output quality.

## Manual operator scripts

The ZIP and the `operator` scripts remain available for individual-case diagnosis and evidence exports. `Enable-ScribbleTestLab.ps1` loads the installed DLL, preferring `%LOCALAPPDATA%/Programs/Scribble/Scribble.dll`. `Prepare-ScribbleTestCase.ps1 -CaseId PP01 -PlanOnly` prints the setup plan without opening apps. The suite supplies `-Suite` to prepare without manual next-step instructions. `Export-ScribbleTestRun.ps1` exports an existing finished run. Do not run a separate manual capture during a suite.

If a script reports `Unable to find type [Scribble.Testing.TestLab]`, install the [current Scribble installer](https://github.com/datap0nd/scribble/releases/latest/download/ScribbleSetup.exe), restart Office, and use its **Test Lab** button. The manual scripts accept `-AssemblyPath` for custom installations.

## Evidence and evaluation

The HTML starts with case/run/build/model identity, UTC timestamps, capture completeness, missing output types and initial error highlights. It includes prompts, expected results, video markers, full error events, model responses, receipts, available slide images, native text previews and the complete recorded timeline. Spreadsheet previews show all recorded rows with formulas and cached values; these are not recalculated or proof of native correctness.

Paste **Copy summary** first. Open **HTML report**, expand a numbered diagnostic part, and click **Copy part**. Paste parts in order as needed; each carries the suite ID and part count. Clipboard restrictions fall back to selecting the text for Ctrl+C. Errors and event bodies are not shortened in the diagnostic parts. Use screenshots from the UGREEN feed of relevant report previews or the visible Office app to review layouts and charts. No HTML, ZIP, PDF or video upload is required. Original outputs remain locally available for inspection; text and screenshots cannot prove every native file property.

Full permitted request and response bodies, tool calls/results, returned provider reasoning, source hashes, handoffs and timestamps are captured locally in encrypted event files. No API key or Authorization header is intentionally captured. The trace contains exactly the context sent through Scribble, which may already be bounded by normal extraction limits. Provider reasoning is available only if the endpoint returned it.

Enable the lab only in a dedicated synthetic session. Guards check case-specific file hashes, selected mail identity/body, saved Office sources and the loopback browser URL; newly created Office drafts carry the run ID. The operator must also start a clean chat and avoid other Scribble activity during capture. This is not an isolation boundary against a compromised local account.

Run `python evaluator-only/evaluate.py path/to/run.zip --kit path/to/scribble-test-kit-v1 --output path/to/report.json`. The evaluator checks archive hashes, trace completeness, formula/chart/slide structure and case facts, then produces machine findings and a native/visual review checklist. Complete native recalculation, source preservation, unsent state and visual review before marking an attempt passed. Artifact presence alone never means pass. Videos and evidence are private outputs and should not be committed to the repository.

The evaluator-only reference files are for human comparison. Never add them to Qwen context. The fixture server never exposes them. Run each priority case three times and preserve all failures. After a fix, compare the same prompts/build/model settings, then use fresh synthetic variants before claiming general improvement. This kit does not train model weights or certify the broader release gate.

## Rebuild in the repository

Use the bundled Python and Node runtime with `python-docx`, `reportlab`, and `@oai/artifact-tool`. Link the generator's `node_modules` to the bundled package directory. Run `generators/build_kit.py`, then `generators/build_office.mjs`, native/render QA, `generators/build_kit.py --pack`, and `evaluator/test_kit.py`. The committed ZIP is independently downloadable under `tests/benchmarks/releases/` on main.

Expected outputs include formula-based reference analysis, a six-slide editable deck, a memo, a draft email, operations PDF, and slide images. All amounts exclude tax. Margin changes are percentage points. Exact expected facts and the scoring rubric live under `evaluator-only/`.
