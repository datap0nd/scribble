# Scribble synthetic testing kit

The kit contains fictional Atlas Office Supplies data, 16 exact prompts and evaluator-only answers. Start with **EX01**, **PP01**, and **XA01**. Their expected June revenue is EUR 120000, budget gap EUR 10000 and margin 38.33%.

## Start a recorded test

1. Install a Scribble build that includes Test Lab. Extract `scribble-test-kit-v1.zip` and verify its adjacent SHA-256. Use classic Office, a clean conversation and only the listed synthetic inputs.
2. In Windows PowerShell 5.1, run `operator/Enable-ScribbleTestLab.ps1`. If the DLL is not found, pass `-AssemblyPath` with the installed `Scribble.dll` path. Activation expires after eight hours and is off by default.
3. Open **Test Lab** in any Scribble pane, choose a case and click **Prepare case**. It opens the required Office apps and fixture documents, PDF files, and exact Outlook case messages in a separate synthetic PST. It also opens apps needed for the expected outputs. Source Office files open read-only; existing unsaved fixture edits require your attention. Other open documents are not closed or saved.
4. For Chrome cases, preparation starts a fixture-only loopback server automatically and opens the required pages in Chrome. Python is not needed. The server chooses an available local port and stops when Test Lab is disabled or expires. Open the Scribble Chrome side panel yourself and check its native connection. Preparation reports missing apps/add-ins and remaining steps.
5. In the case's starting app, open **Test Lab**, choose the same case and click **Start case**. Confirm the synthetic context. Start clears the Scribble conversation, adds the case's file or selected email context, and fills the prompt without submitting it. Close the Test Lab window to return to Chrome; its prompt fills on return. Check the context tray and wait for attachment reading to finish. The prompt is also on the clipboard.
6. Start screen recording, add a video marker, then submit the prepared prompt yourself. Record questions, delays and errors. Do not silently fix outputs. Mark any unscripted coaching as assisted. For a follow-up case, execute its prerequisite as part of the same recorded attempt.
7. Click **Capture new Office drafts** to copy newly created, run-tagged workbooks, decks, documents and open unsent email drafts, with PDFs and slide PNGs. Only drafts created in this run qualify. For draft sheets/slides inside the starter file, use Office **Save As** to a NEW run-folder file, then **Collect saved outputs**. That copies your selected files into the evidence bundle. Never overwrite an input fixture. Manual collection remains available when native copy/export fails.
8. Finish the clean or assisted attempt, then **Export PDF + evidence**. This saves a PDF report, a pasteable `-summary.txt`, an HTML copy, and the evidence ZIP beside each other. The ZIP also contains the PDF, summary and original collected files. Use **Open report PDF** to screenshot its first page, or **Copy report summary** to paste directly into chat. To export later, use `Export-ScribbleTestRun.ps1 -RunId ... -DestinationDirectory ...`.
9. Disable with `operator/Disable-ScribbleTestLab.ps1`. Disabling during a run leaves an incomplete evidence marker.

PowerPoint PP01 uses the starter deck plus three attachments: sales, budget and the operations PDF. Word WD01 uses the open brief plus the same three attachments. For follow-up cases, Start copies the prerequisite prompt first; run it, then use Copy prompt for the actual follow-up. A native output from that prerequisite is the next input, never the reference answer.

Preparation can also run before opening a Scribble pane. From the extracted `operator` folder:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\Prepare-ScribbleTestCase.ps1 -CaseId PP01
```

Replace `PP01` with `EX01`, `XA01`, or another case ID. Enable Test Lab first. `-PlanOnly` lists the required apps/files without opening them. **Stop preparation** stops the setup helper and leaves opened apps intact. If Office is waiting for sign-in or a first-run dialog, resolve it before retrying. Preparing is blocked during an active test run. Manual setup and the existing import/server scripts remain available if native automation is unavailable. A connected add-in check does not certify actual model or native artifact behavior.

If activation reports `Unable to find type [Scribble.Testing.TestLab]`, install Scribble 2.0.132.0 or newer from the [official installer](https://github.com/datap0nd/scribble/releases/latest/download/ScribbleSetup.exe), then open a new Windows PowerShell process. From the extracted `operator` folder, run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Enable-ScribbleTestLab.ps1 -AssemblyPath "$env:LOCALAPPDATA\Programs\Scribble\Scribble.dll"
```

The script is named `Enable-ScribbleTestLab.ps1`. The loader prefers the current install folder over the legacy `%LOCALAPPDATA%\Scribble` folder and prints the DLL it loaded. For a custom installation, supply that installation's DLL path instead.

## Evidence and evaluation

The PDF starts with case/run/build/model identity, UTC timestamps, capture completeness, missing output types and initial error highlights. Later pages contain the prompt, expected result, video markers, full error-like events, model responses, output receipts, available slide PNGs, native text previews and the complete recorded timeline without shortened event bodies. Spreadsheet previews show the first 200 rows per sheet and formulas without recalculation. Original Excel/PowerPoint/Word/PDF/Outlook files remain in the ZIP; their presence does not certify correctness. Edge or Chrome renders the PDF in a separate headless process with an isolated profile. If rendering fails, the ZIP, HTML report, summary and renderer diagnostics are retained with an explicit error.

For quick feedback, paste the summary and add screenshots of page one and the relevant result/error page. For a full diagnosis, attach the PDF or evidence ZIP; a video remains useful for UI delays, clicks and visual problems. Very long runs can produce a long PDF because the full captured timeline is retained.

Full permitted request and response bodies, tool calls/results, returned provider reasoning, source hashes, handoffs and timestamps are captured locally in encrypted event files. No API key or Authorization header is intentionally captured. The trace contains exactly the context sent through Scribble, which may already be bounded by normal extraction limits. Provider reasoning is available only if the endpoint returned it.

Enable the lab only in a dedicated synthetic session. Guards check case-specific file hashes, selected mail identity/body, saved Office sources and the loopback browser URL; newly created Office drafts carry the run ID. The operator must also start a clean chat and avoid other Scribble activity during capture. This is not an isolation boundary against a compromised local account.

Run `python evaluator-only/evaluate.py path/to/run.zip --kit path/to/scribble-test-kit-v1 --output path/to/report.json`. The evaluator checks archive hashes, trace completeness, formula/chart/slide structure and case facts, then produces machine findings and a native/visual review checklist. Complete native recalculation, source preservation, unsent state and visual review before marking an attempt passed. Artifact presence alone never means pass. Videos and evidence are private outputs and should not be committed to the repository.

The evaluator-only reference files are for human comparison. Never add them to Qwen context. The fixture server never exposes them. Run each priority case three times and preserve all failures. After a fix, compare the same prompts/build/model settings, then use fresh synthetic variants before claiming general improvement. This kit does not train model weights or certify the broader release gate.

## Rebuild in the repository

Use the bundled Python and Node runtime with `python-docx`, `reportlab`, and `@oai/artifact-tool`. Link the generator's `node_modules` to the bundled package directory. Run `generators/build_kit.py`, then `generators/build_office.mjs`, native/render QA, `generators/build_kit.py --pack`, and `evaluator/test_kit.py`. The committed ZIP is independently downloadable under `tests/benchmarks/releases/` on main.

Expected outputs include formula-based reference analysis, a six-slide editable deck, a memo, a draft email, operations PDF, and slide images. All amounts exclude tax. Margin changes are percentage points. Exact expected facts and the scoring rubric live under `evaluator-only/`.
