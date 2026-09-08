# Scribble synthetic testing kit

The kit contains fictional Atlas Office Supplies data, 16 exact prompts and evaluator-only answers. Start with **EX01**, **PP01**, and **XA01**. Their expected June revenue is EUR 120000, budget gap EUR 10000 and margin 38.33%.

## Start a recorded test

1. Install a Scribble build that includes Test Lab. Extract `scribble-test-kit-v1.zip` and verify its adjacent SHA-256. Use classic Office, a clean conversation and only the listed synthetic inputs.
2. In Windows PowerShell 5.1, run `operator/Enable-ScribbleTestLab.ps1`. If the DLL is not found, pass `-AssemblyPath` with the installed `Scribble.dll` path. Activation expires after eight hours and is off by default.
3. Run `operator/Import-ScribbleTestMail.ps1` for Outlook cases. It creates a separate local PST and never sends mail. Select the specified messages in that folder and add them as Scribble's locked working set. EML originals and their attachments are in `inputs/outlook/`.
4. For Chrome, run `operator/Serve-ScribbleFixtures.ps1` (Python 3 required) and open `http://127.0.0.1:8765/index.html`. Only the three synthetic browser pages are served. Do not serve the whole kit.
5. Open the **Test Lab** button in any Scribble pane. Select a case, read its setup, prepare its inputs, then click **Start case**. Confirm that the context is synthetic. Start resets the conversation, so add the listed attachment and mail context after starting. The exact prompt is copied to the clipboard. Start the screen recording and add a video marker.
6. Paste the prompt in Scribble and run it normally. Record questions, delays and errors. Do not silently fix outputs. Mark any unscripted coaching as assisted. For a follow-up case, execute its prerequisite as part of the same recorded attempt.
7. Click **Capture new Office drafts** to copy newly created, run-tagged workbooks, decks, documents and open unsent email drafts, with PDFs and slide PNGs. Only drafts created in this run qualify. For draft sheets/slides inside the starter file, use Office **Save As** to a NEW run-folder file, then **Collect saved outputs**. That copies your selected files into the evidence bundle. Never overwrite an input fixture. Manual collection remains available when native copy/export fails.
8. Finish the clean or assisted attempt, then **Export run ZIP**. Share that ZIP and the video together. To export later, use `Export-ScribbleTestRun.ps1 -RunId ... -DestinationDirectory ...`.
9. Disable with `operator/Disable-ScribbleTestLab.ps1`. Disabling during a run leaves an incomplete evidence marker.

PowerPoint PP01 uses the starter deck plus three attachments: sales, budget and the operations PDF. Word WD01 uses the open brief plus the same three attachments. For follow-up cases, Start copies the prerequisite prompt first; run it, then use Copy prompt for the actual follow-up. A native output from that prerequisite is the next input, never the reference answer.

## Evidence and evaluation

Full permitted request and response bodies, tool calls/results, returned provider reasoning, source hashes, handoffs and timestamps are captured locally in encrypted event files. No API key or Authorization header is intentionally captured. The trace contains exactly the context sent through Scribble, which may already be bounded by normal extraction limits. Provider reasoning is available only if the endpoint returned it.

Enable the lab only in a dedicated synthetic session. Guards check case-specific file hashes, selected mail identity/body, saved Office sources and the loopback browser URL; newly created Office drafts carry the run ID. The operator must also start a clean chat and avoid other Scribble activity during capture. This is not an isolation boundary against a compromised local account.

Run `python evaluator-only/evaluate.py path/to/run.zip --kit path/to/scribble-test-kit-v1 --output path/to/report.json`. The evaluator checks archive hashes, trace completeness, formula/chart/slide structure and case facts, then produces machine findings and a native/visual review checklist. Complete native recalculation, source preservation, unsent state and visual review before marking an attempt passed. Artifact presence alone never means pass. Videos and evidence are private outputs and should not be committed to the repository.

The evaluator-only reference files are for human comparison. Never add them to Qwen context. The fixture server never exposes them. Run each priority case three times and preserve all failures. After a fix, compare the same prompts/build/model settings, then use fresh synthetic variants before claiming general improvement. This kit does not train model weights or certify the broader release gate.

## Rebuild in the repository

Use the bundled Python and Node runtime with `python-docx`, `reportlab`, and `@oai/artifact-tool`. Link the generator's `node_modules` to the bundled package directory. Run `generators/build_kit.py`, then `generators/build_office.mjs`, native/render QA, `generators/build_kit.py --pack`, and `evaluator/test_kit.py`. The committed ZIP is independently downloadable under `tests/benchmarks/releases/` on main.

Expected outputs include formula-based reference analysis, a six-slide editable deck, a memo, a draft email, operations PDF, and slide images. All amounts exclude tax. Margin changes are percentage points. Exact expected facts and the scoring rubric live under `evaluator-only/`.
