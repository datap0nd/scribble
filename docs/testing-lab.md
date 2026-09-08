# Testing Lab implementation

The regular Scribble assembly includes an operator-only Test Lab. It is disabled until `Enable-ScribbleTestLab.ps1` validates the fixture kit and writes a current-user DPAPI-protected descriptor. The descriptor expires after eight hours. The Office panes and Chrome native host consult the same descriptor. No model tool can enable the lab, start a run, collect a file or export evidence.

Download the kit from `tests/benchmarks/releases/scribble-test-kit-v1.zip` and follow its `START-HERE.md`. The kit contains 16 cases, five realistic synthetic email messages with attachments, clean/dirty/incomplete workbooks, two starter presentations, Word briefs, a PDF, a loopback website and evaluator-only reference outputs.

## Execution and isolation

Starting from an Office pane clears that pane's conversation and context. Add the case emails/attachments after Start. Chrome clears its old conversation before the first request in a new benchmark run. Later prompts and recovery within the same run retain the normal context. Other panes used manually should also start with a clean chat.

The original production request factories, model choice, tool contracts, limits, policies and draft writers remain in use. Run metadata is not sent to the model. Source attachment hashes are checked against the current case, including attachments from selected fixture emails. Mail reads check source identity/body; Office reads validate a fixture path/hash or a run tag on a newly generated document. The Chrome case stays on the loopback fixture site. Mismatches stop capture and mark the run incomplete.

New documents carry a `ScribbleTestRunId` custom property; new unsent drafts use an Outlook user property. This provides cross-process output identity and enables deliberate operator capture of those specific objects. The mode does not claim to isolate a compromised Windows account. Operators must keep other work out of the test session and never attach reference answers.

## Capture and export

`TaskDiagnostics.Record` forwards untruncated task/inference/tool events to a separate encrypted benchmark sink before normal diagnostic rotation. Pane status and final output events and suite handoff receipts are also recorded. Each event has a run ID, task ID, process-instance ID, local sequence, UTC time and monotonic ticks. Secret redaction runs before encryption, and transport headers are never supplied to the benchmark sink.

Trace storage has a 250 MB limit; individual artifacts have a 100 MB limit and the artifact folder has a 500 MB budget. I/O errors, quota exhaustion, source mismatch, disable-before-finish and unfinished task markers prevent complete evidence certification. Normal extraction truncation remains visible in the recorded payload. Reasoning is available only when the endpoint returned it in the response body.

The operator can capture tagged new Excel/PPT/Word documents and open Outlook drafts. Excel/PowerPoint save copies; Word's native Flat OPC is packaged as DOCX without rebinding its path; PDF/slide exports are derivatives. Outlook captures MSG, rendered HTML and native message/attachment metadata and refuses sent messages. Existing fixture documents with added draft sheets/slides require Save As to a separate output, followed by manual collection. This avoids granting model-facing save capabilities or automatically exporting unrelated source documents.

Export contains `run.json`, a merged `timeline.jsonl`, marker CSV, artifact receipts, `scorecard.json`, incomplete markers and an export hash manifest. Sorting timestamps assists video review; causal ordering uses task/tool IDs and per-instance sequences. The evaluator checks hashes, structure and factual presence and leaves semantic, source-preservation and native/visual review explicit. A score is never declared passed from chat wording or file presence alone.

## Regression and release workflow

The installer workflow validates the committed kit, extracts it and runs the Test Lab infrastructure checks, alongside the suite's existing guardrail and browser tests. `Test-TestLab.ps1` refuses to run over an existing operator descriptor. Native Office readiness and actual Qwen behavior are separate acceptance evidence, not simulated by these tests.

Use `evaluate.py` on each exported run, complete the generated review template against the actual native outputs and video, and rerun evaluation with `--review`. Use `compare_runs.py` to retain all attempts when comparing revisions. Keep evidence/videos out of git. Start with three attempts each of EX01, PP01 and XA01, fix an observed failure, then rerun those cases and fresh synthetic data. Benchmark success does not bypass `Test-ReleaseEvidence.ps1` or certify model-weight training.
