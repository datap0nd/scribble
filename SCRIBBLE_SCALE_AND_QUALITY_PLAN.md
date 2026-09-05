
# Scribble: scalable tasks and Samsung presentation quality

## Repository deliverables

Implementation specification. See `SCALE_IMPLEMENTATION_STATUS.md` for verified progress and remaining acceptance gates.

Save these two files in `C:\Users\keeoh\Documents\ChatGPT\Scribble` when editing is enabled:

- `SCRIBBLE_SCALE_AND_QUALITY_PLAN.md` — this implementation plan.
- `GEMINI_SAMSUNG_SLIDE_EXTRACTION.md` — the copyable Gemini instructions below.

Mark the older `AGENTIC_CONTEXT_PLAN.md` as superseded. Its proposed 500-message ceiling conflicts with the new requirements. Preserve the existing uncommitted Excel work and build on it.

Confirmed decisions: **only Markdown returns from the work PC**, and **long tasks must survive application restarts**.

## Implementation sequence

### 1. Outlook and the shared task coordinator

Build an asynchronous coordinator that separates the complete task from individual model requests. Store the objective, original user instructions, source references, completed batches, outstanding work, and evidence outside the model’s context.

Replace the single-search mailbox workflow with cursor-based pagination. Enumerate metadata first, then read and process bodies in batches. Run small COM operations on Outlook’s owning thread and yield between operations; never block a synchronous mailbox tool while waiting for model requests.

For “review all,” process every matching message. For targeted research, permit relevance filtering and record exclusions. Long bodies and attachments must also be processed incrementally.

Rewrite morning summary to cover its entire time window. Preserve exact coverage through batch analysis and merging; retain source references so later synthesis can verify claims.

**First milestone:** a complete morning summary and a 1,000-email review using the shared coordinator.

### 2. Automatic continuation across every application

Integrate the coordinator into Outlook, Excel, PowerPoint, Word, and Chrome.

Replace task-ending tool-count restrictions with sequential scheduling. Before each model request, budget instructions, tool definitions, evidence, images, and response space. Reduce batch size or compact completed exchanges when necessary. Keep source evidence separately retrievable.

Repeated unsuccessful actions trigger recovery and a different approach. Exhausted recovery produces a resumable task with a concrete blocker. It must never produce a generic “too many tools” failure.

Persist checkpoints using Windows protection for the current user and atomic file replacement. Resume after reopening the application once sources and destinations are revalidated. Browser controls must be rediscovered after restart.

Show ordinary progress and provide Stop/Resume. Stop pauses the task; explicitly discarding it removes its checkpoint.

### 3. Excel transformations and write authorization

Complete the existing selection-batching implementation. Start with approximately 100 rows per batch and adapt to text size. Preserve row identities, source values, shared terminology, and staged results.

The original instruction authorizes the complete transformation within its captured range. Batches and model-request boundaries do not consume that authorization.

Default output goes into an adjacent empty column. Explicit replacement instructions authorize replacing the selected source. Ask when no unambiguous destination exists. Recheck workbook identity and source/destination changes before writing.

Validate all rows for coverage and alignment. Perform semantic review in batches, repair flagged translations, and read back applied output. Journal writes so restart recovery distinguishes completed, missing, and uncertain writes.

### 4. Samsung PowerPoint generation

Use the returned Markdown specification to build a versioned theme and layout library in Scribble. Gemini is used to extract the design standard on the work PC; Scribble continues using its configured model endpoint.

For substantial deck creation, normally ask two or three relevant questions about audience, intended decision, depth, and source completeness. Preserve previously supplied answers.

Generate through:

**Brief → source analysis → storyline → layout selection → editable slides → rendered review → repair.**

Implement the layout families documented by the examples, including dense combinations of tables, charts, annotations, and conclusions. Enforce measured typography, geometry, spacing, and overflow behavior in the renderer.

Permit revision of slides created by the current task during review. Preserve source citations and verify numbers against the source material.

Because only Markdown returns, exact template preservation cannot be verified locally. Final visual acceptance happens on the work PC against the actual slides.

### 5. Chrome interaction and decision-making

Keep the existing extension/native-host architecture. Adapt observation, targeting, and recovery patterns from [Stagehand](https://github.com/browserbase/stagehand) and [Playwright MCP](https://github.com/microsoft/playwright-mcp), retaining applicable attribution when code is reused.

Make control discovery incrementally retrievable across page regions, frames, and shadow DOM. Support custom option cards, radio groups, menus, and native dropdowns without fixed total-option limits.

Resolve controls before acting; verify the intended result afterwards. Follow [Playwright’s actionability principles](https://playwright.dev/docs/actionability) for visibility, stability, enabled state, and receiving input.

For trade-in condition, use the user’s supplied condition. If absent, ask which condition applies and include “compare all” as an option. Explicit comparison requests should evaluate all conditions and return verified quotes with their assumptions.

## Interfaces, defaults, and acceptance

Introduce four shared contracts:

- **Task state:** objective, original user decisions, lifecycle, source bindings, batch ledger, evidence, and checkpoints.
- **Batch result:** stable batch ID, covered source IDs, structured output, evidence references, and failures.
- **Task authorization:** host-created permission tied to the original request, operation, source, and destination; separate from model summaries.
- **Context manager:** constructs each bounded model request and preserves complete tool-call/result pairing.

Task states are running, awaiting user, paused, completed, and discarded. Completion requires a reconciled coverage ledger; a model declaring completion is insufficient.

Saved documents resume automatically only when uniquely identified and validated. Missing documents require reopening or reselection. An unsaved document that disappeared must never be rebound to an unrelated workbook or presentation.

Keep the existing .NET Framework 4.8/C# 7.3 and Chrome architecture. Add no new cloud dependency. Preserve the existing restrictions on sending email and saving user documents.

| Area | Required acceptance |
|---|---|
| Outlook | Process 500 and 1,000 emails; include important late messages, long bodies, attachments, duplicates, and exact time boundaries. |
| Continuation | Exceed existing round limits, shrink oversized requests, recover from context rejection, and retain earlier constraints. |
| Restart | Interrupt before and after a write; resume without duplicate output or binding to the wrong document. |
| Excel | Translate and review 20,000 rows; handle blanks, long cells, formulas, changed sources, and occupied destinations. |
| PowerPoint | Generate representative layouts from the specification; check editability, content accuracy, overflow, and visual match on the work PC. |
| Chrome | Complete condition selection and quote comparison; test custom controls, overlays, delayed updates, stale references, and frames. |

Update tests and static checks that currently enforce obsolete scope limits. Add the browser fixtures to CI. Validate each milestone before continuing to the next.

## Exactly what to do on the work PC

1. **Choose the decks representing the current standard.** Give each an alias such as `D01`, `D02`. Keep the alias-to-filename mapping on the work PC. Note any slides you consider especially good.

2. **Export PDFs containing approximately 20 slides each.** In PowerPoint, use **File → Export → Create PDF/XPS → Options**, choose the slide range, and publish slides at standard quality. Name files such as `D01_001-020.pdf`. Preserve original slide numbering, including a mapping if hidden slides are excluded. [Microsoft’s export instructions](https://support.microsoft.com/en-au/office/save-powerpoint-presentations-as-pdf-files-9b5c786b-9c6e-4fe6-81f6-9372f77c47c8)

3. **Copy the Markdown block below into a file named `GEMINI_SAMSUNG_SLIDE_EXTRACTION.md`.** Attach it to Gemini on the work PC, or paste its contents.

4. **Run one batch in a fresh Gemini conversation.** Attach the batch PDF and, if accepted, the corresponding original PPTX for metadata inspection. Send:

   > MODE: BATCH. Deck alias: D01. Original slides: 1–20. PDF page 1 corresponds to original slide 1. Preferred examples: none specified. Follow the attached extraction instructions. Produce SAMSUNG_BATCH_D01_001-020.md.

   Adjust the alias, range, mapping, and preferred examples for each batch. Check the output accounts for every slide before continuing. Gemini supports file analysis, but large inputs can lose detail, which is why this workflow uses separate batches. [Google’s file-analysis guidance](https://support.google.com/gemini/answer/14903178?hl=en-GB)

5. **Consolidate the batch Markdown files in a new conversation.** Attach this instruction file and the batch reports, then send:

   > MODE: MERGE. Produce one complete SAMSUNG_SLIDE_DESIGN_SPEC.md from all attached batch reports. Preserve the coverage ledger, measured details, layout variants, and unresolved questions.

   For more files than Gemini accepts in one upload, merge groups first, then merge those outputs while preserving their source ledgers.

6. **Audit the specification on the work PC.** Attach the specification and one source PDF batch at a time. Send:

   > MODE: AUDIT. Compare the specification against every slide in this PDF. Produce exact corrections and identify details that need checking in PowerPoint.

   Resolve missing font names, slide dimensions, and other important unknowns using the original PPTX. After collecting corrections, run MERGE again with the specification and all corrections.

7. **Bring back `SAMSUNG_SLIDE_DESIGN_SPEC.md`.** Keep the actual slides and batch PDFs on the work PC. A first batch report is also useful before all 200 slides are finished.

