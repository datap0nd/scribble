# Independent Review of the Scribble Automatic Context Plan

Review target: `AGENTIC_CONTEXT_PLAN.md`

Review status: implementation should not begin from the original plan without
revision.

The review was performed independently against the current Scribble codebase.
It focused on Outlook COM behavior, asynchronous orchestration, endpoint context
budgeting, security provenance, bounded work, and the smallest safe first
release.

## Overall verdict

The product direction is sound: users should speak naturally, context
maintenance should be automatic, and no `analyze_mailbox` command should be
introduced.

The implementation plan is too broad for a first release and contains two
unsafe integration assumptions:

1. A 500-message pipeline cannot reuse the current item-by-item Outlook search
   path without freezing the Outlook UI.
2. Asynchronous mapper/reducer model calls cannot be hidden inside the current
   synchronous mailbox tool host without risking a UI-context deadlock.

The reviewer recommends separating automatic conversation compaction from broad
mailbox processing and shipping the smaller context-continuity change first.

## Priority 0 blockers

### 1. Outlook COM enumeration and cancellation

`MailboxToolHost.Execute` is synchronous and is called from the Outlook UI
thread in `src/Scribble/UI/ChatPane.cs`. The current search path instantiates
Outlook items and calls `MessageReader.CaptureItem`, which reads body, HTML, and
attachment information.

Reusing that path for 500 messages would perform substantial COM work on the UI
thread. The Stop handler cannot respond reliably while that work is running,
and Outlook Object Model calls should not simply be moved to `Task.Run` because
of COM apartment/thread affinity.

Required redesign before broad scanning:

- Use an Outlook `Table`/`GetArray` metadata path instead of opening every item.
- Apply one shared candidate ceiling across all searched folders.
- Process small pages on the owning UI thread and yield between pages.
- Check cancellation and a request deadline between every page and body load.
- Release individual COM objects deterministically.
- Benchmark the path in real Outlook before increasing scope.

Relevant current paths:

- `src/Scribble/UI/ChatPane.cs`, mailbox tool dispatch around lines 2215–2261.
- `src/Scribble/Outlook/MailboxToolHost.cs`, synchronous `Execute` around lines
  69–118.
- `src/Scribble/Outlook/MailboxContextService.cs`, item capture around lines
  235–277.
- `src/Scribble/Outlook/MessageReader.cs`, body/HTML/attachment capture around
  lines 257–276.

### 2. Async orchestration cannot live inside synchronous `search_mailbox`

Mapper and reducer requests are asynchronous. Blocking on them inside
`MailboxToolHost.Execute` risks deadlock because the OpenAI-compatible client's
await continuations may need the UI synchronization context.

The existing `search_mailbox` result contract also promises message metadata
and temporary handles. Returning a multi-minute corpus digest under the same
shape would silently change that semantic contract even if the tool-call ID
were preserved.

Required redesign:

- Put broad processing in an asynchronous orchestration step owned by
  `CompleteMailboxChatAsync`, outside `MailboxToolHost.Execute`; or
- Introduce an explicit, typed result union and test every caller.

Do not synchronously wait for model requests inside the current tool host.

### 3. Context-budget guarantees are overstated

The endpoint model list currently provides an identifier, not the deployed
server's configured context limit. Qwen's advertised native maximum therefore
cannot guarantee the actual endpoint capacity. Token usage arrives only after a
request and cannot make the first preflight exact.

Normal mailbox requests also leave `max_tokens` unset, while Qwen reasoning may
consume significant output capacity. Images consume model context even when
their base64 byte limit is tracked separately.

Required changes:

- Promise a conservative estimate and bounded recovery, not a mathematical
  guarantee for arbitrary OpenAI-compatible endpoints.
- Set explicit output limits for normal and utility requests.
- Disable or minimize thinking for utility compaction when supported.
- Key any learned estimate by normalized endpoint, actual model, and modality.
- Reset learned state when endpoint/model configuration changes.
- Exclude vision from the first compaction MVP or account for it explicitly.
- Classify context-length errors and allow only one reduced retry.

Relevant current paths:

- `src/Scribble/Chat/ChatModels.cs`, model-list response around lines 101–109.
- `src/Scribble/Chat/OpenAiCompatibleClient.cs`, model parsing around lines
  471–483 and infinite HTTP timeout near lines 45–48.
- `src/Scribble/Chat/ChatRequestFactory.cs`, unset normal `max_tokens` around
  lines 168–171.

### 4. Summary provenance and authorization need typed separation

`ChatTurn` stores only role and content. It cannot distinguish original user
text, model-produced compacted state, host-authoritative state, and untrusted
email/tool material.

The compacted snapshot must never become a system message or be passed through
an authorization policy as though it were an original user request. Active
draft identity also does not belong in the model-produced snapshot because the
host already supplies authoritative draft state.

Required changes:

- Introduce typed provenance for original user turns, assistant turns,
  compacted reference state, host state, and tool/email data.
- Insert compacted state only as an explicitly untrusted reference block.
- Pass the latest original prompt and a host-created authorization object
  separately to the broad coordinator and tool host.
- Keep active draft identity entirely host-authored.
- Treat evidence handles as references to messages, not proof that a generated
  claim is true.

Relevant current paths:

- `src/Scribble/Chat/ChatModels.cs`, `ChatTurn` around lines 5–16.
- `src/Scribble/UI/ChatPane.cs`, host-authored request/draft state around lines
  2073–2127.

### 5. Total work, time, cost, and coverage are not bounded

A 500-candidate ceiling alone does not bound body reads, mapper calls,
reductions, retries, wall time, or cost. The current HTTP client uses an
infinite timeout. Exact coverage is also impossible if reducer failures allow
newly processed batches to be counted without an idempotent successful merge.

Required limits:

- Maximum total cleaned body characters/tokens per request.
- Maximum bodies loaded.
- Maximum mapper and reducer calls.
- Maximum retry count.
- Per-call and request-wide deadlines.
- Maximum tolerated failed/skipped batches.
- Idempotent batch identifiers and transactional digest merges.

Coverage must distinguish:

- metadata candidates enumerated;
- bodies successfully read;
- messages successfully mapped;
- messages skipped by policy or budget;
- messages failed because of COM or endpoint errors;
- messages successfully merged into the final digest.

The product contract must also resolve whether “review 500 emails” means all 500
bodies were processed or merely that 500 candidates were considered.

## Recommended smallest safe MVP

Ship conversation continuity before broad mailbox agency.

Scope:

- Outlook text only.
- No vision or attachments.
- No browser or Office integration initially.
- No broad mailbox scan.
- No cross-request evidence ledger.
- No second verifier model call.

Implementation:

1. Serialize model requests so context state cannot be mutated concurrently.
2. Capture endpoint token usage where available.
3. Add conservative request-size estimation and explicit output reserves.
4. Set finite HTTP/request deadlines.
5. Classify context-length errors and allow one smaller retry.
6. Compact only old, complete user/assistant pairs.
7. Produce one bounded JSON snapshot through a no-tools utility call.
8. Validate its schema, size, and provenance locally.
9. Insert it as untrusted reference data while preserving recent raw turns.
10. Keep the visible transcript unchanged and make cancellation transactional.

MVP exit criteria:

- Seeded goals and constraints survive beyond the current history window.
- Draft authorization still depends exclusively on the latest original user
  prompt.
- No tool-call/tool-result pair is split.
- No utility request exposes tools.
- A failed, cancelled, oversized, or malformed summary never replaces the last
  valid state.
- One classified context error produces at most one reduced retry.

## Recommended second-stage Outlook pilot

After the MVP is stable, prototype broad processing with deliberately smaller
limits:

- At most 100 metadata rows through an Outlook Table query.
- At most 25 bodies.
- Fixed byte-sized mapper batches.
- Text body only; no attachments, HTML images, or vision.
- Finite per-call and overall deadlines, initially 60–120 seconds depending on
  measured endpoint behavior.
- Cancellation between every COM page, body load, and mapper batch.
- Exact transactional coverage accounting.

Increase the limits toward 500 only after real Outlook and target-endpoint
benchmarks demonstrate acceptable latency, cancellation, memory use, and cost.

## Complexity to remove from the first release

- Second LLM verifier pass.
- Learned adaptive context ceilings.
- Token-driven variable mailbox batch sizing.
- Active-draft state inside the compacted snapshot.
- Browser and Office context integration.
- Cross-request evidence ledger.
- Hiding a corpus digest behind the current `search_mailbox` result shape.

These can be revisited after the smaller MVP establishes safe provenance,
budgeting, cancellation, and failure semantics.

## Review conclusion

Keep the product vision but split delivery:

1. Automatic conversation compaction and conservative request budgeting.
2. A separately designed asynchronous Outlook metadata/body pilot.
3. Expansion toward 500 only from measured evidence.

The first implementation should not begin until the original plan is revised to
reflect this split and the Priority 0 blockers above.

## Final reviewer corrections

- Much of the proposed Phase 0 already exists in
  `tests/GuardrailTests/Program.cs`: the three mailbox tools, ten-message working
  set, bounded history, latest-prompt draft authorization, and Gemini
  disablement already have guardrails. New baseline work should concentrate on
  request budgeting, compaction provenance, broad-intent authorization, COM
  cancellation, and transactional coverage state.
- `PaneMemory` intentionally survives pane recreation and is keyed by host kind,
  so pane disposal is not currently the correct clearing boundary. Browser
  history also lives in extension JavaScript rather than the desktop pane
  memory. This reinforces the recommendation to keep the first release
  Outlook-only and define state ownership explicitly before expanding to all
  chat paths.

External technical references used by the reviewer:

- Microsoft Outlook API/threading guidance:
  https://learn.microsoft.com/en-us/office/client-developer/outlook/selecting-an-api-or-technology-for-developing-solutions-for-outlook
- Outlook Table API:
  https://learn.microsoft.com/en-us/office/vba/api/outlook.table
- Qwen3.8-27B model information:
  https://huggingface.co/Qwen/Qwen3.8-27B-FP8
