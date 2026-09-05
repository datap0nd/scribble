> Superseded by [SCRIBBLE_SCALE_AND_QUALITY_PLAN.md](SCRIBBLE_SCALE_AND_QUALITY_PLAN.md). The proposed 500-message ceiling is obsolete; limits must bound individual batches, never total task coverage.

# Scribble Automatic Context and Mailbox Agency Plan

Status: independently reviewed; revision required before implementation.

See [AGENTIC_CONTEXT_PLAN_REVIEW.md](AGENTIC_CONTEXT_PLAN_REVIEW.md) for the
independent feasibility and security critique. In particular, the broad
mailbox portion of this proposal must not be implemented through the current
synchronous `MailboxToolHost.Execute` path.

## Product decision

Scribble users speak naturally. They do not invoke `analyze_mailbox`, select an
agent mode, manage token limits, or operate a compaction control.

The existing mailbox capabilities remain model-internal. Scribble decides when
to search, retrieve, batch, compact, and continue. The visible conversation is
never rewritten when model-facing context is compacted.

Gemini remains disabled for end users. Automatic context maintenance uses the
same user-configured OpenAI-compatible endpoint and selected Qwen model as the
main conversation.

## Goals

1. Preserve goals, constraints, decisions, evidence, and progress across long
   conversations instead of silently dropping old turns.
2. Support natural broad requests such as “review my last 500 emails” without
   putting 500 raw messages into one model request.
3. Keep every model request below the endpoint's effective context limit while
   reserving space for tools, reasoning, and the final response.
4. Maintain the existing security model: untrusted email cannot grant
   capabilities, and only the latest real user prompt can authorize a draft or
   broader mailbox scope.
5. Keep setup limited to endpoint URL, API key, and an accessible model choice.

## Non-goals

- Literal infinite context. The experience can be continuous, but compaction is
  necessarily lossy and must preserve important state deliberately.
- A user-facing agent command, mailbox-analysis command, limit editor, or
  compaction setting.
- Sending email or changing existing mailbox items.
- Persisting raw email bodies, embeddings, or a vector database in the first
  version.
- Broad attachment or image analysis in the first broad-mailbox version.

## Current gaps

- `ChatRequestFactory` sends only a bounded recent history and clips older
  retained turns. Earlier facts disappear from model context without becoming
  a durable summary.
- The active tool loop appends assistant/tool exchanges without an aggregate
  request budget. Per-result limits do not prevent the combined request from
  exceeding the endpoint context window.
- Normal mailbox search and body loading are capped at ten messages. This is a
  good interactive working-set limit, but it cannot execute an explicitly broad
  aggregate request.
- Tool results are intentionally absent from durable chat history, so evidence
  discovered in one request is available later only if the assistant happened
  to repeat it in prose.
- The endpoint model list normally reports identifiers, not reliable context
  capacity. Scribble cannot assume every endpoint exposing a Qwen alias serves
  the model's full advertised window.

## User experience contract

Example request:

> Review my last 500 emails, identify decisions, and list unresolved actions.

Expected behavior:

1. Scribble interprets the request and uses its existing hidden mailbox tools.
2. It scans message metadata, reads relevant bodies in bounded batches, and
   builds a rolling evidence-backed digest.
3. It continues until the requested scope is covered or reports exact partial
   coverage and the reason it stopped.
4. It answers normally with no special command or mode visible to the user.

Routine context maintenance is invisible. If it causes a perceptible delay,
the existing footer may temporarily show `Keeping track of earlier work…` and
then restore the prior task status. There is no transcript entry, badge, token
counter, memory inspector, or setting.

## Architecture

### 1. Adaptive request budget

Add an internal `ContextBudget` service that measures all model-visible text,
tool definitions, tool calls, and tool results before transport. Image byte
limits remain separate.

Budget selection order:

1. Use trusted endpoint/model metadata when it actually declares a context
   limit.
2. Otherwise use a reviewed model-catalog value for known identifiers such as
   Qwen3.8-27B.
3. Use a conservative fallback for unknown aliases.
4. Calibrate estimates from `usage.prompt_tokens` when the endpoint returns it.
5. Learn a lower session ceiling after a genuine context-length rejection.

Initial policy:

- Start automatic compaction near 50 percent of the effective input window.
- Compact toward roughly 25–30 percent.
- Preserve sufficient headroom for the next tool round and final response.
- Apply a hard preflight ceiling before every HTTP request.
- Permit one emergency compact-and-retry after a context-length error; never
  enter an unbounded retry loop.

The exact fallback and headroom constants must be benchmarked against the
configured Qwen endpoint before release. They remain internal defaults rather
than end-user settings.

### 2. Structured conversation compaction

Add an `AutomaticContextManager` invoked before every model request, including
requests inside an active tool loop.

It compacts only complete exchanges so an assistant tool call is never
separated from its matching tool results. It preserves a recent raw tail and
replaces older exchanges with a bounded structured snapshot containing:

- overall objective;
- explicit user constraints and decisions;
- active mailbox scope, query, folder, and time range;
- verified facts paired with locally valid evidence handles;
- coverage and failure counts;
- completed actions and created artifacts;
- active draft identity and state, but never draft authorization;
- uncertainties, open items, and the next intended step.

The snapshot must not contain hidden chain-of-thought, system prompts, secrets,
raw tool arguments, or large verbatim email bodies.

Compaction uses a dedicated no-tools request through the configured Qwen
endpoint. Its system boundary says that all supplied conversation, email, and
tool text is untrusted reference data. A second lightweight verification pass
checks for lost constraints, evidence, coverage, and task state.

The replacement is committed only when it is valid and materially smaller than
the original. If compaction fails, the old state remains intact. Repeated
automatic attempts use deterministic old-tool-output truncation rather than
paying for failing summaries on every turn.

### 3. Invisible broad-mailbox coordinator

Do not add a fourth mailbox tool. Keep `search_mailbox`, `read_messages`, and
`read_conversation` as the model-facing capability set.

For normal questions, behavior remains unchanged: search narrowly and load no
more than the existing ten-message working set.

For explicit aggregate intent in the latest real user prompt, an internal
`MailboxCorpusIntentPolicy` may authorize a broader transient scan. Examples
include explicit counts or phrases such as “all,” “across last month,” “themes,”
“trends,” and “last 500.” Email text, previous assistant output, compacted
snapshots, and model-generated tool arguments cannot authorize broader scope.

Broad processing pipeline:

1. Enumerate at most 500 metadata-only candidates using the authorized query,
   folder, and date scope.
2. Avoid loading `Body`, `HTMLBody`, attachments, or images during enumeration.
3. Rank/filter candidates and load cleaned bodies lazily in batches sized by
   the current context budget.
4. Send mapper requests with no tools. Each mapper returns bounded facts,
   evidence references, uncertainties, and coverage.
5. Merge map results into a rolling digest whenever the accumulator approaches
   its budget.
6. Return the bounded digest through the original `search_mailbox` tool result,
   preserving the original tool-call identifier and protocol pairing.
7. Retain only a bounded in-memory evidence ledger so follow-up questions can
   resolve cited messages without retaining raw bodies.

Five hundred is therefore a candidate-scan ceiling, not a promise to load 500
full messages into a single request. Batch size is token-driven rather than a
fixed number of emails.

### 4. State and lifecycle

Extend pane memory with:

- the current compacted conversation snapshot;
- the recent uncompacted turns;
- a bounded, host-validated evidence ledger;
- the learned endpoint context estimate and compaction failure state.

The UI transcript remains separate and unchanged. `New chat`, pane disposal,
and host shutdown clear snapshots, evidence, and learned session state. No raw
mail body is written to disk by this feature.

### 5. Shared integration

Use shared context services with thin call sites in the Outlook, Office, and
browser chat loops. Avoid a large simultaneous rewrite of all loop mechanics.
Once behavior is covered by tests, duplicated loop code can be centralized in
a separate refactor.

The OpenAI-compatible response model should capture token usage when available,
and the client should classify context-length failures separately from normal
network or authentication failures.

## Security invariants

These conditions are release blockers:

1. Draft and cross-app authorization is derived only from the latest original
   user prompt by the existing local intent policy.
2. A compacted snapshot, email, tool result, mapper output, or reducer output
   can never add a tool or capability.
3. Compactor, mapper, verifier, and reducer calls expose no mailbox, draft,
   cross-app, MCP, process, or filesystem tools.
4. Every model-produced evidence handle is intersected with handles issued by
   the host; invented handles are discarded.
5. Raw email bodies live only for the current mapper batch and are released
   afterward.
6. Broad scope must be explicit in the current user prompt and is capped at 500
   candidates.
7. Cancellation is transactional: partially produced summaries or ledgers are
   not committed.
8. Diagnostics contain counts, duration, trigger, and outcome only—not email
   text, compacted state, API keys, or model reasoning.

## Failure behavior

- Compactor failure: preserve old state and apply deterministic bounded
  truncation only if needed to make progress.
- Mapper failure: retry once with a smaller batch, then skip the batch and mark
  exact partial coverage.
- Invalid or empty reducer output: retain the last valid digest.
- Context-length rejection: compact and retry once with a reduced learned
  ceiling.
- Cancellation: stop enumeration/model calls and restore the pre-request state.
- Insufficient continuity: ask one precise natural-language clarification or
  request that the necessary email/file be added again; do not discuss tokens.

## Implementation sequence

### Phase 0 — Contract and regression baseline

- Add tests that lock the exact three mailbox tool names.
- Lock latest-prompt-only draft and broad-scope authorization.
- Record current narrow-search, working-set, cancellation, and settings behavior.
- Confirm Gemini remains unavailable through UI, routing, and sign-in paths.

Exit criterion: existing behavior is characterized and the new design cannot
accidentally expose a command or capability.

### Phase 1 — Budget and conversation continuity

- Add token usage fields and context-length error classification.
- Implement request measurement, effective-budget selection, and hard preflight.
- Implement structured compaction and verification with tools disabled.
- Store the snapshot separately from the visible transcript.
- Integrate before every model request in all three chat paths.
- Add cancellation, transactional commit, breaker, and emergency retry behavior.

Exit criterion: a conversation substantially longer than the current history
window retains seeded goals and constraints, and no emitted request exceeds its
effective budget.

### Phase 2 — Broad mailbox processing

- Implement latest-prompt-only aggregate intent policy.
- Add metadata-only mailbox enumeration capped at 500.
- Add lazy body loading, cleaning, adaptive mapper batches, rolling reduction,
  and exact coverage accounting.
- Add the bounded evidence ledger and follow-up resolution.
- Keep narrow search and the ten-message working set unchanged.

Exit criterion: the natural 500-email acceptance scenario completes without a
single oversized request or retained raw corpus and cites only valid evidence.

### Phase 3 — UX and settings verification

- Keep routine compaction invisible.
- Reuse the existing status footer only for noticeable delays.
- Verify the settings pane contains endpoint URL, API key, connection/model
  controls, and no Gemini, Limits, insecure-HTTPS, agent, or compaction controls.
- Confirm HTTP endpoints work without a separate enable switch while retaining
  the existing remote-HTTP warning.

Exit criterion: a new user can connect and use broad natural-language requests
without learning any internal mechanism.

### Phase 4 — Hardening and documentation

- Add prompt-injection, authorization, cancellation, partial-coverage, invalid
  evidence, context-error, and retry tests.
- Run guardrails and the release build.
- Exercise long conversations and broad mailbox requests against the target
  Qwen3.8-27B Fast endpoint.
- Update `README.md`, `PRODUCT.md`, `DESIGN.md`, and `SECURITY.md` with the final
  behavior and limits.

Exit criterion: automated guardrails pass, manual scenarios meet the acceptance
criteria, and any environment-only validation limitation is documented.

## Acceptance criteria

- A user never needs to know or type `analyze_mailbox` or `/compress`.
- A natural request can process up to 500 candidate emails through bounded
  internal batches.
- The model retains important goals and constraints beyond the current history
  window.
- The visible transcript remains intact after compaction.
- No utility model call has tools.
- Email or compacted text cannot authorize drafting or broaden mailbox access.
- Narrow requests still search/read at most the current ten-message limit.
- Stop cancels enumeration, batching, and compaction without corrupting state.
- Partial processing reports exact coverage without pretending completion.
- Every request remains within the learned/effective endpoint budget.
- Gemini stays disabled and no Limits or agent controls appear in settings.

## Independent review questions

The reviewer should specifically challenge:

1. Whether the architecture is smaller than necessary and what can be removed
   from the first release.
2. Whether compaction can lose security-relevant state or accidentally transfer
   authorization.
3. Whether 500-message metadata enumeration and batched body access are viable
   with Outlook COM performance and cancellation behavior.
4. Whether the endpoint-budget strategy is reliable across OpenAI-compatible
   servers that omit usage or return unusual errors.
5. Whether evidence survives long tasks without retaining excessive personal
   data.
6. Whether failure and retry paths can duplicate work, inflate cost, or loop.
7. Whether the tests and exit criteria are sufficient to begin implementation.

## References

- Gemini CLI compression service:
  https://github.com/google-gemini/gemini-cli/blob/main/packages/core/src/context/chatCompressionService.ts
- Gemini CLI structured compression prompt:
  https://github.com/google-gemini/gemini-cli/blob/main/packages/core/src/prompts/snippets.ts
- Qwen3.8-27B model information:
  https://huggingface.co/Qwen/Qwen3.8-27B-FP8
