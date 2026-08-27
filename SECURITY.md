# Security model

## Security objective

Untrusted email, document, webpage, screenshot, MCP, user-prompt, conversation,
and model content must never expand the locally available capabilities or reach
an Outlook send or source-message mutation capability. Model
tool calls may select bounded read-only mailbox context. The only mutation this
add-in permits is creating one unsent draft after explicit, request-scoped local
authorization and updating that same locally linked unsent item after user
feedback.

## Capability separation

```text
Prompt + optional selected-message, ten-email working set, and bounded files
    |
    v
OpenAiCompatibleClient -> messages + read-only tool schema -> endpoint
    |
    v
allowlisted tool call -> MailboxToolHost -> bounded local Outlook reads
    |
    v
temporary handles + bounded untrusted text -> endpoint
    |
    v
bounded response -> local emphasis normalizer -> native RichTextBox spans

Latest user prompt passes the local drafting-intent policy
    |
    v
request includes create_draft + exact reply handle -> DraftToolHost -> consume once
    |
    v
DraftService -> Save + Display one unsent Outlook draft

Later user feedback passes the local revision-intent policy
    |
    v
request includes update_draft -> DraftToolHost -> consume once
    |
    v
SafeDraftHtml -> remove Markdown markers + encode text + fixed visual layout -> Save + Display same item

Explicit Settings click -> clean at most 15 recent Sent Items samples
    |
    v
no-tool style request -> editable local writing profile -> draft-only prompt data
```

`OpenAiCompatibleClient` has no reference to the Outlook application object or
`DraftService`. The mailbox host remains separate from `DraftService`. The
dedicated draft host can reach only the internal draft service. Creation requires
an atomic one-request authorization created by deterministic local intent rules
from the latest user-written prompt. Email bodies and model output do not enter
that decision. Updates are available only for locally recognized revision intent
while the host retains exactly one linked draft, and remain limited to one
mutation attempt per user request.

The Edge/Chrome path is separate and read-only:

```text
Explicit toolbar, context-menu, or attach-button gesture
    |
    v
temporary activeTab access -> bounded selection/page text/visible screenshot
    |
    v
fixed extension origin -> framed per-user native bridge -> BrowserChatService
    |
    v
configured endpoint + explicitly allowlisted read-only MCP tools only
```

The extension has no host permissions, content script, remote code, browsing
history, cookie, download, debugger, or browser-management access. It cannot
click, navigate, fill forms, enter credentials, upload, download, purchase,
post, message, or mutate a page. The native host accepts only the fixed public
extension identity and exposes neither browser control nor Office/mailbox tools.
MCP is disabled in browser chat unless the user enters exact tool names and
affirms that each is read-only; only one approved server and one call are
allowed per request.

## Enforced invariants

1. Without a working set, the model request schema exposes exactly
   `search_mailbox`, `read_messages`, and `read_thread`. With a locked working
   set, it exposes only `read_messages` and only accepts its `context1` through
   `context10` handles. It exposes `create_draft` only when local
   code recognizes drafting intent in the latest user prompt. Once a draft is
   linked, recognized revision intent may expose `update_draft` instead. It
   never exposes both.
2. `MailboxToolHost` has one public dispatcher and rejects any tool name outside
   that compile-time allowlist.
3. Model-selected searches are limited to one search of the primary Inbox and
   Sent Items per request and return no more than ten summaries. No request can
   load more than ten unique message bodies, including thread reads. Body
   lengths, calls per round, and tool rounds are also capped.
4. Search results receive temporary handles. Read operations accept only handles
   issued within the current request, plus the optional `selected` handle or a
   locally approved ten-email working set.
   Reply creation also requires one of those exact handles. Missing, expired,
   and fabricated handles are rejected without consuming draft permission, and
   the host never substitutes the selected or latest item.
5. The mailbox host has no reference to `DraftService`, and the endpoint client
   has no Outlook application object.
6. Response text is stripped of control characters and truncated before display
   or drafting.
7. A `RichTextBox` displays the response as literal text. No browser or HTML
   renderer is used.
8. `DraftService` and `DraftSession` are internal implementation types. The
   public draft host exposes state plus one guarded dispatcher and no send path.
9. `DraftToolHost` accepts only `create_draft` and `update_draft`, requires a
   draft operation to be the only tool call in that response, bounds every
   field, rejects unknown properties, and atomically consumes local permission
   before mutation.
10. A chat can link at most one draft. A request can make at most one creation
    or update attempt. Starting a new chat releases the COM link without deleting
    the unsent Outlook item.
11. Draft operations call Outlook save and display behavior only. Subject and
    recipient fields are bounded single-line text. Shared local code removes
    Markdown emphasis markers and HTML-encodes the remaining text. It may add
    only fixed paragraph, heading, subheading, list, divider, and `<strong>`
    elements. At most 12 exact phrases may be bolded. BCC and arbitrary model
    HTML are not accepted.
12. Source scans fail on Outlook send, delete, move, Outbox, or send/receive
   capabilities.
13. Conversation history, the active working set, and external context are held
    in memory and cleared by **New** or Outlook shutdown. `/search clear` removes
    only the current selection and working set. The visible **Clear** action
    removes both mailbox and external context.
14. `/search` is parsed and executed locally without calling the endpoint. It
    stores only the newest ten matching metadata records. A later normal prompt
    exposes only those handles, so the model cannot broaden the approved set.
    Ctrl+click multi-selection uses the same one-to-ten normalization and cap.
15. The chat never evaluates Markdown or HTML. A bounded local parser removes
    emphasis markers and produces plain text plus bold character ranges. The
    RichTextBox applies those ranges natively. The draft path consumes the same
    ranges but continues to HTML-encode all text before inserting fixed
    `<strong>` elements and other compile-time visual layout tags.
16. External context requires an explicit file selection or drop. It is limited
    to three supported text files, 2 MB per file before reading, 12,000 text
    characters per file, and 24,000 total. It is labeled as untrusted reference
    data and cannot add instructions or capabilities.
17. Writing-style analysis never runs automatically. It requires an explicit
    Settings action, reads at most 15 recent usable Sent Items messages, removes
    obvious quoted history, and uses a no-tool model request. The generated
     profile is visible and editable. It is added only to locally authorized
     draft requests and is subordinate to every capability boundary.
18. Browser page access requires a user gesture and the temporary `activeTab`
    grant. Context remains visibly attached in the side panel until the user
    clears it; it is never collected from other tabs or in the background.
19. Browser context is capped at 16,000 selection characters, 48,000 page-text
    characters, a validated 5 MB JPEG/PNG/WebP screenshot, 12 history turns, and
    a 16,000-character prompt. Native framing and responses have independent
    hard byte caps and timeouts.
20. The native-host manifest allowlists one stable extension origin. The host
    independently requires that exact origin argument, uses strict binary
    stdin/stdout framing, and returns no settings secrets or stack traces.
21. `BrowserChatService` can expose only exact, case-sensitive MCP tool names
    the user separately allowlisted and affirmed as read-only. Browser MCP is
    default-off, limited to one approved server, one call, one tool round, and
    shorter per-operation timeouts. It rejects every other requested tool and
    labels page, screenshot, and MCP results as untrusted data. It has no Outlook
    or document host object and no page-control API.
22. Model and webpage output is inserted using DOM `textContent`; it is never
    parsed as HTML or evaluated as script.

The system prompt reinforces these limits, but no security property depends on
the model obeying it.

Classic Outlook COM add-ins do not have a permission manifest that can deny a
`Send` scope. AI365's guarantee is capability-based: its source and compiled
assembly contain no Outlook `Send`, `Submit`, or send/receive invocation; its
model tools expose no such operation; and the model client never receives the
Outlook application object. CI scans both source and compiled IL for this
boundary. Replacing the installed binary or compromising the Windows process is
outside this threat model.

## Secrets

The API key is encrypted with Windows Data Protection API using the current-user
scope. The encrypted value is stored in:

```text
%LOCALAPPDATA%\OutlookLocalAIChat\settings.json
```

Any process running as the same Windows user can potentially invoke DPAPI and
recover current-user secrets. This protects the key at rest from casual file
inspection, not from a compromised user session.

The key is sent in the Authorization header of the configured endpoint. HTTPS
protects that header and submitted mailbox context in transit. Loopback HTTP is
permitted automatically for local model servers. Non-local HTTP requires an
explicit persisted opt-in in Settings and displays a warning because the API
key, prompts, and retrieved email context are then sent without transport
encryption.

The browser extension never receives the API key, Gemini refresh token, or MCP
headers. The native bridge loads them under the same current-user process and
sends them only through the existing configured provider/MCP clients.

The Settings endpoint check may send the same Authorization header to
`GET /v1/models`. It then submits a synthetic tool-call request containing no
selected-message metadata, email bodies, or mailbox search results. The returned
tool call is validated but never executed.

## Logging

The diagnostic log records UTC time, operation name, exception type, diagnostic
code, and HRESULT category. It does not record email content, prompts, provider
response bodies, endpoints, request IDs, or API keys.

## Installation trust

The repository does not contain a code-signing certificate. Unsigned installers
and assemblies can trigger Windows warnings and may be blocked by corporate
application-control policy.

For organizational distribution:

1. Build in a controlled Windows pipeline.
2. Sign the DLL, browser native host, and installer with the organization's trusted code-signing
   certificate.
3. Publish hashes and retain build provenance.
4. Allowlist the publisher rather than a mutable file path.

## Out of scope

The design cannot guarantee safety if:

- the installed binary or registry entries are replaced;
- the Windows account is compromised;
- Outlook, .NET Framework, or Windows has an exploitable vulnerability;
- Edge, Chrome, the extension platform, or the native-messaging channel has an
  exploitable vulnerability;
- another Outlook add-in modifies the draft after creation;
- a user-configured MCP server exercises capabilities outside AI365's own
  read-only browser/Office hosts;
- the configured AI endpoint mishandles or retains submitted data.

Review the endpoint provider's privacy, retention, and data-residency controls
before using real work email.
