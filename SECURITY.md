# Security model

## Security objective

Untrusted email, document, local Topic, webpage, screenshot, MCP, user-prompt, conversation,
saved Local skill, packaged Public skill, and model content must never expand the locally available capabilities or reach
an Outlook send or source-message mutation capability. Model
tool calls may select bounded read-only mailbox context. The only mutation this
add-in permits is creating one unsent draft after explicit, request-scoped local
authorization and updating that same locally linked unsent item after user
feedback.

## Capability separation

```text
Prompt + optional selected-message, bounded email working set, and bounded files
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

The Chrome path is separate:

```text
Explicit toolbar, context-menu, or attach-button gesture
    |
    v
active-tab capture (title/URL/selection/page text) + model-driven
http/https operation of up to five registered Scribble work tabs
    |
    v
fixed extension origin -> framed per-user native bridge -> BrowserChatService
    |
    v
configured endpoint + native BrowserActionPolicy + browser tools
```

The extension holds http/https host permissions and the required `debugger`
permission. That is an intentional capability-class change: compromise of the
extension would be more powerful than its exposed tool surface. Compensating
controls keep `chrome.debugger` calls in `background.js`, require side-panel
port registration of a live Scribble work-tab ID, allow only
`Input.dispatchMouseEvent`, `Input.dispatchKeyEvent`, and `Input.insertText`,
and detach in `finally` after each atomic action and on port disconnect. The
active tab remains read-only. Inspection uses `chrome.scripting`, traverses
open Shadow DOM and same-origin frames, leaves cross-origin frames opaque, and
never reads values from sensitive fields.

Typing is itself an exfiltration channel because page JavaScript can observe
keystrokes before submission. Typed values are therefore capped at 200
characters and may use normalized words from the user request or clarification
answers, or a small compile-time map of public aliases such as Dubai to DXB.
Any other inferred public term requires the user to approve that exact text in
an `ask_user` card before it can be typed. Values are evaluated independently
by the DOM-independent native
`BrowserActionPolicy`. Before dispatch, the value appears once in a
plain-language Pixel Pal status and remains in the bounded internal tool
transcript; raw refs are never rendered as chat activity. Search
queries are normalized to user-supplied tokens; harmless
reordering and singular/plural differences do not cause the search to fail, while
page-only tokens are removed before dispatch. The policy blocks credential, personal/traveler identity,
payment, purchase/booking, messaging, upload/download, and destructive fields,
forms, and controls, including Enter or a benign-looking button in a sensitive
form. Passenger-count controls remain allowed. `ask_user` pauses at the panel
for material ambiguities. The native host accepts only the
fixed public extension identity; its write-shaped capabilities are opening
one unsent Outlook draft and one new unsaved Excel workbook per request -
both visible, both left for the user to review, and neither able to send
or save anything. MCP is disabled in browser chat
unless the user enters exact tool names and affirms that each is read-only.

## Enforced invariants

1. Without a working set, the model request schema exposes the suite-wide
   read-only `ask_user` prompt helper plus `search_mailbox`, `read_messages`,
   and `read_thread`. With a locked working set, it exposes `ask_user` and
   `read_messages`, and only accepts its `context1` through
   `context10` handles. It exposes `create_draft` only when local
   code recognizes drafting intent in the latest user prompt. Once a draft is
   linked, recognized revision intent may expose `update_draft` instead. It
   never exposes both.
2. `MailboxToolHost` has one public dispatcher and rejects any tool name outside
   that compile-time allowlist.
3. Model-selected searches are limited to one search of the primary Inbox and
   Sent Items per request and return no more summaries than the configured
   working-set size. No request can load more unique message bodies than that
   size, including thread reads. Body lengths, calls per round, and tool
   rounds are also capped.
   Optional inclusive `received_after` and `received_before` filters and the
   `unread_only` filter narrow results only. They do not add mark-read or any
   other source-mail mutation. A capped search may inspect one extra matching
   metadata item solely to report `truncated: true`; it never exposes or reads
   that extra message body.
4. Search results receive temporary handles. Read operations accept only handles
   issued within the current request, plus the optional `selected` handle or a
   locally approved email working set (sized by the user in Settings >
   Limits, default ten).
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
    stores only the newest matching metadata records, bounded by the
    configured working-set size. A later normal prompt exposes only those
    handles, so the model cannot broaden the approved set. Ctrl+click
    multi-selection uses the same normalization and cap.
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
18. Browser context is the active tab of the panel's own window, captured at
    send time and shown in the panel header; beyond that, reads touch only
    Scribble's own work tabs. Model-driven navigation is restricted to
    user-supplied http/https URLs in up to five Scribble-created background tabs; open-ended discovery uses Google's visible UI and observed result refs - the
    user's current tab is never navigated - and every tab is a normal,
    visible browser tab the user can inspect or close.
19. Browser context is capped at 16,000 selection characters, 48,000 page-text
    characters, a validated 5 MB JPEG/PNG/WebP screenshot, 12 history turns, and
    a 16,000-character prompt. Native framing and responses have independent
    hard byte caps and timeouts.
20. The native-host manifest allowlists one stable extension origin. The host
    independently requires that exact origin argument, uses strict binary
    stdin/stdout framing, and returns no settings secrets or stack traces.
21. `BrowserChatService` exposes the fixed browser tools (navigate, read,
    Google UI search, snapshot, act, extension-validated evidence, the shared
    `ask_user` prompt helper, one unsent Outlook draft per request, and one unsaved Excel
    workbook per request) plus exact, case-sensitive MCP tool names the
    user separately allowlisted and affirmed as read-only. Progress is measured
    from stable page-state fingerprints. Twenty consecutive browser calls with
    no meaningful state change stop as a loop; progressing work may continue up
    to a 120-round emergency cost/safety fuse, with four calls but no more than
    one state-changing browser call per round. It rejects every
    other requested tool and
    labels page, screenshot, and MCP results as untrusted data. Every page
    read returns a bounded link list so multi-step navigation follows exact
    URLs and ref-scoped controls instead of guessing. Older browser results are
    compacted while clarification answers, validated evidence, and the six
    newest snapshots remain full. Its only
    Office capabilities are `OutlookDraftLauncher` (displays one unsent draft;
    no send, read, delete, or mailbox access) and `ExcelTableLauncher` (opens
    one brand-new unsaved workbook with a bounded table and optional chart; no
    save, print, protect, or close capability, and it never touches existing
    files). Page control is limited to the separately authorized background
    operator described above.
21a. The Excel, PowerPoint, and Word panes expose `fetch_web_page`: one
    bounded read-only HTTP GET per call (http/https only, no cookies, no
    credentials, 3 MB / 48,000-character caps, results marked untrusted).
    The Outlook mailbox pane deliberately does not get this tool, so
    attacker-authored email text can never choose a URL sink for mailbox
    data.
21b. Every Scribble pane exposes `ask_user`. Office panes advertise one bounded
    question with at most four bounded options plus free text. Browser chat
    advertises one to three related questions in one card and returns answers
    keyed by question id. The shared parser accepts both shapes during version
    skew and resumes only with the user's answer. A local
    structural preflight forces this tool for a narrow set of obviously vague
    prompts. Mixed `ask_user` plus action-tool rounds are rejected without
    running any of the requested actions; Stop cancels a pending question.
21c. A local Topic is inert until the user selects it for a new chat. The
    selection locks after the first message. `search_topic` may run once per
    request and returns at most ten opaque handles; `read_topic_files` accepts
    only handles bound to that Topic, chat, and request and reads at most three
    documents / 120,000 characters.
21d. Topic roots must be existing non-network local folders. Recursive indexing
    skips hidden, system, and reparse-point files and directories, canonicalizes
    and revalidates containment before extraction, and never accepts a path from
    the model. Absolute paths and cached text are excluded from model output and
    logs.
21e. Topic indexes are bounded plaintext atomic caches under the current user's
    LocalAppData. Removing or repointing a Topic removes only its cache; source
    repositories are read-only. Relevant excerpts are transmitted to the
    configured model with the same untrusted-data treatment as attached files.
21f. User-selected and Outlook attachments are accepted up to 100 MiB each and
    share a 250 MiB source budget per operation. Source files remain local.
    Parsers stream or seek through file-backed input, cap decompressed archive
    parts and PDF streams, reject unsafe image dimensions, and retain the
    existing 48,000-character per-file, 120,000-character total, and 800 KiB
    vision-image payload limits. Topic indexing remains capped at 25 MiB.
21g. Excel's **Send to Scribble** context-menu command captures one contiguous
    range before focusing the task pane. The attachment is inert; only a later
    user prompt containing an edit action can unlock writes. An eligible
    single-column attachment mints an opaque request-scoped handle which
    expires after the request and is invalidated by context removal or a new
    chat. `write_selection_output` stages at most 500 one-to-one literal values
    in five bounded calls without consuming permission. Excel selection
    requests may use up to eight bounded tool rounds so five staging calls,
    one clarification, and a rejected retry can still finish. The final call
    revalidates saved `FullName` (or unsaved workbook name), window handle,
    worksheet name, source address, exact row count, and either a fully blank,
    unmerged, formula-free destination or an explicitly authorized exact-source
    replacement before consuming one draft permission and performing one bulk
    write. Values beginning with `=`, `+`, `-`, or `@`
    are forced to inert text. Source cells are preserved unless the user's own
    prompt or `ask_user` answer explicitly says replace, overwrite, or in place;
    that choice remains bound to the captured source range and the same final
    identity checks. No file is saved.
21h. Browser action policy is operation-aware. Sensitive-name classification
    applies to value entry rather than ordinary links or category-card clicks;
    password and file controls remain denied regardless of operation. A safe
    HTTP(S) anchor labelled Buy may navigate, while a button/submit with the
    same text remains consequential. Password forms, sensitive fields,
    personal/payment submits, authentication, upload, messaging, download,
    destructive actions, and final purchase/trade-in submission are hard
    denials with no model-controlled confirmation override.
21i. A completed browser price, valuation, availability, or configured-product
    claim requires `browser_record_evidence`. The extension validates the
    bounded fields against its current work tab, revision, fingerprint, visible
    DOM text, and verified action receipts, then sends the structured record to
    the native host as untrusted data. Stale revisions, different tabs,
    invented or Google result URLs, absent excerpts, unresolved actions, and
    answer/record mismatches cannot complete. Work tabs survive until Clear chat
    so **Open evidence tab** can reactivate the exact observed page.
22. Model and webpage output is never parsed as HTML or evaluated as
    script. Assistant replies pass through a bounded local formatter that
    builds paragraph, list, table, bold, and code DOM nodes itself and
    inserts every piece of text with `textContent`; user and page text is
    inserted as literal `textContent` directly.

The system prompt reinforces these limits, but no security property depends on
the model obeying it.

Classic Outlook COM add-ins do not have a permission manifest that can deny a
`Send` scope. Scribble's guarantee is capability-based: its source and compiled
assembly contain no Outlook `Send`, `Submit`, or send/receive invocation; its
model tools expose no such operation; and the model client never receives the
Outlook application object. CI scans both source and compiled IL for this
boundary. Replacing the installed binary or compromising the Windows process is
outside this threat model.

## Secrets

The API key is encrypted with Windows Data Protection API using the current-user
scope. The encrypted value is stored in:

```text
%LOCALAPPDATA%\Scribble\settings.json
```

Any process running as the same Windows user can potentially invoke DPAPI and
recover current-user secrets. This protects the key at rest from casual file
inspection, not from a compromised user session.

Direct Gemini is disabled in the standard build. Dormant Gemini tokens are not
decrypted during settings load and are removed from the settings file on the
next successful save.

The key is sent in the Authorization header of the configured endpoint. HTTPS
protects that header and submitted mailbox context in transit. HTTP is accepted
without a separate opt-in so local and LAN-hosted model servers are easy to
configure. Settings displays a prominent warning for non-loopback HTTP because
the API key, prompts, and retrieved email context are then sent without
transport encryption. Use HTTPS for every endpoint outside a trusted local
development setup.

The browser extension never receives the API key, dormant Gemini credentials, or MCP
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
- Chrome, the extension platform, or the native-messaging channel has an
  exploitable vulnerability;
- another Outlook add-in modifies the draft after creation;
- a user-configured MCP server exercises capabilities outside Scribble's own
  read-only browser/Office hosts;
- the configured AI endpoint mishandles or retains submitted data.

Review the endpoint provider's privacy, retention, and data-residency controls
before using real work email.
