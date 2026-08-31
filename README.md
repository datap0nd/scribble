# Scribble

A Windows-only AI assistant suite for classic Microsoft Office
(Professional Plus 2021) and Google Chrome: one installer adds
an **Scribble** sidebar to **Outlook, Excel, PowerPoint, Word, and the web**. Every
pane shares the same chat stack -
an OpenAI-compatible endpoint, the same
settings and writing soul, rich markdown output with tables, optional MCP
tool servers, and the same hard guardrails: the model can read bounded
context but can never send email, save a file, or delete anything.

- **Outlook**: chat with the local Inbox and Sent Items, read attachments and
  images, and open or revise one linked **unsent** draft for human review.
- **Excel**: chat with the open workbook (bounded sheet and cell reads); the
  only write surface is a clearly marked **Scribble Draft** worksheet that the
  add-in never saves.
- **PowerPoint**: chat with the open presentation (bounded slide and notes
  reads); the only write surface is **[Scribble draft]** slides that the add-in
  never saves, laid out and styled by the built-in **METO corporate theme**.
- **Word**: chat with the open document (bounded text reads); the only write
  surface is a brand-new, unsaved **[Scribble draft]** document that the add-in
  never saves.
- **Chrome**: chat about the tab you are on - its title, address,
  selection, and readable text are attached automatically - and let Scribble
  browse http/https pages in up to five of its own background work tabs;
  your current tab is never navigated away. When
  you ask, it opens an unsent Outlook draft or a new unsaved Excel workbook
  with a table and chart built from what it found. A Stop button halts a
  running chain of steps, and it asks you a clickable clarifying question
  when something important is ambiguous. It can click one benign, visible
  control at a time (cookie banners, location choosers, continue) - never
  anything that buys, signs in, or pays - and it never types into fields,
  fills forms, downloads, uploads, or sends anything.
- **The web from Office**: the Excel, PowerPoint, and Word panes can fetch
  http/https pages read-only (`fetch_web_page`) - bounded text and links, no
  cookies, no sign-ins, no forms - so "get the top 5 iPhone prices from
  amazon.ae and chart them" works straight into a Scribble Draft sheet. The
  Outlook pane deliberately has no web access: mailbox text plus an
  attacker-chosen URL would be a data-exfiltration channel.
- **Connected**: from any document pane, "email this to ..." opens an
  unsent Outlook draft (optionally attaching the saved file); "put this in
  PowerPoint" / "put this in Excel" / "put this in Word" drafts into the
  sibling app; and a deliberate **Share to Scribble apps** hand-off moves one
  bounded snippet between panes. One **Update** click in Settings refreshes
  the whole installed suite together.

It does not use Microsoft 365 add-in deployment, Microsoft Graph, or
Entra ID.

## Install

1. Close Outlook, Excel, PowerPoint, and Word.
2. Download
   [ScribbleSetup.exe](https://github.com/datap0nd/scribble/releases/latest/download/ScribbleSetup.exe).
   This link tracks the **Latest** release, which is rebuilt automatically on
   every push to `main`.
3. Run the installer for your Windows account. It asks which apps get Scribble -
   all five are selected by default; untick any you do not
   want. Re-running the installer later lets you change the selection, and
   deselected apps are cleanly unregistered. Silent installs (and the
   in-app updater) keep your previous selection.
4. If you selected Chrome, leave **Finish setting up Scribble in Google
   Chrome** ticked. Setup opens the Extensions page and the exact Scribble folder.
   Turn on **Developer mode**, select **Load unpacked**, and choose that folder.
   The installer prepares the
   files and secure local bridge, but the browser must receive this approval
   from you because Scribble is not in a browser store. If a managed computer
   disables Developer mode, your IT policy must allow the extension.
5. Start classic Outlook, Excel, PowerPoint, or Word.
6. In Outlook choose **Scribble > Scribble** on the ribbon; in Excel, PowerPoint,
   and Word the **Scribble** button sits on the **Home** tab.
7. Open **Settings** and enter:
   - the OpenAI-compatible endpoint or base URL;
   - the API key.
8. Click **Connect & load models** to load available model IDs from
   `GET /v1/models`, then choose or type a model that supports
   OpenAI-compatible chat tool calls. If the field is empty, Scribble prefers
   a suitable discovered Qwen model, with Qwen3.8-27B preferred when available.
   If there is only one other usable model it is selected; otherwise the choice
   remains with the user.
9. Click **Test selected model**. Save only after authentication, the selected
   model, and mailbox tool calling pass.
10. Optional: open **Writing style**, click **Analyze 15 sent emails**, review
   the generated drafting instructions, edit them, and enable the profile.

To update later, open **Settings** in any Scribble pane and click
**Update Scribble**. You confirm twice - the second dialog warns that the
Office apps are about to close, so save your work first - and Scribble then
downloads the latest installer, **closes Outlook, Excel, PowerPoint, and
Word itself**, and installs silently for your Windows account. Chrome
stay open. Only the
apps that actually have Scribble installed are closed; a host still sitting
on a save prompt after about thirty seconds is closed forcibly, so an
update can never stall unfinished. Outlook reopens automatically when the
update was started there. One update refreshes the whole suite. If the browser
extension changed, open `chrome://extensions`, find
Scribble, and click **Reload**; unpacked extensions do not reload changed files
automatically.

Settings shows the **Installed version** (for example `2.0.27.0`); the
release page states the version it publishes, so you can confirm an
update actually landed.

Examples:

```text
https://ai.example.test/v1
https://ai.example.test/v1/chat/completions
http://127.0.0.1:1234/v1
```

HTTP and HTTPS URLs are accepted without a separate opt-in. HTTPS is strongly
recommended outside the local computer. Settings shows a warning for remote
HTTP because the API key, prompts, and retrieved email context then cross the
network without transport encryption. Loopback HTTP such as `localhost` and
`127.0.0.1` remains the normal setup for a local Qwen server.

The first unsigned build may trigger a Windows SmartScreen warning. A trusted
code-signing certificate is required to remove that warning for normal company
distribution.

The chat sidebar renders in Microsoft Edge WebView2, which ships with
Windows 10/11 and Microsoft Edge. The embedded page is network-isolated by
CSP, never navigates anywhere, and inserts all model and mailbox text as
inert text nodes - model output is still never parsed as HTML.

## Model choice

Direct Google Gemini access is unavailable to end users. The retained Gemini
implementation is protected by a default-off product gate and request-level
checks; Gemini model IDs are removed from saved and discovered model lists.
Every selectable model is sent only to the configured OpenAI-compatible
endpoint.

After you enter an endpoint and API key, use **Connect & load models** in
Settings to populate the dropdown from `GET /v1/models`. Scribble prefers a
discovered Qwen model when no model is already selected. You can still type any
endpoint model ID manually. Every
dropdown entry is tagged **Vision** (reads email images) or **Text**
(filename-only), and the sidebar header shows the same tag for the saved model.

For email **images**, pick a model tagged **Vision**. Vision capability is
detected from the model ID: `vl` or `vision` in the name (for example
`qwen3-vl-30b`), multimodal Gemma generations (`gemma3`/`gemma-4` and later),
and common vision families such as LLaVA, Pixtral, MiniCPM-V, InternVL,
Moondream, and SmolVLM. Scribble loads image attachments automatically and sends
them as multimodal input, capped at eight images per request. Multimodal Gemma
requires the server to load its vision projector; if the server rejects image
input, use a `vl` model instead. Text-only models get spreadsheet text and
image metadata only, and the chat will say so if you ask about an image.
Optional: enable **Auto-switch to vision for images** to temporarily use a
discovered vision model while keeping your everyday text model saved. Save after
**Connect & load models** so auto-switch knows which vision models are available.

Embedding-only and Gemini models are excluded from discovery. **Test selected model** verifies
authentication with a lightweight `search_mailbox` tool-call probe. It does not
read mailbox data during the check. Model discovery allows up to 15 seconds;
the tool-call probe allows up to 90 seconds.

## Use

1. Open **Scribble**. The chat appears as a sidebar inside Outlook and starts
   with **no context** - opening the pane never pulls in whatever email
   happens to be selected. You add context deliberately: drag messages onto
   the pane, use **Add email**, right-click messages and choose **Send to
   Scribble**, or run `/search`. Common `RE:`, `FW:`, and `FWD:` prefixes are
   hidden where a subject is shown.
   Right-click works inside the pane for the usual copy, paste, and select-all
   actions.
2. To choose a bounded group first, enter `/search person or topic`. Scribble
   searches locally and keeps the newest matching Inbox or Sent Items
   emails as the working set (up to the per-request size configured in
   Settings > Limits; the default is ten). No email body is sent during this
   command.
3. Review the listed subjects and send another `/search` to replace the set if
   it is wrong. Results appear in a collapsible working-set layer as
   distinct email cards with subject, sender, and date. Use `/search clear` to
   remove it. The layer collapses automatically when you send a normal AI
   prompt and can be reopened with **Show**.
4. Alternatively, Ctrl+click emails in Outlook (up to the working-set
   limit), then choose
   **Add email**, right-click **Send to Scribble**, or drag the selected messages
   onto the Scribble pane. Multiple messages become the same locked working set.
5. Use the **+** menu or drag files from Windows Explorer to add external
   context: up to three documents and four images. Documents go through the
   same extractors as email attachments (PDF, Office, text formats), and
   images become vision input with a tray thumbnail. HTML files are read as
   inert text, not rendered. Each file may be up to 25 MB on disk; extracted
   text is bounded to 48,000 characters per document and 120,000 characters
   total. A file that exceeds a cap still appears in the tray as an amber
   warning chip explaining what was kept — oversized files are noted, and
   over-length text keeps its first 48,000 characters with a truncation
   notice the model can see.
6. Ask a normal mailbox question. When a working set exists, the model can read
   only those emails. Without one, it may perform one bounded mailbox search
   and load no more unique email bodies than the configured working-set size
   for the request. Meeting
   invites and calendar items are readable like email — subject, body, time,
   location, and attachments — but Scribble can never accept, decline, or
   schedule anything. When a body
   is loaded, Scribble also reads up to ten **email attachments** per
   message: images (PNG, JPEG, GIF, BMP, WebP, TIFF), spreadsheets (XLSX,
   XLSM, XLSB, XLTX, XLTM, XLS, CSV, TSV — all worksheets, including
   binary BIFF12 workbooks), documents (PDF, PPTX, PPTM, PPSX, PPSM,
   POTX, DOCX, DOCM, DOTX, DOTM, PPT, DOC, RTF, and OpenDocument
   ODT/ODS/ODP), attached Outlook messages (MSG, OFT — subject, sender,
   and body), and text files (TXT, MD, LOG, JSON, XML, YAML, INI, HTML,
   EML). Unknown extensions are identified by content (image, OOXML or
   binary Office, OpenDocument, MSG, PDF, or plain text). PDF
   extraction reads the text layer, including CID-font PDFs from Word and
   Chrome — scanned PDFs yield a clear "no readable text" note. Legacy
   binary Office files get best-effort extraction. Every attachment is
   listed; anything unreadable is noted rather than silently skipped.
   Attachments up to 25 MB are read; extraction is streamed and bounded
   to 48,000 characters per attachment and 120,000 characters per
   message, with an explicit truncation notice when more content
   remains.
   Small inline images embedded in the body (64 KB or less) are treated
   as signature graphics and ignored, with a note in the tool result;
   pasted screenshots and photos are far larger and are always read.
   Attachments are decrypted locally through Outlook COM before reading.
7. The sidebar records which bounded context operations ran.
8. Ask explicitly, for example "create a reply draft" or "write an email."
   Local code recognizes that drafting intent and exposes one creation attempt
   for that request. The draft opens unsent in Outlook. You can also
   right-click an email and choose **Scribble - Suggest a response**: the
   sidebar asks up to three quick questions (reply tone plus up to two
   model-suggested questions specific to that email), and your answers
   shape the reply draft. Skipping the questions goes straight to a
   draft. The composed request goes through the same drafting pipeline,
   so it still authorizes exactly one draft that opens unsent for review.
9. A mailbox question without explicit drafting language cannot expose draft
   creation. Loaded email text and model output cannot authorize it.
10. The same Outlook draft stays linked to that chat. Follow-up instructions such
   as "make it shorter" or "bold the deadline" update and redisplay that exact
   unsent item. No second draft is created.
11. Review, edit, address, and send the message using Outlook's normal editor.

Selecting an email is optional for mailbox questions. When one is selected, the
model receives its metadata and may request its body using the temporary
`selected` handle. A multi-email selection is stored as a locked working
set with `context1` through `contextN` handles (N is the configured
working-set size, default 10). The conversation and working set
remain in memory until cleared or Outlook closes. `/search clear` removes the
email working set but retains external files. **Clear** removes all context, and
**New** starts a new conversation with no retained context.

Settings is organized into four tabs: **Connection** (endpoint, API key,
model discovery, compatibility test, updates), **MCP**, **Writing style**, and
**Support** (describe a
problem and Scribble opens a pre-filled, unsent report email to the creator
with the recent diagnostic log — timestamps, operations, and error codes
only — for you to review and send yourself).

The **writing soul** is a small, editable portrait of how you write.
Analysis never runs automatically: it requires a click, reads at most 15
recent usable Sent Items messages, removes obvious quoted history, and
sends bounded samples to the configured AI endpoint. The result is visible
and editable before saving. A **soul strength** slider (10–100) controls
how strongly drafts follow your voice, and **hard draft rules** (one per
line) are followed exactly in every draft. Soul, strength, and rules apply
only to draft creation and revision, and only to wording, greeting,
cadence, and sign-off. They cannot alter any capability or security rule.

## Chrome

Click the Scribble toolbar button to open its side panel. The current tab's
title, address, selection, and readable text are captured automatically and sent
with every message - there is nothing to attach. The panel header shows which
tab is being shared; the selection is capped at 16,000 characters and page text
at 48,000. A right-click **Ask Scribble about this page** opens the panel too.

Scribble can also browse for you: ask it to look something up and the model
opens http/https pages in up to five of its own background work tabs
(`browser_navigate` with a tab number 1-5), re-reads them
(`browser_read_page`), and compares sites side by side across up to twelve
bounded tool rounds per request. Your current tab is never navigated away;
**Clear chat** closes Scribble's work tabs. Ask it to email someone and it opens one unsent Outlook
draft window for your review (`open_outlook_draft`); ask for Excel and it
opens one unsaved workbook (`open_excel_table`). Neither is ever sent or
saved by Scribble. When an interstitial blocks reading - a cookie banner, a
country or language chooser - Scribble can click that one visible control
(`browser_click`); a hard blocklist refuses anything that buys, checks out,
signs in, registers, subscribes, or submits a credential or payment form, and
typing into fields is impossible. When your request is ambiguous (location,
recipient, budget), it asks you one clarifying question with clickable
options (`ask_user`) and waits for your answer. It never fills or submits
forms, enters credentials, uploads, downloads, purchases, or posts. Browser settings
pages, extension galleries, and some protected viewers block page-text
extraction; those tabs are shared as address-only.

The **Settings** button in the panel opens the same Scribble Settings window as
the Office add-ins (it appears on your desktop). Messages go through a per-user
native bridge to the same Scribble connection and model configured there. The
extension never receives the API key, Gemini token, or MCP headers. Page
content is always labelled as untrusted data. MCP stays off in browser chat by
default: to use a web-search MCP, list its exact tool name under that server's
**Chrome tool allowlist** in Scribble Settings and tick the read-only
approval; never approve a tool that writes or takes actions.

To repeat the one-time setup, use **Set up Scribble in Google Chrome** from the
Scribble Start menu folder. After a Scribble update, click **Reload** on
`chrome://extensions` - Chrome requires this manual reload for an unpacked
extension whose installed files changed.

## Excel, PowerPoint, and Word panes

The Excel, PowerPoint, and Word sidebars reuse the same chat page, models,
streaming, context tray, and settings as Outlook. Their tool surface is
document-shaped and read-only:

- Excel: `list_worksheets` and `read_cells` (at most 500 rows x 50 columns
  per read, 500 characters per cell, tab-separated, truncation flagged).
- PowerPoint: `list_slides` and `read_slide` (bounded slide text including
  speaker notes).
- Word: `read_document` (bounded plain-text slices of the active document).
- **+ > Add current selection / Add current slide** snapshots what you have
  selected into the bounded context tray; files and pictures work exactly as
  in Outlook.

Writes unlock only when your own latest message asks to produce something
("put this in a draft sheet", "build a slide with this", "do a bar chart
with this in a slide", "put this table into word", "fix the formula in
column B", "email this to ..."). Asking to edit, fix, fill, chart, or move
content all count - no confirmation round-trips. Every write stays in
memory - the add-in has no save capability at all - and one request
produces one complete deliverable. A deck or workbook may be built over
a few bounded calls (a small local model cannot emit a whole dense deck
in one payload, so it adds it in batches instead of thinning it out);
an unsent email draft stays strictly one per request. By default output lands in a
clearly marked draft surface; when you explicitly ask to change the file
you are working on ("fill in the missing totals on my sheet", "continue
this document"), it goes there instead - still unsaved:

- `write_draft_sheet` fills a brand-new numbered **Scribble Draft** worksheet
  (Scribble Draft, Scribble Draft 2, ... at the end of the workbook). Earlier
  drafts and your own sheets are never touched, so a follow-up draft can
  never destroy a previous one - delete draft sheets you no longer want,
  or close without saving to discard everything.
  The table always lands at **A3** (title in A1) so formulas can reference
  it deterministically. Cells starting with `=` become **live Excel
  formulas** that may reference other sheets of the same workbook
  (`=SUM(Data!B2:B9)`) or the draft table itself; locale-style formulas
  fall back through `FormulaLocal`, numbers and dates are typed
  automatically, the title and header row are styled, and columns autofit.
  Ask for a chart and a **native Excel chart** is drawn below the table,
  sourced live from it (column, bar, line, pie, area, scatter). Formulas
  that could reach the network, native code, or other files (`WEBSERVICE`,
  `RTD`, `CALL`, `HYPERLINK`, external `[Book]` references) are rejected
  and land as visible text instead.
- `write_cells` (Excel only) writes values and live formulas **directly
  into your active worksheet** starting at a cell you name - offered only
  for explicit change-my-sheet requests ("fill in the missing totals",
  "fix the formulas in column D"). Same formula safety rules; nothing is
  saved, so closing without saving discards everything.
- `add_draft_slides` adds slides marked **[Scribble draft]**; existing slides
  are never modified. Every slide is painted from the corporate theme (see
  below). By default new slides append at the end, or ask for a position
  ("after slide 2", "at the start") and they insert there. A slide can
  carry bullets, a **strategy grid** of numbered cards, a **data table**
  with automatic performance highlighting, and a **native chart** built
  from data the model supplies - "do a bar chart with this in a slide"
  draws a real PowerPoint chart, not a picture or a text list.
- `write_draft_document` writes into Word with a placement you control:
  by default it **appends to the document you are working on** (Ctrl+Z
  undoes it), "replace the selection" rewrites just what you selected, and
  asking for a separate draft opens a brand-new unsaved document headed
  **[Scribble draft]**. `#`/`##`/`###` headings, `-`/`1.` lists, and
  `**bold**` render as real Word styles, and `| cell | cell |` rows become
  **real formatted Word tables** - "put this table into word with an
  analysis" produces a styled table plus prose.
- `create_email_draft` opens an unsent Outlook draft for review - Outlook
  starts if needed, the draft can attach the current file when it is saved
  on disk, and sending stays impossible. When the model picks a recipient
  you never mentioned in your request, the pane calls it out so you can
  check the address before sending. Email bodies support headings, lists,
  dividers, and `| cell | cell |` **tables** rendered as real bordered
  HTML tables.
- `send_to_powerpoint` / `send_to_excel` / `send_to_word` create a
  **brand-new unsaved draft file every time** - the sibling app starts if
  needed, a fresh workbook/deck/document opens, and nothing that happens
  to be open in that app is ever touched. The Outlook pane has them too:
  "create an excel", "create a word", "create a powerpoint", "build me a
  slide of my day", or "put these emails in excel" all work straight from
  your mailbox, first try.

### Corporate slide theme

Draft slides are not generic bullet pages. Scribble ships the **METO executive
deck theme** compiled into the add-in, and the model supplies content only -
it can never choose a font, color, size, or position, so even a small local
model produces on-brand slides:

- **Typography**: Samsung Sharp Sans Bold titles (40pt, stepped down when a
  title runs long), Samsung Sharp Sans Medium subtitles (15pt), Calibri body
  text (11pt), Arial chart labels (9pt), Malgun Gothic footnotes (7pt).
- **Palette**: brand blue `#1428A0` rules and accents, charcoal titles,
  pure white slides, slate `#E7ECF0` cards and table headers.
- **Layouts**, picked automatically from the content supplied: a royal-blue
  **cover**, a centered **agenda**, the three-column **strategy grid** with
  circled numbers, the dense **performance matrix**, a full-width **chart**,
  or bullets beside a chart.
- **Selective highlighting**: table cells marked with the corporate
  performance vocabulary - up/down arrows, the deficit triangle, or a signed
  number like `+12%` / `-8%` - are filled light green (growth) or light
  yellow (shortfall) automatically, capped at four cells each so a table
  stays low-noise. The model never picks a color.
- **Charts**: stacked, 100% stacked, clustered, bar, line-with-markers,
  pie, area, and scatter, always flat (never 3-D), with fine `#D9D9D9`
  gridlines, 9pt charcoal labels, brand-blue-led series colors, and an
  optional unit indicator such as `(K unit)`.

Ask in the corporate vocabulary and it carries through: takeaway titles,
`M/S`, `G/R`, `A/R`, `S/I`, `S/O`, `YTD`, `MP`, and arrow markers are all
part of the drafting instructions.

Decks are built to be **dense**: one slide can carry a table *and* a chart
*and* its takeaway bullets at once, every content slide is expected to hold
a table, chart, or card grid rather than bare bullets, and the model is
told to read a source document to the end before drafting and to keep
adding slides over a few calls instead of emitting one thin pass.

The add-in never calls Save, SaveAs, Delete, Print, Protect, Close, or Quit
on your documents - saving stays a human action, so even a discarded draft
sheet or slide costs nothing.

Office opens every document in its own window, so each window gets its own
Scribble pane - the ribbon button works in a freshly created draft document
too, not just the first window. Closing and reopening a pane keeps the
conversation for as long as that Office application stays open; all windows
of one app share the same chat, which lives only in process memory and is
forgotten when the app closes (nothing is written to disk). **New chat**
clears it immediately.

## Rich answers and the pixel pal

Assistant replies render as safe rich text: tables with alignment, nested
lists, headings, code blocks, quotes, and inline emphasis. The page builds
everything as inert DOM nodes - model output is never parsed as HTML, and
links never navigate (the URL shows on hover). While the model thinks, a
small pixel robot types away next to the usual dots.

## MCP tool servers

Settings has an **MCP** tab where you can register up to eight Model Context
Protocol servers - local commands (stdio) or HTTP(S) endpoints. Their tools
appear to the model as `mcp_<server>_<tool>` (40 tools max), every result
comes back bounded and marked as untrusted data, and slow servers are timed
out (stdio servers are killed rather than left blocking).

Only you can add a server: nothing in email, document, or model text can
register one. A server runs with your Windows account's own permissions,
outside Scribble's guardrails - Scribble itself still cannot send email or save or
delete documents, but a server you add acts with whatever powers it has.
Only add servers you trust, and prefer read-only ones.

Browser chat has a separate, default-off boundary. For one server only, you may
enter up to 20 exact, case-sensitive MCP tool names in its **Chrome tool
allowlist** and affirm that you verified them as read-only. Only those names can
be exposed there, with at most one call in one tool round. Adding a server for
Office does not automatically make any of its tools available to webpage
content.

HTTP(S) servers can carry per-server request headers (one per line as
`Name: value`, typically `Authorization: Bearer ...`). Headers are sent only
to that server's own endpoint, never logged, and stored DPAPI-encrypted like
the API key.

## Request budgets

Scribble uses reviewed, non-editable per-field text budgets as a conservative
Qwen baseline: 4,000 prompt characters, 12,000 retained answer characters, 12
history turns, six tool rounds, and four calls per round. Legacy custom values
for those budgets are ignored on load and replaced on save.

The number of emails per request is yours: **Settings > Limits > Emails per
request** sets the working set, search-result, and request-wide body-loading
size, from 1 to 10,000 (default ten). Scribble's tool-result boundary scales
with your choice. Be deliberate with large values: hundreds of email bodies can
reach millions of characters, overflow a small model's context window, and slow
requests down. These values do not reveal or guarantee the server's actual
context window; attachment-heavy multi-round requests still depend on the
endpoint's configured limit.

## Diagnostics and administration

**+ > Copy diagnostics** in any pane copies a bounded report of the last
five requests to the clipboard: whether the local intent gate unlocked
drafting, which tools were exposed to the model, every tool call with its
status, and how the request ended. It contains no API keys, settings, or
message bodies - paste it into a bug report to show exactly what happened.

Direct Google Gemini is disabled for end users in this build. The registry
value `DisableGemini` = `1` (DWORD) under
`HKLM\Software\Policies\Scribble` or `HKCU\Software\Policies\Scribble` remains
a defense-in-depth kill switch for a possible future managed build. Policies
can only remove capabilities, never add any.

## Hard security boundary

The model is not given general Outlook access. Draft creation is a narrowly
scoped exception, not a general mutation permission.

- A request without a working set exposes three read-only tools:
  `search_mailbox`, `read_messages`, and `read_thread`. A request with a locked
  working set exposes only `read_messages`, and only for its approved
  handles.
- `create_draft` is added only when local code recognizes an explicit drafting
  instruction in the latest user-written prompt, such as "create a draft" or
  "write a reply." Model output and email content never enter that decision.
- Once one draft exists, `create_draft` disappears and only `update_draft` is
  eligible. Local code exposes it only for a recognized revision instruction,
  and it can mutate only that linked unsent item once per user request.
- The draft host requires `create_draft` to be the only tool call in its model
  response, validates strict arguments, and atomically consumes permission
  before creating anything.
- One chat can link at most one draft, and a request may open at most one
  unsent email draft. Document deliverables (slides, sheets, documents) may
  be written over a small, fixed number of bounded calls within that same
  authorized request - the permission is granted once, by the user's own
  wording, and every call still lands in a marked, unsaved surface.
- The local hosts reject every other tool name and cap tool calls, tool rounds,
  result counts, message bodies, draft fields, and total returned context.
- Search results use temporary handles. The model cannot submit arbitrary COM
  objects, Outlook commands, or executable code.
- A reply draft must include the exact temporary handle for its source email.
  The local host rejects missing, expired, or invented handles and never falls
  back to the selected or latest mailbox item.
- The model client never receives the Outlook application object or draft
  service.
- Model output is length-limited text displayed in a Windows control. It is
  never evaluated, executed, or rendered as model-provided HTML. Local code removes Markdown
  emphasis markers and applies only native bold spans in the transcript.
- Only a one-request authorization derived locally from explicit user drafting
  intent can create the linked draft. Later revisions require both recognized
  revision intent and that local linked-draft session.
- The model-invoked mailbox host remains read-only. A separate draft host accepts
  only the bounded `create_draft` and `update_draft` operations.
- The draft path exposes no send, move, delete, schedule, BCC, arbitrary HTML,
  or mailbox traversal operation.
- Classic Outlook COM has no permission-manifest switch for sending. Scribble
  instead hardcodes the absence of a send capability, keeps the Outlook object
  outside the model client, and verifies the source plus compiled assembly in CI.
- Drafts are saved and displayed as unsent Outlook items.
- CI fails if forbidden Outlook action calls are introduced.
- The browser host accepts messages only from Scribble's fixed extension
  identity and validates bounded native-message framing. Browser tools are
  limited to navigating and reading the user's own visible tab (http/https
  only, executed by the extension) and opening one unsent Outlook draft when
  the user asks for one (once per request, always unsent/unsaved). Page
  content and tool results are
  explicitly untrusted reference data.

These controls let model output select read-only context and, after explicit
local authorization, create one unsent draft. They prevent it from reaching an
email-send or source-mailbox mutation capability. They do not claim protection
against a compromised Windows account, modified add-in binary, vulnerabilities
in Outlook or .NET, or an administrator replacing installed files.

See [SECURITY.md](SECURITY.md) for the full threat model.

## Data flow

Every chat request initially sends the configured endpoint:

- selected email metadata, or metadata for the working-set emails (up to the
  configured per-request size);
- up to 12 recent chat turns;
- the current prompt;
- up to three explicitly added bounded text files;
- the editable writing profile only when drafting is locally authorized and the
  profile is enabled.

A browser request instead sends the latest prompt, up to 12 recent chat
turns, and the active tab's title, address, selection, and readable text,
captured automatically at send time - plus, during a browsing request, the
bounded text of pages the model navigated to in that same visible tab. It never
sends browser history, cookies, other tabs, or pages read in the background.

The model may then request:

- one search with bounded result summaries (up to the configured per-request
  size) from the primary Inbox and Sent Items when no working set is locked;
- no more unique message bodies than the configured per-request size across
  the entire request;
- conversation messages only within that same request-wide body limit;
- at most four tool calls per round and four context-retrieval rounds.

`/search` is handled locally before an LLM request is created. It returns
metadata matches (up to the configured per-request size) and does not transmit
bodies. A later normal prompt
sends the working-set metadata and exposes only the body-read tool for those
exact handles.

When the latest user prompt explicitly asks to create or open a draft, local
intent rules expose `create_draft` for that request. Its bounded arguments may
contain a new-message subject,
recipients, CC recipients, and body, or a reply body plus the exact temporary
handle of a searched or selected source message. The tool can only save and
display one unsent Outlook draft. While that draft is
linked, a recognized revision request exposes `update_draft` instead. Each
update supplies the complete bounded body and optional exact phrases to bold.
The local formatter HTML-encodes all text and inserts only fixed headings,
subheadings, lists, dividers, paragraphs, and `<strong>` markup. The model can
request these visual structures with a small text layout syntax, but raw HTML is
rejected. If a model returns Markdown emphasis markers, the shared local formatter
removes them and applies real bold formatting in both Scribble and Outlook. Stray
formatting asterisks are removed. Arbitrary model HTML is never accepted. Neither
tool has a send operation.

Email bodies are sent only when the model requests them through an approved
read-only tool. The add-in does not index, upload, or transmit the entire
mailbox automatically.

The optional Settings check sends the API key to `GET /v1/models` when that route
is available, then submits a lightweight synthetic chat request that contains no
mailbox data. **Connect & load models** loads the dropdown without running that probe.
A successful check proves that authentication, the entered model, and the
tool-call response shape work before the first real mailbox question.

Nothing is sent to Microsoft 365 by the add-in. Outlook itself continues to use
whatever mail server your organization configured.

## API compatibility

The endpoint must support:

```http
POST /v1/chat/completions
Authorization: Bearer YOUR_KEY
Content-Type: application/json
```

The request uses `model`, `messages`, `stream: false`, and standard
OpenAI-compatible function tools. The endpoint and selected model must support
chat-completions tool calling. The final response must provide
`choices[0].message.content` as text. `GET /v1/models` is optional. When
available, **Connect & load models** or **Test selected model** populates the editable model
list in Settings.

Direct Gemini model IDs and Google browser sign-in are unavailable in this
build. The dormant translation code remains for a possible future managed
release, but its OAuth, discovery, streaming, and non-streaming entry points
all fail closed while the product gate is off.

If a request fails, the sidebar shows diagnostic identifiers such as:

```text
HTTP_401_UNAUTHORIZED
HTTP_400_BAD_REQUEST
NETWORK_CONNECT_FAILURE
NETWORK_NAME_RESOLUTION
TLS_SECURE_CHANNEL_FAILURE
AI_TIMEOUT
RESPONSE_INVALID_JSON
RESPONSE_MISSING_CONTENT
TOOL_ROUND_LIMIT
DRAFT_PERMISSION_NOT_AVAILABLE
DRAFT_UPDATE_NOT_AVAILABLE
DRAFT_ALREADY_LINKED
DRAFT_TOOL_MUST_BE_EXCLUSIVE
DRAFT_CREATION_FAILED
DRAFT_UPDATE_FAILED
TONE_SAMPLES_INSUFFICIENT
TONE_ANALYSIS_FAILED
EXTERNAL_CONTEXT_FAILED
OUTLOOK_COM_0x800...
```

For HTTP failures it also shows the provider error message/code, request ID when
present, and a bounded response excerpt. The local diagnostic log records the
operation, exception type, diagnostic code, and HRESULT without email content,
prompts, endpoint responses, or API keys.

For connection failures, the sidebar also shows the target host and port,
exception chain, HRESULT, `WebExceptionStatus`, and Windows socket/native error
when available. This distinguishes DNS, connection refusal, proxy, TLS, and
timeout failures before the endpoint returns an HTTP response.

## Remove

1. In Chrome, open `chrome://extensions` and select **Remove** on the
   Scribble card.
2. Close Outlook, Excel, PowerPoint, and Word.
3. Open Windows **Installed apps** or **Apps & features**.
4. Uninstall **Scribble**.

Windows uninstallation removes the private native bridge, its registration, and
the staged extension files. It deliberately does not edit browser profiles, so
if you skip step 1 the browser can retain a broken unpacked-extension card until
you remove it there.

Endpoint settings remain under:

```text
%LOCALAPPDATA%\Scribble
```

Delete that folder manually if you also want to remove the encrypted API key and
local diagnostic log.

## Build

Requirements:

- Windows 10 or newer
- Visual Studio 2022 Build Tools with .NET Framework 4.8 targeting pack
- Inno Setup 6

Build the assembly and tests:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Restore-StrongNameKey.ps1
msbuild Scribble.sln /m /p:Configuration=Release
tests\GuardrailTests\bin\Release\GuardrailTests.exe
powershell -ExecutionPolicy Bypass -File scripts\Test-Guardrails.ps1
```

The solution builds the Office assembly, guardrail tests, and the .NET Framework
browser bridge. The browser extension is plain Manifest V3 HTML, CSS, and
JavaScript, so it needs no npm build. The repository stores the stable
strong-name key as Base64 so local and CI builds use the same COM identity. A
strong name is an assembly identity mechanism, not a trusted publisher
signature.

Build the installer:

```powershell
& "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe" `
  "installer\Scribble.iss"
```

The installer is written to:

```text
artifacts\ScribbleSetup.exe
```

GitHub Actions builds, smoke-tests, and publishes the same single-file installer.
Every push to `main` updates the **Latest** GitHub release so the install link
above always points at the newest build.

### Code signing (optional)

CI signs the add-in DLL, browser bridge, and installer when the repository has the
`SIGNING_PFX` (base64 PFX) and `SIGNING_PFX_PASSWORD` Actions secrets. To set
that up once, run on any Windows machine:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\New-SigningCert.ps1
```

It creates a personal self-signed code-signing certificate, prints the two
secret values to add, and leaves you a `scribble-signing.cer` to trust on each
machine that runs Scribble:

```powershell
certutil -user -addstore Root scribble-signing.cer
certutil -user -addstore TrustedPublisher scribble-signing.cer
```

Without the secrets the signing steps are skipped and the build stays
unsigned, exactly as before. A self-signed certificate does not earn
SmartScreen reputation like a paid one, but once trusted it makes every
installer and DLL verifiable as yours, and any tampered build fails
verification.

## Compatibility

- Classic Outlook, Excel, PowerPoint, and Word for Windows
- Google Chrome 116 or newer for the browser side panel
- Microsoft Office Professional Plus 2021
- 32-bit or 64-bit Office on Windows
- .NET Framework 4.8
- OpenAI-compatible endpoint and model with tool-calling support
- HTTP or HTTPS endpoint URL (HTTPS recommended outside the computer)

The new Outlook for Windows does not load COM add-ins.
