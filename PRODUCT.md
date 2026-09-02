# Product

<!-- impeccable:product-schema 1 -->

## Platform

windows

## Stack

Delegated: C# on .NET Framework 4.8 as a classic Outlook COM add-in, with a
Windows Forms control hosted in an Outlook Custom Task Pane and a single-file
Windows installer. This targets
Microsoft Office Professional Plus 2021 on Windows without Microsoft 365 add-in
deployment.

## Users

The primary user works in classic Outlook on a managed Windows work PC and wants
to ask questions across their local mailbox, let the model retrieve only relevant
context, refine a response through conversation, and open an unsent draft for
final review.

## Product Purpose

The add-in provides a native chat sidebar inside Outlook. A local `/search`
command or Outlook multi-selection creates a reviewable working set bounded by
the per-request size the user sets in Settings > Limits (default ten
emails). A normal prompt sends recent in-memory conversation plus optional
selected-message or working-set metadata to a user-configured OpenAI-compatible
endpoint. The model may request bounded read-only searches and message bodies
from the primary Inbox and Sent Items, but it can never load more unique
bodies per request than that configured size. After the user explicitly arms one request, the model may
create one unsent Outlook reply or new-message draft. That visible Outlook item
then remains linked to the chat so later feedback updates the same draft.

Success means installation is understandable, configuration takes one endpoint,
model name, and API key, and no model response can invoke an Outlook send action.

## Positioning

Model output can invoke only a compile-time allowlist of bounded mailbox read
tools. `create_draft` appears only when local code recognizes explicit drafting
intent in the latest user-written prompt.
After creation it is replaced by `update_draft`, which can modify only the one
locally linked item. The dedicated host exposes no send operation.

## Operating Context

- Microsoft Office Professional Plus 2021 with classic Outlook on Windows.
- Per-user local installation is preferred.
- The user opens Scribble from the ribbon or right-clicks one to ten selected
  emails and chooses **Send to Scribble**, then works in a right-docked Outlook
  Custom Task Pane. One selected email receives a temporary read handle. Selecting
  two or more emails (up to the configured working-set size) creates the locked
  working set.
- Configuration is stored for the current Windows user. The API key is encrypted
  with Windows Data Protection API.
- Conversations are kept in memory and disappear when Outlook closes or the user
  starts a new chat.

## Capabilities and Constraints

- Search and read bounded context from the primary Inbox and Sent Items.
- Accept explicitly selected local files and Outlook attachments up to 100 MiB
  each, with a 250 MiB source budget per operation. Extraction stays local;
  only bounded text or downscaled image input reaches the configured endpoint.
- Launch saved prompts immediately from a two-part Skills shelf. Public skills
  are read-only packages shipped with Scribble; Local skills are created per
  Office app and remain in the current Windows user's LocalAppData.
- Ship an Outlook-only **Morning unread summary** Public skill. It starts a
  fresh in-memory chat, searches the primary Inbox from 5:00 PM yesterday
  through click time in the PC's local timezone, and summarizes no more than
  the configured email limit without changing read state.
- Configure up to twenty named local Topics, explicitly select one per chat,
  and search and read only bounded relevant excerpts from its recursively
  indexed document folder.
- Handle `/search person or topic` locally, retain only the newest ten metadata
  matches, show them as distinct collapsible context cards, and allow another
  `/search` to replace that set before an LLM call.
- Accept Ctrl+click multi-selection of two or more Outlook emails (up to the
  configured working-set size) as the same locked working set.
- Hold a text conversation about the mailbox, a selected message, or a retrieved
  conversation.
- Generate text suitable for a reply or a new message.
- Create and display at most one unsent Outlook draft per chat, then update that
  same item at most once per later user request.
- Bind reply drafts to the exact temporary handle returned for the searched or
  selected source message. Never fall back to the latest mailbox item.
- Never send, schedule, move, delete, mark, categorize, or modify the source email.
- Skills are manual prompt shortcuts, not scheduled or background jobs, and
  cannot expand the tool allowlist or draft authorization rules.
- Without a working set, expose only `search_mailbox`, `read_messages`, and
  `read_thread`, with one search and ten unique bodies per request. With a
  working set, expose only `read_messages` for those approved handles.
  Conditionally expose `create_draft` for a locally recognized drafting request
  or `update_draft` for a locally recognized revision of the one linked draft,
  never both.
- Reject all other model tool calls and cap calls, rounds, results, and returned
  text.
- Never render model output as HTML or execute it as code. A shared local
  formatter removes Markdown emphasis notation and maps bounded spans to native
  RichTextBox bold in chat or locally encoded `<strong>` in Outlook drafts.
  Draft formatting also accepts exact bold phrases. Stray formatting asterisks
  are removed before display.
- Support an OpenAI-compatible `/v1/chat/completions` endpoint.
- Recommend the Qwen3.8-27B family as the fast agentic default while preserving
  editable model identifiers and user choice when several other models are
  available.
- Verify authentication, optional `/v1/models` discovery, and actual read-only
  tool-call compatibility from Settings without loading mailbox data.
- Permit HTTP and HTTPS endpoint URLs without a separate opt-in control.
- Warn clearly for non-loopback HTTP that the API key and mailbox context will
  travel without transport encryption; recommend HTTPS outside the computer.
- Keep Topic indexes under the current user's LocalAppData, never modify source
  folders, never expose absolute paths to the model, and disclose that selected
  excerpts travel to the configured model endpoint.
- Target both 32-bit and 64-bit Office from one installer when practical.
- A production installer should be code-signed. Signing credentials are not
  included in the repository.

## Brand Commitments

The user-facing and name-bearing technical identity is Scribble. Upgrade
compatibility rests on the unchanged installer AppId, COM CLSIDs, strong-name
key, assembly version, project GUIDs, and browser extension key/ID. Assemblies,
ProgIDs, settings paths, installer artifacts, and repository names all use the
`Scribble` identity. The UI should feel like a restrained Windows productivity
utility, not an AI showcase. Language must be direct, calm, and explicit about
what data is read and when a draft is created.

## Evidence on Hand

The product brief is the user's requested workflow. There are no approved company
logos, claims, screenshots, or signing certificates, and future work must not
fabricate them.

## Product Principles

- Capabilities, not prompts, define the security boundary.
- Nothing leaves the active conversation and the locally or model-selected
  ten-email read boundary.
- Drafting always ends in Outlook's normal editor with the user in control.
- Local configuration should be inspectable, reversible, and per-user.
- Familiar Windows behavior is more important than decorative novelty.

## Accessibility & Inclusion

The chat sidebar must support keyboard-only operation, visible focus, system text
scaling, high-contrast-compatible colors, and plain-language error recovery.
