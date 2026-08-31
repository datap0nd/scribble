---
name: Scribble
description: A restrained Outlook sidebar for chatting with bounded mailbox context and opening user-reviewed drafts.
colors:
  action-blue: "#005fb8"
  window: "Canvas"
  window-text: "CanvasText"
  control: "ButtonFace"
  control-text: "ButtonText"
  highlight: "Highlight"
  highlight-text: "HighlightText"
  secondary-text: "#505050"
  muted-surface: "#f4f6f8"
  border: "ButtonBorder"
  error: "#a32626"
  high-contrast-secondary: "GrayText"
  high-contrast-error: "LinkText"
typography:
  title:
    fontFamily: "system-ui, sans-serif"
    fontSize: "12pt"
    fontWeight: 700
    lineHeight: 1.2
  body:
    fontFamily: "system-ui, sans-serif"
    fontSize: "10pt"
    fontWeight: 400
    lineHeight: 1.4
  label:
    fontFamily: "system-ui, sans-serif"
    fontSize: "9pt"
    fontWeight: 700
    lineHeight: 1.3
  hint:
    fontFamily: "system-ui, sans-serif"
    fontSize: "8pt"
    fontWeight: 400
    lineHeight: 1.3
rounded:
  square: "0px"
spacing:
  tight: "4px"
  compact: "8px"
  control-gap: "10px"
  toolbar-x: "12px"
  content-x: "18px"
  dialog-x: "24px"
components:
  button-primary:
    backgroundColor: "{colors.action-blue}"
    textColor: "{colors.highlight-text}"
    typography: "{typography.label}"
    rounded: "{rounded.square}"
    padding: "0 16px"
    height: "34px"
  button-secondary:
    backgroundColor: "{colors.window}"
    textColor: "{colors.window-text}"
    typography: "{typography.body}"
    rounded: "{rounded.square}"
    padding: "0 14px"
    height: "34px"
  button-link:
    backgroundColor: "{colors.window}"
    textColor: "{colors.action-blue}"
    typography: "{typography.body}"
    rounded: "{rounded.square}"
    padding: "0 4px"
    height: "28px"
  input:
    backgroundColor: "{colors.window}"
    textColor: "{colors.window-text}"
    typography: "{typography.body}"
    rounded: "{rounded.square}"
    padding: "4px 6px"
    height: "34px"
---

# Design System: Scribble

## Overview

**Creative North Star: "The Guardrailed Desk Tool"**

Scribble is a compact native Outlook sidebar. It should feel native to classic Outlook, not like an AI showcase: familiar system typography, quiet white and cool-gray surfaces, square fields, restrained density, and blue reserved for direct actions.

The screen tells one ordered story: confirm mailbox scope and the optional selected message, ask a question, observe which bounded read-only context was loaded, then explicitly request one unsent draft in Outlook. The visual hierarchy must keep the local-intent and one-linked-draft boundary obvious. No control may imply that the utility can send mail.

**Key Characteristics:**

- Native Windows behavior and system settings take priority over decorative identity.
- One mailbox-scope strip anchors the sidebar and optional selected message.
- A large plain-text transcript carries the work without chat bubbles or HTML treatment.
- The composer and AI action are visually bounded together.
- Draft actions are separate, secondary, and enabled only when usable.
- Disclosure and status text state what data is used and what action occurred.

## Colors

The palette is mostly Windows system color roles, with a fixed Outlook-adjacent blue for direct actions in standard contrast and system highlight colors in high contrast.

### Primary

- **Action Blue:** Used for the Send to AI button, transcript speaker emphasis, and toolbar links. In high contrast, use the Windows `Highlight` system role instead.

### Neutral

- **Window:** The main transcript, toolbar, secondary buttons, and draft-action surface use the Windows canvas role.
- **Window Text:** Primary copy follows the Windows canvas text role.
- **Control:** Settings and high-contrast muted bands use the Windows control surface.
- **Control Text:** Settings text follows the Windows control text role.
- **Secondary Text:** Supporting metadata, hints, disclosures, and normal status messages use a quieter neutral in standard contrast and `GrayText` in high contrast.
- **Muted Surface:** Header, composer, and status bands use a cool gray in standard contrast, then fall back to the Windows control surface in high contrast.
- **Border:** Square field and secondary-button outlines use the Windows control-dark or button-border role.
- **Error:** Recoverable errors use restrained dark red in standard contrast and a system-provided high-contrast role when high contrast is active.

### Named Rules

**The System Color Rule.** System canvas, text, control, highlight, and border roles outrank fixed palette values whenever Windows high contrast is active.

**The One Accent Rule.** Blue marks direct actions and the user's transcript label. It is not decoration and must not spread into background panels or assistant content.

**The Text-Plus-State Rule.** Color never carries status alone. Every busy, error, disabled, disclosure, and draft state also has explicit text or native control state.

## Typography

**Display Font:** None

**Body Font:** Windows message-box system font, normally Segoe UI, with the active Windows fallback

**Label Font:** The same system family at bold weight

**Character:** Typography is familiar, compact, and subordinate to the task. It inherits Windows font settings and scales through WinForms `AutoScaleMode.Font`; nominal sizes describe the standard presentation, not a hard override of user settings.

### Hierarchy

- **Title** (bold, nominally 12pt): Scribble in the top strip.
- **Body** (regular, nominally 10pt): Transcript turns, composer text, settings fields, and primary reading content.
- **Label** (bold, nominally 9pt): Speaker names, field labels, and the primary action.
- **Hint** (regular, no smaller than 8pt): Keyboard guidance, draft disclosure, metadata, and status copy.

### Named Rules

**The System Font Rule.** Use `SystemFonts.MessageBoxFont` and relative size adjustments. Do not bundle a custom font or pin typography in a way that defeats Windows text scaling.

**The Safe-Rich-Text Rule.** Transcript content remains bounded plain text. A local parser may apply native RichTextBox bold to bounded character ranges after removing Markdown emphasis markers. Drafts may use a fixed local renderer for headings, subheadings, lists, dividers, and bold text. Model-provided HTML, links, images, scripts, and executable affordances are never rendered.

## Layout

The chat is a single-column, vertically stacked Outlook Custom Task Pane, initially 380 pixels wide with a 300-pixel minimum usable width. The top mailbox-scope strip is 92 pixels high and includes either the selected subject or ten-email working-set count plus the exact active model, followed by a compact 38-pixel toolbar. When active, a collapsible 322-pixel context layer appears above the transcript. It shows bounded email cards and external text-file cards in one scrollable ledger. The transcript consumes all remaining flexible height. The 154-pixel composer band and 64-pixel status band remain anchored at the bottom.

Horizontal content padding is 14 pixels in the sidebar work areas. The toolbar uses 8 pixels, while the modal settings form uses 24 pixels. Vertical rhythm is compact, generally 3 to 10 pixels between related controls. The transcript stays visually open and scrolls vertically instead of becoming a stack of cards.

The composer is a two-column grid: a fluid multiline text field and a fixed action column. The send button fills the message row. A persistent safety line below the message field says "Say 'create a draft' to open one. Scribble cannot send." After creation it becomes a blue linked state: "One draft linked. Revision requests update this draft only." Long message subjects and status text ellipsize rather than breaking the frame.

The settings window is a centered modal with Connection, MCP, Writing style, and Support tabs. Connection leads with endpoint URL and API key, then an explicit model-discovery action, editable model selector, remote-HTTP warning, and optional compatibility test. Direct Gemini and user-adjustable Limits tabs are absent. Writing style contains an explicit consent action, a visible disclosure that no analysis runs automatically, an editable profile field, and a draft-only enable checkbox. Save and Cancel remain common bottom actions.

The context toolbar uses short native actions for Add email, Add files, New, and Settings. Email drops resolve the current Outlook Explorer selection. File drops and the file picker accept only a small supported text-format set, show each accepted file in the context ledger, and expose the fixed three-file boundary in status text.

### Named Rules

**The Ordered Boundary Rule.** Keep mailbox scope, conversation and context ledger, composer, draft controls, and status in that order. This sequence is the interface's security explanation.

**The Transcript-Breathes Rule.** Fixed utility bands yield height to the transcript. Never shrink the conversation into a decorative card or make drafting controls compete with it.

## Elevation & Depth

The application defines no shadows, gradients, blur, overlays, or custom elevation. Depth comes from native window chrome, alternating white and muted system surfaces, and one-pixel control borders. Any outer window shadow is owned by Windows and must not be replicated inside the client area.

### Named Rules

**The Flat-Utility Rule.** App content is flat by default. Use tonal bands and native borders to separate regions, not card shadows or floating panels.

## Shapes

Controls use square native geometry. Text fields have fixed single borders; flat buttons have either a single outline or no border for link-like toolbar actions. Surfaces are rectangular bands with no clipping, pills, avatars, speech bubbles, or decorative silhouettes.

### Named Rules

**The Native Rectangle Rule.** Keep controls square and familiar. Rounded chat bubbles, pill buttons, and oversized AI ornaments contradict the product's utility character.

## Components

### Mailbox Scope Strip

- **Character:** Quiet context anchor, not a card.
- **Structure:** Muted full-width band with a bold "Scribble" title, an ellipsized optional `Selected: subject` or `Working set: N of M emails` (M is the configured per-request size, default 10), and exactly `Model: model_name`.
- **State:** When no context is selected, the same region says mailbox search remains available. `/search person or topic` replaces the local email working set, while `/search clear` removes it.

### Toolbar Links

- **Shape:** Borderless, square, 28-pixel-high buttons.
- **Color:** White canvas with action-blue text in standard contrast.
- **State:** Refresh selection, New chat, and Settings disable while a request is active. Refresh accepts one selected email or a Ctrl+click selection of up to ten. Native focus and disabled rendering remain visible.

### Transcript

- **Character:** A spacious, read-only plain-text document.
- **Structure:** Borderless white surface with vertical scrolling and no automatic URL detection.
- **Turns:** Speaker names are bold. "You" uses the action color; "Assistant" uses primary text. Model emphasis markers are removed and represented as native bold spans. Context-loading entries are italic secondary text and endpoint errors use explicit diagnostic codes.
- **Accessibility:** Accessible name is "Scribble conversation"; the description identifies it as a plain-text mailbox conversation and context-loading ledger.

### Working-Set Layer

- **Structure:** A quiet bordered layer between the toolbar and transcript contains a bold count, a Show or Hide action, and up to five separate email blocks.
- **Email blocks:** Each block presents a blue ordinal, bold subject, sender, and received date. Long text ellipsizes. The blocks are informational, not clickable, and the layer scrolls vertically when needed.
- **State:** `/search` and Outlook multi-selection replace the complete card set and open the review layer. Search failure preserves the previous cards. Sending a normal AI prompt collapses the layer to its count header so the transcript regains space. `/search clear`, a single selected email, and New chat remove the layer.
- **Accessibility:** The layer and every email block expose names and descriptions. The Show or Hide action remains keyboard reachable and never relies on color alone.

### Composer

- **Style:** Multiline square field with a fixed single border, vertical scrolling, and bounded input length.
- **Authorization:** Deterministic local code recognizes explicit drafting or revision intent only from the latest user-written prompt. Model output and loaded email content cannot authorize mutation.
- **Instruction:** A persistent safety line explains how to request a draft and states that Scribble cannot send. A separate hint says Ctrl+Enter submits the chat prompt.
- **Focus:** Keep the native Windows focus indication. Do not replace it with color-only styling.
- **Busy State:** Disable the field while waiting, change "Send to AI" to "Cancel," and restore the user's prompt if the request fails, times out, or is discarded.

### Primary Button

- **Shape:** Square, flat, filled action control.
- **Primary:** Action-blue background, system highlight text, bold label.
- **High Contrast:** Replace the fixed fill with the system highlight role.
- **State:** "Send to AI" starts the bounded request. "Cancel" is the only alternate label and cancels the in-flight request.

### Linked Draft State

- **Creation:** Explicit user wording such as "create a draft" locally authorizes one model-requested unsent draft attempt for that request.
- **Visible result:** Outlook displays the created item immediately. The chat then links to that exact item.
- **Revision:** Later drafting feedback updates and redisplays the same item. No second draft button or manual copy action exists.
- **Formatting:** The model provides plain text and optional exact bold phrases. Shared local code removes emphasis notation, then encodes text and applies fixed safe bold markup.
- **Boundary:** No send, schedule, move, delete, mark, categorize, BCC, arbitrary HTML, or source-message modification action belongs in this component family.

### Status Band

- **Character:** A full-width operational ledger at the bottom of the sidebar.
- **Content:** States mailbox scope, context retrieval, waiting, cancellation, diagnostic code, configuration, and unsent-draft outcomes in plain language.
- **Accessibility:** Exposes a status-bar role. Errors use error color plus explicit recovery copy.

### Settings Fields and Actions

- **Fields:** Endpoint URL and API key come first, followed by an editable Model selector populated after connection. Inputs use bold labels and accessible descriptions; the API key uses the system password character.
- **Disclosure:** Explain that prompts, recent conversation, and model-requested bounded mailbox context go to the configured endpoint and that the key is encrypted for the current Windows user.
- **HTTP warning:** Do not require an opt-in checkbox. Accept HTTP by default, but show an adjacent alert whenever a non-loopback endpoint will receive the API key, prompts, and retrieved email context without transport encryption.
- **Actions:** **Connect & load models** is the primary setup action. **Test selected model** is secondary. Save is primary; Cancel is secondary. Enter activates Save and Escape activates Cancel.
- **Endpoint Check:** Show progress, allow cancellation, and report the actual diagnostic code. A successful state means authentication, the selected model, and one synthetic read-only tool call passed. The probe must not load or execute against mailbox data.
- **Errors:** Validation failures appear inline as an accessible alert without closing the modal.

## Do's and Don'ts

### Do:

- **Do** keep mailbox scope and optional selected-message or working-set identity visible before conversation content.
- **Do** state that Inbox and Sent Items are available only through bounded read tools.
- **Do** keep the automatic local intent boundary visible and replace its guidance with linked-draft state after creation.
- **Do** keep `Selected: subject` at the top and hide repeated `RE:`, `FW:`, and `FWD:` display prefixes.
- **Do** cap `/search`, Outlook multi-selection, and request-wide body loading at the working-set size the user configured in Settings > Limits (default ten emails).
- **Do** say "unsent draft" and "for your review" in successful draft status text.
- **Do** inherit Windows system fonts, focus behavior, text scaling, and high-contrast colors.
- **Do** preserve keyboard operation, including Ctrl+Enter to send, Enter to save settings, and Escape to cancel settings.
- **Do** restore user input after request failure, timeout, or cancellation.

### Don't:

- **Don't** add a Send Mail control, auto-send behavior, scheduling, source-message mutation, or language that implies any of those capabilities.
- **Don't** expose arbitrary model commands, HTML, clickable output, or rich interactive response cards.
- **Don't** hide which mailbox scope or selected-message reference is available.
- **Don't** rely on fixed colors when Windows high contrast is active.
- **Don't** replace native focus and disabled states with color-only signals.
- **Don't** use chat bubbles, avatars, glowing AI motifs, gradients, rounded cards, or ornamental motion.
