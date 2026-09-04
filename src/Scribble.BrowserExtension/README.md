# Scribble browser extension

This is a Manifest V3 extension for Google Chrome. It opens Scribble in the browser side panel and sends chat requests to the local native-messaging host named `com.scribble.browser`.

The extension is intentionally store-independent: it contains plain HTML, CSS, and JavaScript and can be installed with **Load unpacked**. No build step or npm installation is required.

## Before installing

Install Scribble browser support first. Scribble Setup must install and register the native bridge, and its native-host manifest must allow this extension origin:

```text
chrome-extension://olkepladbgkfkhlglooilnmalckpdada/
```

The fixed public key in `manifest.json` keeps that extension ID stable across computers and folder locations.

## Install in Google Chrome

1. Keep this entire folder in a permanent location. Moving or deleting it breaks a manually loaded extension.
2. Open `chrome://extensions` in Chrome.
3. Turn on **Developer mode**.
4. Choose **Load unpacked**.
5. Select this `Scribble.BrowserExtension` folder—the folder that directly contains `manifest.json`.
6. Pin Scribble from Chrome's Extensions menu if you want the toolbar button to remain visible.
7. Click the Scribble toolbar button to open the side panel.

After replacing extension files with a newer version, return to the browser's Extensions page and choose **Reload** on Scribble. Version 1.2 adds the required `debugger` permission; Chrome may disable an older installation until you review and accept the permission increase.

The panel footer compares the running extension version with the extension
bundled by the installed Scribble suite. It says **latest** when they match. If
the files have been updated but Chrome still has older code loaded, it shows
the available version and a **Reload extension update** button. A newly added
permission may still require approval on `chrome://extensions`.

## Using page context

The current tab's title, address, selection (up to 16,000 characters), and
readable text (up to 48,000 characters) are captured automatically when you
press **Send** - there is nothing to attach. The panel header shows which tab
is being shared; protected pages (browser settings, extension galleries, some
viewers) are shared as address-only.

Scribble can also browse for you. When you ask it to look something up, the
model opens http/https pages in up to five of its own background work tabs
(numbered 1-5, so it can compare sites side by side), searches through Google's
visible UI (including localized Google ccTLDs and `name=q` fields), inspects
ref-scoped controls, and performs bounded trusted input. Each mutation expires
the old refs, waits for a stable page, and returns a fresh snapshot with an
observed outcome instead of assuming that a click worked.
Version 1.5.1 fixes a pre-action rescan that could invalidate a freshly returned
Google result ref immediately before Scribble tried to click it. Version 1.5.2
also makes the model-facing contract explicit that a user-supplied bare domain,
such as `samsungtradein.ae`, can be opened directly with HTTPS.
Your current tab is never navigated away, and **Clear chat** closes Scribble's
work tabs. Public search and travel criteria can be typed, selected, clicked,
scrolled, and filtered. Typed values are capped at 200 characters and may come
from your request, a locally validated public alias such as Dubai to DXB, or a
clarification answer. Scribble asks you to approve the exact text before using
an unfamiliar inferred public term. Values appear verbatim in
one plain-language status beside the Pixel Pal. Raw control refs stay in the
internal tool transcript rather than appearing as duplicate chat cards. A
native operation-aware policy permits reversible product cards and HTTP(S)
product links while refusing sensitive typing, credential/password forms,
purchase/booking submits, payment, identity, message, upload/download, and
destructive actions. When several related details are ambiguous, it asks one
to three questions together with clickable options and waits. When
you ask Scribble to email someone, it opens one unsent Outlook draft window
for your review; ask for Excel and it opens one unsaved workbook. It can never
send or save either.

Long tasks continue while the page is making observable progress. Scribble
stops after 20 consecutive browser calls with no meaningful state change, with
a separate 120-round emergency cost/safety fuse. Public travel inputs expose
safe displayed values so the model can verify airport codes and dates before
submitting; sensitive field values remain unreadable.

Completed price, valuation, availability, and configured-product answers need
a structured evidence record validated by the extension against the final DOM,
work tab, revision, and action receipts. The evidence card can reopen that same
tab until **Clear chat**. Scribble-authored browser messages use first person;
page text, evidence quotations, error codes, and protocol markers stay verbatim.

Chrome displays its normal “Scribble is debugging this browser” banner during
trusted input. Scribble attaches only for an atomic action and detaches in a
`finally` block, so the banner may flicker. Clicking the banner's **Cancel**
control stops the run and requires user attention; Scribble never bypasses a
CAPTCHA, bot check, protected page, cross-origin widget, or sign-in wall.

The **Settings** button in the panel opens the shared Scribble Settings window
on your desktop - the same one the Office add-ins use. The extension treats
page and model text as untrusted plain text and never renders either as HTML.

MCP tools are unavailable here by default. To use a read-only web-search tool,
open Scribble Settings in an Office app, enter that tool's exact MCP name in the
server's **Chrome tool allowlist**, and tick the read-only approval. Browser
chat uses at most one approved server and one tool call per request. Never
approve an MCP tool that writes data or takes actions.

## Permissions

- `tabs` + `http://*/*`, `https://*/*` host permissions: read the active tab's
  bounded context and create/manage at most five inactive Scribble work tabs.
- `activeTab`, `scripting`: read the page text of the tab you are on.
- `debugger`: dispatch only allowlisted mouse, keyboard, and text input in a
  registered Scribble work tab after native policy authorization. It is never
  used on the active context tab and is detached after each atomic action.
- `contextMenus`: provides “Ask Scribble about this page.”
- `nativeMessaging`: talks to the locally installed `com.scribble.browser` bridge.
- `sidePanel`: hosts the Scribble conversation beside the current page.

There are no content scripts, no remote scripts, and no background page collection.

## Troubleshooting

- **Browser support is not installed:** rerun Scribble Setup with browser support enabled, then restart the browser.
- **The extension is not authorized:** the extension and Scribble bridge are from different builds, or the native-host manifest does not contain the stable origin above. Reinstall matching versions.
- **The footer says an extension update is available:** choose **Reload extension update**. If Chrome asks about new permissions, approve or re-enable Scribble on `chrome://extensions`.
- **No model is configured:** open Scribble Settings from an Office app and select/configure a model.
- **The page cannot be read:** browser settings pages, extension galleries, PDF viewers, and some protected pages block injected scripts; those tabs are shared as address-only. Try a regular webpage.

For diagnostics, open the extension entry on `chrome://extensions` and inspect the service worker or side-panel developer tools.
