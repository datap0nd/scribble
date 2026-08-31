# Scribble browser extension

This is one Manifest V3 extension for both Microsoft Edge and Google Chrome. It opens Scribble in the browser side panel and sends chat requests to the local native-messaging host named `com.scribble.browser`.

The extension is intentionally store-independent: it contains plain HTML, CSS, and JavaScript and can be installed with **Load unpacked**. No build step or npm installation is required.

## Before installing

Install Scribble browser support first. Scribble Setup must install and register the native bridge, and its native-host manifest must allow this extension origin:

```text
chrome-extension://olkepladbgkfkhlglooilnmalckpdada/
```

The fixed public key in `manifest.json` keeps that extension ID stable across computers and folder locations. Edge and Chrome use the same ID.

## Install in Microsoft Edge

1. Keep this entire folder in a permanent location. Moving or deleting it breaks a manually loaded extension.
2. Open `edge://extensions` in Edge.
3. Turn on **Developer mode**.
4. Choose **Load unpacked**.
5. Select this `Scribble.BrowserExtension` folder—the folder that directly contains `manifest.json`.
6. Pin Scribble from Edge's Extensions menu if you want the toolbar button to remain visible.
7. Click the Scribble toolbar button to open the side panel.

## Install in Google Chrome

1. Keep this entire folder in a permanent location. Moving or deleting it breaks a manually loaded extension.
2. Open `chrome://extensions` in Chrome.
3. Turn on **Developer mode**.
4. Choose **Load unpacked**.
5. Select this `Scribble.BrowserExtension` folder—the folder that directly contains `manifest.json`.
6. Pin Scribble from Chrome's Extensions menu if you want the toolbar button to remain visible.
7. Click the Scribble toolbar button to open the side panel.

After replacing extension files with a newer version, return to the browser's Extensions page and choose **Reload** on Scribble.

## Using page context

The current tab's title, address, selection (up to 16,000 characters), and
readable text (up to 48,000 characters) are captured automatically when you
press **Send** - there is nothing to attach. The panel header shows which tab
is being shared; protected pages (browser settings, extension galleries, some
viewers) are shared as address-only.

Scribble can also browse for you. When you ask it to look something up, the
model may navigate this same visible tab to http/https pages and read them,
across up to eight bounded tool rounds per request. It cannot click page
controls, fill or submit forms, sign in, purchase, download, or upload. When
your own message asks for an email draft, Scribble can open one unsent Outlook
draft window for your review; it can never send it.

The **Settings** button in the panel opens the shared Scribble Settings window
on your desktop - the same one the Office add-ins use. The extension treats
page and model text as untrusted plain text and never renders either as HTML.

MCP tools are unavailable here by default. To use a read-only web-search tool,
open Scribble Settings in an Office app, enter that tool's exact MCP name in the
server's **Edge/Chrome tool allowlist**, and tick the read-only approval. Browser
chat uses at most one approved server and one tool call per request. Never
approve an MCP tool that writes data or takes actions.

## Permissions

- `tabs` + `http://*/*`, `https://*/*` host permissions: read the active tab's
  title/URL/text and navigate that same visible tab when the model browses for
  you. Other tabs and background pages are never read.
- `activeTab`, `scripting`: read the page text of the tab you are on.
- `contextMenus`: provides “Ask Scribble about this page.”
- `nativeMessaging`: talks to the locally installed `com.scribble.browser` bridge.
- `sidePanel`: hosts the Scribble conversation beside the current page.

There are no content scripts, no remote scripts, and no background page collection.

## Troubleshooting

- **Browser support is not installed:** rerun Scribble Setup with browser support enabled, then restart the browser.
- **The extension is not authorized:** the extension and Scribble bridge are from different builds, or the native-host manifest does not contain the stable origin above. Reinstall matching versions.
- **No model is configured:** open Scribble Settings from an Office app and select/configure a model.
- **The page cannot be read:** browser settings pages, extension galleries, PDF viewers, and some protected pages block injected scripts; those tabs are shared as address-only. Try a regular webpage.

For diagnostics, open the extension entry on `edge://extensions` or `chrome://extensions` and inspect the service worker or side-panel developer tools.
