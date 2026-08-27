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

Opening the side panel does not read the page. Use one of these explicit actions:

- **Attach selection** reads up to 16,000 characters currently selected in the active page.
- **Attach page** reads up to 48,000 characters of visible page text.
- **Attach visible screenshot** captures only the currently visible browser viewport, compressed to no more than 5 MB.
- **Clear context** removes all attached page data from the side panel.
- The page and selection context menus attach the corresponding text and open the side panel.

Attached context is included with chat requests only after the user presses **Send**, and remains attached until **Clear context** is selected. The extension treats page and model text as untrusted plain text and never renders either as HTML.

MCP tools are unavailable here by default. To use a read-only web-search tool,
open Scribble Settings in an Office app, enter that tool's exact MCP name in the
server's **Edge/Chrome tool allowlist**, and tick the read-only approval. Browser
chat uses at most one approved server and one tool call per request. Never
approve an MCP tool that writes data or takes actions.

## Permissions

- `activeTab`: temporary access to the tab where the user invoked Scribble; there is no permanent access to every website.
- `scripting`: reads selection or page text only after an attach button or Scribble context-menu action.
- `contextMenus`: provides “Ask Scribble about this page” and “Ask Scribble about this selection.”
- `nativeMessaging`: talks to the locally installed `com.scribble.browser` bridge.
- `sidePanel`: hosts the Scribble conversation beside the current page.

There are no host permissions, no `<all_urls>` access, no content scripts, no remote scripts, and no background page collection.

## Troubleshooting

- **Browser support is not installed:** rerun Scribble Setup with browser support enabled, then restart the browser.
- **The extension is not authorized:** the extension and Scribble bridge are from different builds, or the native-host manifest does not contain the stable origin above. Reinstall matching versions.
- **No model is configured:** open Scribble Settings from an Office app and select/configure a model.
- **The page cannot be read:** browser settings pages, extension galleries, PDF viewers, and some protected pages block injected scripts. Try a regular webpage or attach a visible screenshot.
- **A new tab cannot be read while the side panel stays open:** click the Scribble toolbar button on that tab to grant temporary `activeTab` access, then attach it.

For diagnostics, open the extension entry on `edge://extensions` or `chrome://extensions` and inspect the service worker or side-panel developer tools.
