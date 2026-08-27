"use strict";

const MENU_PAGE = "ai365-ask-page";
const MENU_SELECTION = "ai365-ask-selection";
const MAX_SELECTION_CHARS = 16_000;
const MAX_PAGE_TEXT_CHARS = 48_000;
const MAX_TITLE_CHARS = 512;
const MAX_URL_CHARS = 4_096;

const pendingContextMessages = new Map();

chrome.runtime.onInstalled.addListener(() => {
  void rebuildContextMenus();
});

chrome.runtime.onStartup.addListener(() => {
  void rebuildContextMenus();
});

chrome.action.onClicked.addListener((tab) => {
  if (!Number.isInteger(tab.windowId)) {
    return;
  }

  // This call happens directly in the toolbar-click user gesture.
  void chrome.sidePanel.open({ windowId: tab.windowId }).catch((error) => {
    console.warn("AI365 could not open the side panel.", error);
  });
});

chrome.contextMenus.onClicked.addListener((info, tab) => {
  if (info.menuItemId !== MENU_PAGE && info.menuItemId !== MENU_SELECTION) {
    return;
  }

  void handleContextMenuClick(info, tab);
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (sender.id !== chrome.runtime.id ||
      message?.type !== "consumePendingBrowserContext" ||
      !Number.isInteger(message.windowId)) {
    return undefined;
  }

  const pending = pendingContextMessages.get(message.windowId) || null;
  pendingContextMessages.delete(message.windowId);
  sendResponse({ pending });
  return false;
});

async function rebuildContextMenus() {
  try {
    await chrome.contextMenus.removeAll();
    chrome.contextMenus.create({
      id: MENU_PAGE,
      title: "Ask AI365 about this page",
      contexts: ["page"]
    });
    chrome.contextMenus.create({
      id: MENU_SELECTION,
      title: "Ask AI365 about this selection",
      contexts: ["selection"]
    });
  } catch (error) {
    console.warn("AI365 could not create its context menus.", error);
  }
}

async function handleContextMenuClick(info, tab) {
  if (!Number.isInteger(tab?.id) || !Number.isInteger(tab?.windowId)) {
    return;
  }

  // Open immediately while Chrome/Edge still considers this a user gesture.
  const panelPromise = chrome.sidePanel.open({ windowId: tab.windowId });

  let panelMessage;
  try {
    if (info.menuItemId === MENU_SELECTION) {
      const selection = boundText(info.selectionText, MAX_SELECTION_CHARS);
      if (!selection.trim()) {
        throw new Error("The selection is empty. Select some page text and try again.");
      }

      panelMessage = {
        type: "applyBrowserContext",
        targetWindowId: tab.windowId,
        kind: "selection",
        context: {
          title: boundText(tab.title, MAX_TITLE_CHARS),
          url: boundText(info.pageUrl || tab.url, MAX_URL_CHARS),
          selection
        }
      };
    } else {
      const context = await readPage(tab.id);
      if (!context.pageText.trim()) {
        throw new Error("AI365 could not find readable text on this page.");
      }

      panelMessage = {
        type: "applyBrowserContext",
        targetWindowId: tab.windowId,
        kind: "page",
        context
      };
    }
  } catch (error) {
    panelMessage = {
      type: "browserContextError",
      targetWindowId: tab.windowId,
      error: describePageAccessError(error)
    };
  }

  try {
    await panelPromise;
  } catch (error) {
    console.warn("AI365 could not open the side panel.", error);
  }

  await deliverToSidePanel(panelMessage);
}

async function readPage(tabId) {
  const results = await chrome.scripting.executeScript({
    target: { tabId },
    func: (pageLimit, titleLimit, urlLimit) => {
      const root = document.body || document.documentElement;
      const pageText = root && typeof root.innerText === "string"
        ? root.innerText.slice(0, pageLimit)
        : "";

      return {
        title: String(document.title || "").slice(0, titleLimit),
        url: String(location.href || "").slice(0, urlLimit),
        pageText
      };
    },
    args: [MAX_PAGE_TEXT_CHARS, MAX_TITLE_CHARS, MAX_URL_CHARS]
  });

  const context = results?.[0]?.result;
  if (!context || typeof context !== "object") {
    throw new Error("The page did not return readable content.");
  }

  return {
    title: boundText(context.title, MAX_TITLE_CHARS),
    url: boundText(context.url, MAX_URL_CHARS),
    pageText: boundText(context.pageText, MAX_PAGE_TEXT_CHARS)
  };
}

async function deliverToSidePanel(message) {
  if (!Number.isInteger(message?.targetWindowId)) {
    return;
  }

  pendingContextMessages.set(message.targetWindowId, message);

  try {
    const response = await chrome.runtime.sendMessage(message);
    if (response?.accepted === true) {
      pendingContextMessages.delete(message.targetWindowId);
    }
  } catch {
    // A newly opened panel asks for pending context as soon as it loads.
  }
}

function boundText(value, limit) {
  return typeof value === "string" ? value.slice(0, limit) : "";
}

function describePageAccessError(error) {
  const message = error instanceof Error ? error.message : String(error || "");
  const lower = message.toLowerCase();

  if (
    lower.includes("cannot access contents") ||
    lower.includes("missing host permission") ||
    lower.includes("cannot be scripted")
  ) {
    return "This page does not allow extensions to read its text. Try a regular webpage, or attach a visible screenshot instead.";
  }

  return message || "AI365 could not read this page. Click the AI365 toolbar button on the tab and try again.";
}
