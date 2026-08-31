"use strict";

const MENU_OPEN = "scribble-open-panel";

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
    console.warn("Scribble could not open the side panel.", error);
  });
});

chrome.contextMenus.onClicked.addListener((info, tab) => {
  if (info.menuItemId !== MENU_OPEN || !Number.isInteger(tab?.windowId)) {
    return;
  }

  // The panel captures the current tab's context automatically.
  void chrome.sidePanel.open({ windowId: tab.windowId }).catch((error) => {
    console.warn("Scribble could not open the side panel.", error);
  });
});

async function rebuildContextMenus() {
  try {
    await chrome.contextMenus.removeAll();
    chrome.contextMenus.create({
      id: MENU_OPEN,
      title: "Ask Scribble about this page",
      contexts: ["page", "selection"]
    });
  } catch (error) {
    console.warn("Scribble could not create its context menus.", error);
  }
}
