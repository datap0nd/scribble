"use strict";

const MENU_OPEN = "scribble-open-panel";
const OPERATOR_PORT = "scribble-browser-operator";
const CDP_VERSION = "1.3";
const ALLOWED_CDP_COMMANDS = new Set([
  "Input.dispatchMouseEvent",
  "Input.dispatchKeyEvent",
  "Input.insertText"
]);
const ALLOWED_KEYS = new Set([
  "Enter", "Escape", "Tab", "Backspace", "Delete", " ",
  "ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight",
  "Home", "End", "PageUp", "PageDown", "a"
]);
const operatorStates = new Map();
const attachedTabs = new Map();
const intentionalDetaches = new Set();

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

chrome.runtime.onConnect.addListener((port) => {
  const sidePanelUrl = chrome.runtime.getURL("sidepanel.html");
  if (port.name !== OPERATOR_PORT ||
      typeof port.sender?.url !== "string" ||
      !port.sender.url.startsWith(sidePanelUrl)) {
    port.disconnect();
    return;
  }

  const state = { chatId: "", workTabs: new Set() };
  operatorStates.set(port, state);
  port.onMessage.addListener((message) => {
    void handleOperatorMessage(port, state, message);
  });
  port.onDisconnect.addListener(() => {
    operatorStates.delete(port);
    void detachForPort(port);
  });
});

chrome.debugger.onDetach.addListener((source, reason) => {
  const tabId = source?.tabId;
  if (!Number.isInteger(tabId)) {
    return;
  }
  const owner = attachedTabs.get(tabId);
  attachedTabs.delete(tabId);
  if (intentionalDetaches.delete(tabId)) {
    return;
  }
  if (owner) {
    postOperator(owner, {
      type: "operatorDetached",
      tabId,
      reason: reason === "canceled_by_user"
        ? "canceled_by_user"
        : String(reason || "unexpected")
    });
  }
});

async function handleOperatorMessage(port, state, message) {
  if (!message || typeof message !== "object") {
    return;
  }

  if (message.type === "registerWorkTabs") {
    state.chatId = boundWorkerText(message.chatId, 100);
    state.workTabs = new Set(
      (Array.isArray(message.tabIds) ? message.tabIds : [])
        .filter((tabId) => Number.isInteger(tabId) && tabId >= 0)
        .slice(0, 5)
    );
    postOperator(port, {
      type: "workTabsRegistered",
      requestId: boundWorkerText(message.requestId, 100)
    });
    return;
  }

  if (message.type === "detachAll") {
    await detachForPort(port);
    postOperator(port, {
      type: "detachedAll",
      requestId: boundWorkerText(message.requestId, 100)
    });
    return;
  }

  if (message.type !== "cdpAction") {
    return;
  }

  const requestId = boundWorkerText(message.requestId, 100);
  const tabId = Number(message.tabId);
  try {
    if (!Number.isInteger(tabId) || !state.workTabs.has(tabId)) {
      throw new Error("CDP actions are allowed only in registered Scribble work tabs.");
    }
    const tab = await chrome.tabs.get(tabId);
    if (!/^https?:/i.test(String(tab.url || ""))) {
      throw new Error("CDP actions are restricted to HTTP and HTTPS work tabs.");
    }
    const requestedCommands = Array.isArray(message.commands)
      ? message.commands
      : [{ command: message.command, params: message.params }];
    if (requestedCommands.length === 0 || requestedCommands.length > 50) {
      throw new Error("A CDP action must include 1-50 commands.");
    }
    const commands = requestedCommands.map((requested) => {
      const command = boundWorkerText(requested?.command, 80);
      if (!ALLOWED_CDP_COMMANDS.has(command)) {
        throw new Error("The requested CDP command is not allowlisted.");
      }
      return { command, params: validateCdpParams(command, requested?.params) };
    });
    intentionalDetaches.delete(tabId);
    attachedTabs.set(tabId, port);
    await chrome.debugger.attach({ tabId }, CDP_VERSION);
    const results = [];
    try {
      for (const item of commands) {
        results.push(await chrome.debugger.sendCommand(
          { tabId },
          item.command,
          item.params
        ) || {});
      }
    } finally {
      intentionalDetaches.add(tabId);
      try {
        await chrome.debugger.detach({ tabId });
      } finally {
        attachedTabs.delete(tabId);
        setTimeout(() => intentionalDetaches.delete(tabId), 1_000);
      }
    }
    postOperator(port, { type: "cdpResult", requestId, results });
  } catch (error) {
    attachedTabs.delete(tabId);
    postOperator(port, {
      type: "cdpError",
      requestId,
      error: boundWorkerText(error instanceof Error ? error.message : String(error || ""), 500)
    });
  }
}

function validateCdpParams(command, raw) {
  const params = raw && typeof raw === "object" ? raw : {};
  if (command === "Input.insertText") {
    const text = typeof params.text === "string" ? params.text : "";
    if (!text || text.length > 200 || /[\u0000-\u001f\u007f]/.test(text)) {
      throw new Error("CDP text must contain 1-200 characters.");
    }
    return { text };
  }

  if (command === "Input.dispatchKeyEvent") {
    const type = String(params.type || "");
    const key = String(params.key || "");
    const modifiers = Number(params.modifiers || 0);
    if (!new Set(["keyDown", "keyUp", "rawKeyDown"]).has(type) ||
        !ALLOWED_KEYS.has(key) ||
        !new Set([0, 2]).has(modifiers) ||
        (modifiers === 2 && key !== "a")) {
      throw new Error("The requested keyboard event is not allowlisted.");
    }
    return { type, key, modifiers };
  }

  const type = String(params.type || "");
  if (!new Set(["mouseMoved", "mousePressed", "mouseReleased", "mouseWheel"]).has(type)) {
    throw new Error("The requested mouse event is not allowlisted.");
  }
  const x = boundedCoordinate(params.x);
  const y = boundedCoordinate(params.y);
  if (type === "mouseWheel") {
    return {
      type,
      x,
      y,
      deltaX: boundedDelta(params.deltaX),
      deltaY: boundedDelta(params.deltaY)
    };
  }
  if (type === "mouseMoved") {
    return { type, x, y, button: "none" };
  }
  return {
    type,
    x,
    y,
    button: "left",
    clickCount: 1
  };
}

function boundedCoordinate(value) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < -10_000 || number > 100_000) {
    throw new Error("The requested pointer coordinate is invalid.");
  }
  return Math.round(number);
}

function boundedDelta(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return 0;
  }
  return Math.max(-2_000, Math.min(2_000, Math.round(number)));
}

async function detachForPort(port) {
  const owned = [];
  for (const [tabId, owner] of attachedTabs) {
    if (owner === port) {
      owned.push(tabId);
    }
  }
  for (const tabId of owned) {
    intentionalDetaches.add(tabId);
    try {
      await chrome.debugger.detach({ tabId });
    } catch {
      intentionalDetaches.delete(tabId);
    } finally {
      attachedTabs.delete(tabId);
    }
  }
}

function postOperator(port, message) {
  try {
    port.postMessage(message);
  } catch {
    // The side panel closed; disconnect cleanup owns any attachment.
  }
}

function boundWorkerText(value, maximum) {
  return String(value || "").replace(/[\u0000-\u001f\u007f]/g, " ").slice(0, maximum);
}

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
