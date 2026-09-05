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
const KEY_DETAILS = Object.freeze({
  Enter: { code: "Enter", windowsVirtualKeyCode: 13 },
  Escape: { code: "Escape", windowsVirtualKeyCode: 27 },
  Tab: { code: "Tab", windowsVirtualKeyCode: 9 },
  Backspace: { code: "Backspace", windowsVirtualKeyCode: 8 },
  Delete: { code: "Delete", windowsVirtualKeyCode: 46 },
  " ": { code: "Space", windowsVirtualKeyCode: 32 },
  ArrowUp: { code: "ArrowUp", windowsVirtualKeyCode: 38 },
  ArrowDown: { code: "ArrowDown", windowsVirtualKeyCode: 40 },
  ArrowLeft: { code: "ArrowLeft", windowsVirtualKeyCode: 37 },
  ArrowRight: { code: "ArrowRight", windowsVirtualKeyCode: 39 },
  Home: { code: "Home", windowsVirtualKeyCode: 36 },
  End: { code: "End", windowsVirtualKeyCode: 35 },
  PageUp: { code: "PageUp", windowsVirtualKeyCode: 33 },
  PageDown: { code: "PageDown", windowsVirtualKeyCode: 34 },
  a: { code: "KeyA", windowsVirtualKeyCode: 65 }
});
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

  if (message.type === "inspectSemantics") {
    const requestId = boundWorkerText(message.requestId, 100), tabId = Number(message.tabId);
    try {
      if (!Number.isInteger(tabId) || !state.workTabs.has(tabId)) throw new Error("Perception requires a registered work tab.");
      const tab = await chrome.tabs.get(tabId);
      if (!/^https?:/i.test(String(tab.url || ""))) throw new Error("This browser page does not permit inspection.");
      attachedTabs.set(tabId, port);
      await chrome.debugger.attach({ tabId }, CDP_VERSION);
      try {
        const dom = await chrome.debugger.sendCommand({ tabId }, "DOMSnapshot.captureSnapshot", {
          computedStyles: ["display", "visibility", "opacity"], includeDOMRects: true
        });
        const ax = await chrome.debugger.sendCommand({ tabId }, "Accessibility.getFullAXTree", {});
        const observation = collectPerception(dom, ax, message.query, message.offset);
        let screenshotDataUrl = "";
        if (message.captureImage === true) {
          const screenshot = await chrome.debugger.sendCommand({ tabId }, "Page.captureScreenshot", {
            format: "jpeg", quality: 65, captureBeyondViewport: false
          });
          if (screenshot.data?.length <= 4 * 1024 * 1024) screenshotDataUrl = "data:image/jpeg;base64," + screenshot.data;
        }
        postOperator(port, { type: "perceptionResult", requestId, observation, screenshotDataUrl });
      } finally {
        intentionalDetaches.add(tabId);
        await chrome.debugger.detach({ tabId });
        attachedTabs.delete(tabId);
        setTimeout(() => intentionalDetaches.delete(tabId), 1000);
      }
    } catch (error) {
      attachedTabs.delete(tabId);
      postOperator(port, { type: "cdpError", requestId, error: boundWorkerText(error?.message, 500) });
    }
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

// A fixed read-only protocol, separate from the trusted-input command allowlist.
// Never return DOM input values, scripts, cookies, headers or arbitrary protocol data.
function collectPerception(snapshot, accessibility, query = "", offset = 0) {
  const strings = snapshot.strings || [], output = [];
  const roles = new Set(["button", "link", "checkbox", "radio", "switch", "slider", "spinbutton", "combobox", "textbox", "searchbox", "tab", "treeitem", "menuitem", "menuitemcheckbox", "menuitemradio", "option", "gridcell"]);
  const axByNode = new Map((accessibility.nodes || []).filter(n => !n.ignored && n.backendDOMNodeId).map(n => [n.backendDOMNodeId, n]));
  const filter = String(query || "").toLowerCase().slice(0, 200);
  for (const [documentIndex, doc] of (snapshot.documents || []).entries()) {
    const nodes = doc.nodes || {}, layout = doc.layout || {};
    const clickable = new Set(nodes.isClickable?.index || []);
    const children = new Map(), renderedText = new Map();
    for (let index = 0; index < (nodes.parentIndex || []).length; index++) {
      const parent = nodes.parentIndex[index];
      if (!children.has(parent)) children.set(parent, []);
      children.get(parent).push(index);
    }
    for (let row = 0; row < (layout.nodeIndex || []).length; row++) {
      const index = layout.nodeIndex[row], styles = (layout.styles?.[row] || []).map(i => strings[i]);
      if (strings[nodes.nodeName?.[index]] === '#text' && styles[0] !== 'none' &&
          !['hidden','collapse'].includes(styles[1]) && styles[2] !== '0')
        renderedText.set(index, strings[layout.text?.[row]] || strings[nodes.nodeValue?.[index]] || '');
    }
    const textLabel = index => {
      const queue = [index], pieces = [];
      let visited = 0, size = 0;
      while (queue.length && visited++ < 128 && size < 220) {
        const node = queue.shift(), tag = strings[nodes.nodeName?.[node]];
        if (/^(INPUT|TEXTAREA|SELECT|SCRIPT|STYLE|NOSCRIPT)$/.test(tag)) continue;
        const value = renderedText.get(node);
        if (value) {pieces.push(value);size += value.length;}
        queue.push(...(children.get(node) || []).slice(0,128-visited));
      }
      return pieces.join(' ').replace(/\s+/g,' ').trim().slice(0,220);
    };
    for (let row = 0; row < (layout.nodeIndex || []).length; row++) {
      const index = layout.nodeIndex[row], backendId = nodes.backendNodeId?.[index], ax = axByNode.get(backendId);
      const role = ax?.role?.value || "", listener = clickable.has(index);
      if (!listener && !roles.has(role)) continue;
      const styles = (layout.styles?.[row] || []).map(i => strings[i]);
      if (styles[0] === "none" || ["hidden", "collapse"].includes(styles[1])) continue;
      const bounds = layout.bounds?.[row];
      if (!bounds || bounds[2] < 1 || bounds[3] < 1) continue;
      const attrs = nodes.attributes?.[index] || [], attributes = {};
      for (let i = 0; i < attrs.length; i += 2) attributes[strings[attrs[i]]] = strings[attrs[i + 1]];
      if (["hidden", "password"].includes(attributes.type)) continue;
      const name = String(ax?.name?.value || attributes["aria-label"] || attributes.title || attributes.alt || textLabel(index)).slice(0, 220);
      if (filter && !`${role} ${name}`.toLowerCase().includes(filter)) continue;
      output.push({ backendId, role, name, tag: strings[nodes.nodeName?.[index]], listener,
        topDocument: documentIndex === 0, frameId: strings[doc.frameId],
        x: bounds[0] - (doc.scrollOffsetX || 0), y: bounds[1] - (doc.scrollOffsetY || 0), width: bounds[2], height: bounds[3],
        states: (ax?.properties || []).filter(p => ["checked", "selected", "expanded", "disabled", "readonly", "level"].includes(p.name))
          .map(p => ({ name: p.name, value: p.value?.value })) });
    }
  }
  const start = Math.max(0, Number(offset) || 0);
  return { methods: ["DOMSnapshot", "Accessibility"], controls: output.slice(start, start + 80),
    total: output.length, nextOffset: start + 80 < output.length ? start + 80 : null };
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
    const details = KEY_DETAILS[key];
    return {
      type,
      key,
      code: details.code,
      windowsVirtualKeyCode: details.windowsVirtualKeyCode,
      modifiers
    };
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
