"use strict";

const NATIVE_HOST = "com.ai365.browser";
const MAX_SELECTION_CHARS = 16_000;
const MAX_PAGE_TEXT_CHARS = 48_000;
const MAX_HISTORY_TURNS = 12;
const MAX_HISTORY_CONTENT_CHARS = 48_000;
const MAX_PROMPT_CHARS = 16_000;
const MAX_SCREENSHOT_BYTES = 5 * 1024 * 1024;
const MAX_TITLE_CHARS = 512;
const MAX_URL_CHARS = 4_096;
const PING_TIMEOUT_MS = 10_000;
const CHAT_TIMEOUT_MS = 300_000;

const elements = {
  retryConnection: document.getElementById("retryConnection"),
  connectionDot: document.getElementById("connectionDot"),
  connectionLabel: document.getElementById("connectionLabel"),
  attachSelection: document.getElementById("attachSelection"),
  attachPage: document.getElementById("attachPage"),
  attachScreenshot: document.getElementById("attachScreenshot"),
  clearContext: document.getElementById("clearContext"),
  contextSource: document.getElementById("contextSource"),
  contextEmpty: document.getElementById("contextEmpty"),
  contextItems: document.getElementById("contextItems"),
  contextNotice: document.getElementById("contextNotice"),
  messages: document.getElementById("messages"),
  welcome: document.getElementById("welcome"),
  composer: document.getElementById("composer"),
  prompt: document.getElementById("prompt"),
  promptCount: document.getElementById("promptCount"),
  send: document.getElementById("send"),
  activity: document.getElementById("activity"),
  connectionDetails: document.getElementById("connectionDetails")
};

let pageContext = emptyContext();
let conversationHistory = [];
let isSending = false;
let isPinging = false;
let panelWindowId = null;
let connection = {
  connected: false,
  configured: false,
  model: "",
  supportsVision: false,
  version: ""
};

elements.retryConnection.addEventListener("click", () => {
  void pingNativeHost();
});

elements.attachSelection.addEventListener("click", () => {
  void attachSelection();
});

elements.attachPage.addEventListener("click", () => {
  void attachPage();
});

elements.attachScreenshot.addEventListener("click", () => {
  void attachVisibleScreenshot();
});

elements.clearContext.addEventListener("click", () => {
  pageContext = emptyContext();
  setContextNotice("Context cleared. Nothing from the page is attached.");
  renderContext();
});

elements.prompt.addEventListener("input", () => {
  if (elements.prompt.value.length > MAX_PROMPT_CHARS) {
    elements.prompt.value = elements.prompt.value.slice(0, MAX_PROMPT_CHARS);
  }
  renderComposerState();
});

elements.prompt.addEventListener("keydown", (event) => {
  if (event.key === "Enter" && !event.shiftKey && !event.isComposing) {
    event.preventDefault();
    elements.composer.requestSubmit();
  }
});

elements.composer.addEventListener("submit", (event) => {
  event.preventDefault();
  void sendChatMessage();
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (sender.id !== chrome.runtime.id ||
      !Number.isInteger(panelWindowId) ||
      message?.targetWindowId !== panelWindowId) {
    return undefined;
  }

  if (message?.type === "applyBrowserContext") {
    applyContextPatch(message.context, message.kind);
    sendResponse({ accepted: true });
    return false;
  }

  if (message?.type === "browserContextError") {
    setContextNotice(boundText(message.error, 1_000), true);
    sendResponse({ accepted: true });
    return false;
  }

  return undefined;
});

void initialize();

async function initialize() {
  renderContext();
  renderComposerState();

  try {
    const currentWindow = await chrome.windows.getCurrent();
    panelWindowId = Number.isInteger(currentWindow?.id)
      ? currentWindow.id
      : null;
    const response = await chrome.runtime.sendMessage({
      type: "consumePendingBrowserContext",
      windowId: panelWindowId
    });
    const pending = response?.pending;
    if (pending?.type === "applyBrowserContext") {
      applyContextPatch(pending.context, pending.kind);
    } else if (pending?.type === "browserContextError") {
      setContextNotice(boundText(pending.error, 1_000), true);
    }
  } catch {
    // There is no pending context on a normal toolbar open.
  }

  await pingNativeHost();
}

async function pingNativeHost() {
  if (isPinging) {
    return;
  }

  isPinging = true;
  setConnectionView("connecting", "Connecting");
  setActivity("Connecting to the AI365 browser bridge…");

  try {
    const response = await sendNativeMessage({ type: "ping" }, PING_TIMEOUT_MS);
    updateConnectionFromResponse(response);

    if (response?.ok !== true) {
      throw new NativeResponseError(describeHostResponseError(response), response?.errorCode);
    }

    if (response.configured !== true) {
      connection.configured = false;
      setConnectionView("warning", "Setup needed");
      setActivity("AI365 is connected, but no model is configured. Open AI365 Settings in an Office app and choose a model.");
    } else {
      connection.connected = true;
      connection.configured = true;
      setConnectionView("connected", "Connected");
      setActivity("Ready. Page content is shared only when you attach it and send a message.");
    }
  } catch (error) {
    connection.connected = false;
    connection.configured = false;
    setConnectionView("error", "Not connected");
    setActivity(describeNativeMessagingError(error));
  } finally {
    isPinging = false;
    renderConnectionDetails();
    renderComposerState();
  }
}

async function attachSelection() {
  await withContextButton(elements.attachSelection, "Attaching…", async () => {
    const tab = await getActiveTab();
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: (selectionLimit, titleLimit, urlLimit) => ({
        title: String(document.title || "").slice(0, titleLimit),
        url: String(location.href || "").slice(0, urlLimit),
        selection: String(window.getSelection?.().toString() || "").slice(0, selectionLimit)
      }),
      args: [MAX_SELECTION_CHARS, MAX_TITLE_CHARS, MAX_URL_CHARS]
    });

    const context = results?.[0]?.result;
    if (!context?.selection?.trim()) {
      throw new Error("No text is selected. Select text on the page, then try again.");
    }

    applyContextPatch(context, "selection");
  });
}

async function attachPage() {
  await withContextButton(elements.attachPage, "Attaching…", async () => {
    const tab = await getActiveTab();
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: (pageLimit, titleLimit, urlLimit) => {
        const root = document.body || document.documentElement;
        return {
          title: String(document.title || "").slice(0, titleLimit),
          url: String(location.href || "").slice(0, urlLimit),
          pageText: String(root?.innerText || "").slice(0, pageLimit)
        };
      },
      args: [MAX_PAGE_TEXT_CHARS, MAX_TITLE_CHARS, MAX_URL_CHARS]
    });

    const context = results?.[0]?.result;
    if (!context?.pageText?.trim()) {
      throw new Error("AI365 could not find readable text on this page.");
    }

    applyContextPatch(context, "page");
  });
}

async function attachVisibleScreenshot() {
  await withContextButton(elements.attachScreenshot, "Capturing…", async () => {
    const tab = await getActiveTab();
    let screenshotDataUrl = "";

    // Chromium permits only two captureVisibleTab calls per second. A second,
    // more aggressive JPEG pass keeps us within that API limit.
    for (const quality of [82, 40]) {
      screenshotDataUrl = await chrome.tabs.captureVisibleTab(tab.windowId, {
        format: "jpeg",
        quality
      });

      if (dataUrlByteLength(screenshotDataUrl) <= MAX_SCREENSHOT_BYTES) {
        break;
      }
    }

    const size = dataUrlByteLength(screenshotDataUrl);
    if (!isSupportedScreenshot(screenshotDataUrl) || size <= 0) {
      throw new Error("The browser did not return a usable screenshot.");
    }
    if (size > MAX_SCREENSHOT_BYTES) {
      throw new Error("The visible page is too large to attach as a 5 MB screenshot. Make the browser window smaller and try again.");
    }

    applyContextPatch({
      title: boundText(tab.title, MAX_TITLE_CHARS),
      url: boundText(tab.url, MAX_URL_CHARS),
      screenshotDataUrl
    }, "screenshot");

    if (connection.connected && connection.supportsVision !== true) {
      setContextNotice("Screenshot attached. The currently selected model may not support vision; choose a vision model in AI365 Settings if the request fails.");
    }
  });
}

async function withContextButton(button, busyLabel, operation) {
  const originalLabel = button.textContent;
  button.disabled = true;
  button.textContent = busyLabel;
  setContextNotice("");

  try {
    await operation();
  } catch (error) {
    setContextNotice(describePageAccessError(error), true);
  } finally {
    button.textContent = originalLabel;
    button.disabled = false;
  }
}

async function getActiveTab() {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  const tab = tabs?.[0];

  if (!Number.isInteger(tab?.id) || !Number.isInteger(tab?.windowId)) {
    throw new Error("No active webpage was found.");
  }

  return tab;
}

function applyContextPatch(rawPatch, kind) {
  if (!rawPatch || typeof rawPatch !== "object") {
    setContextNotice("AI365 received invalid page context.", true);
    return;
  }

  const patch = {
    title: boundText(rawPatch.title, MAX_TITLE_CHARS),
    url: boundText(rawPatch.url, MAX_URL_CHARS),
    selection: boundText(rawPatch.selection, MAX_SELECTION_CHARS),
    pageText: boundText(rawPatch.pageText, MAX_PAGE_TEXT_CHARS),
    screenshotDataUrl: isSupportedScreenshot(rawPatch.screenshotDataUrl)
      && dataUrlByteLength(rawPatch.screenshotDataUrl) <= MAX_SCREENSHOT_BYTES
      ? rawPatch.screenshotDataUrl
      : ""
  };

  if (patch.url && pageContext.url && patch.url !== pageContext.url) {
    pageContext = emptyContext();
  }

  pageContext.title = patch.title || pageContext.title;
  pageContext.url = patch.url || pageContext.url;

  if (kind === "selection" && patch.selection) {
    pageContext.selection = patch.selection;
  } else if (kind === "page" && patch.pageText) {
    pageContext.pageText = patch.pageText;
  } else if (kind === "screenshot" && patch.screenshotDataUrl) {
    pageContext.screenshotDataUrl = patch.screenshotDataUrl;
  }

  setContextNotice(contextAttachedMessage(kind));
  renderContext();
}

function renderContext() {
  const hasSelection = pageContext.selection.length > 0;
  const hasPage = pageContext.pageText.length > 0;
  const hasScreenshot = pageContext.screenshotDataUrl.length > 0;
  const hasContext = hasSelection || hasPage || hasScreenshot;

  elements.contextSource.textContent = pageContext.title || pageContext.url || "Nothing attached";
  elements.contextSource.title = pageContext.url || pageContext.title || "";
  elements.contextEmpty.hidden = hasContext;
  elements.clearContext.disabled = !hasContext;

  elements.contextItems.replaceChildren();
  if (hasSelection) {
    appendContextItem(`Selection · ${formatNumber(pageContext.selection.length)} characters`);
  }
  if (hasPage) {
    appendContextItem(`Page text · ${formatNumber(pageContext.pageText.length)} characters`);
  }
  if (hasScreenshot) {
    appendContextItem(`Visible screenshot · ${formatBytes(dataUrlByteLength(pageContext.screenshotDataUrl))}`);
  }
}

function appendContextItem(label) {
  const item = document.createElement("li");
  item.className = "context-item";
  item.textContent = label;
  elements.contextItems.append(item);
}

function setContextNotice(message, isError = false) {
  elements.contextNotice.textContent = message || "";
  elements.contextNotice.classList.toggle("error", isError);
}

async function sendChatMessage() {
  const prompt = boundText(elements.prompt.value, MAX_PROMPT_CHARS).trim();
  if (!prompt || isSending) {
    return;
  }

  if (!connection.connected || !connection.configured) {
    setActivity("AI365 is not ready. Retry the connection, or configure a model in AI365 Settings.");
    return;
  }

  let context;
  try {
    context = validatedContextSnapshot();
  } catch (error) {
    setContextNotice(error instanceof Error ? error.message : String(error), true);
    return;
  }

  const request = {
    type: "chat",
    requestId: createRequestId(),
    prompt,
    history: conversationHistory.slice(-MAX_HISTORY_TURNS).map((turn) => ({
      role: turn.role === "assistant" ? "assistant" : "user",
      content: boundText(turn.content, MAX_HISTORY_CONTENT_CHARS)
    })),
    context
  };

  appendMessage("user", prompt);
  elements.prompt.value = "";
  isSending = true;
  renderComposerState();
  setActivity("AI365 is thinking…");

  try {
    const response = await sendNativeMessage(request, CHAT_TIMEOUT_MS);
    updateConnectionFromResponse(response);

    if (response?.ok !== true) {
      throw new NativeResponseError(describeHostResponseError(response), response?.errorCode);
    }

    const content = typeof response.content === "string" ? response.content.trim() : "";
    if (!content) {
      throw new NativeResponseError("AI365 returned an empty response.", "EMPTY_RESPONSE");
    }

    appendMessage("assistant", content);
    conversationHistory.push(
      { role: "user", content: prompt },
      { role: "assistant", content }
    );
    conversationHistory = conversationHistory.slice(-MAX_HISTORY_TURNS);
    setActivity("Ready.");
  } catch (error) {
    const description = error instanceof NativeResponseError
      ? error.message
      : describeNativeMessagingError(error);
    appendMessage("error", description);
    setActivity(description);

    if (!(error instanceof NativeResponseError)) {
      connection.connected = false;
      connection.configured = false;
      setConnectionView("error", "Not connected");
    }
  } finally {
    isSending = false;
    renderConnectionDetails();
    renderComposerState();
    elements.prompt.focus();
  }
}

function validatedContextSnapshot() {
  const screenshotDataUrl = isSupportedScreenshot(pageContext.screenshotDataUrl)
    ? pageContext.screenshotDataUrl
    : "";

  if (screenshotDataUrl && dataUrlByteLength(screenshotDataUrl) > MAX_SCREENSHOT_BYTES) {
    throw new Error("The attached screenshot exceeds the 5 MB limit. Clear it and attach a new screenshot.");
  }

  return {
    title: boundText(pageContext.title, MAX_TITLE_CHARS),
    url: boundText(pageContext.url, MAX_URL_CHARS),
    selection: boundText(pageContext.selection, MAX_SELECTION_CHARS),
    pageText: boundText(pageContext.pageText, MAX_PAGE_TEXT_CHARS),
    screenshotDataUrl
  };
}

function appendMessage(role, content) {
  elements.welcome.hidden = true;

  const article = document.createElement("article");
  article.className = `message ${role}`;

  const label = document.createElement("p");
  label.className = "message-role";
  label.textContent = role === "user" ? "You" : role === "assistant" ? "AI365" : "Connection error";

  const body = document.createElement("div");
  body.className = "message-body";
  // Never interpret model or webpage text as HTML.
  body.textContent = String(content || "");

  article.append(label, body);
  elements.messages.append(article);
  elements.messages.scrollTop = elements.messages.scrollHeight;
}

function renderComposerState() {
  const length = Math.min(elements.prompt.value.length, MAX_PROMPT_CHARS);
  elements.promptCount.textContent = `${formatNumber(length)} / ${formatNumber(MAX_PROMPT_CHARS)}`;
  elements.send.disabled = isSending
    || !connection.connected
    || !connection.configured
    || !elements.prompt.value.trim();
  elements.send.textContent = isSending ? "Sending…" : "Send";
}

function updateConnectionFromResponse(response) {
  if (!response || typeof response !== "object") {
    return;
  }

  // A structured response proves that the local bridge is reachable even
  // when the model request itself failed (for example, a rate limit).
  connection.connected = true;
  connection.configured = response.configured === true;
  connection.model = boundText(response.model, 200);
  connection.supportsVision = response.supportsVision === true;
  connection.version = boundText(response.version, 100);

  if (connection.configured) {
    setConnectionView("connected", "Connected");
  } else {
    setConnectionView("warning", "Setup needed");
  }
  renderConnectionDetails();
}

function setConnectionView(state, label) {
  elements.connectionLabel.textContent = label;
  elements.connectionDot.classList.toggle("connected", state === "connected");
  elements.connectionDot.classList.toggle("error", state === "error");
}

function renderConnectionDetails() {
  const details = [];
  if (connection.model) {
    details.push(connection.model);
  }
  if (connection.version) {
    details.push(`bridge ${connection.version}`);
  }
  if (connection.connected && connection.supportsVision) {
    details.push("vision ready");
  }
  elements.connectionDetails.textContent = details.join(" · ");
}

function setActivity(message) {
  elements.activity.textContent = message;
}

function sendNativeMessage(message, timeoutMs) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => {
      if (settled) {
        return;
      }
      settled = true;
      reject(new Error("AI365 did not respond in time."));
    }, timeoutMs);

    chrome.runtime.sendNativeMessage(NATIVE_HOST, message, (response) => {
      const runtimeError = chrome.runtime.lastError;
      if (settled) {
        return;
      }

      settled = true;
      clearTimeout(timer);

      if (runtimeError) {
        reject(new Error(runtimeError.message));
        return;
      }

      if (!response || typeof response !== "object") {
        reject(new Error("The AI365 browser bridge returned an invalid response."));
        return;
      }

      resolve(response);
    });
  });
}

function describeNativeMessagingError(error) {
  const message = error instanceof Error ? error.message : String(error || "");
  const lower = message.toLowerCase();

  if (lower.includes("specified native messaging host not found")) {
    return "AI365 browser support is not installed. Run AI365 Setup with browser support enabled, then restart the browser.";
  }
  if (lower.includes("access to the specified native messaging host is forbidden")) {
    return "This extension is not authorized to use the AI365 browser bridge. Reinstall matching versions of AI365 and the extension.";
  }
  if (lower.includes("native host has exited") || lower.includes("communicating with the native messaging host")) {
    return "The AI365 browser bridge stopped unexpectedly. Retry; if it happens again, repair AI365 from Setup.";
  }
  if (lower.includes("message length exceeded") || lower.includes("too large")) {
    return "The request is too large for the AI365 browser bridge. Clear some attached context and try again.";
  }
  if (lower.includes("did not respond in time")) {
    return "AI365 took too long to respond. Retry, or choose a faster model in AI365 Settings.";
  }

  return message
    ? `AI365 could not connect to its browser bridge: ${message}`
    : "AI365 could not connect to its browser bridge. Run AI365 Setup with browser support enabled.";
}

function describeHostResponseError(response) {
  const code = typeof response?.errorCode === "string" ? response.errorCode.toUpperCase() : "";
  const hostMessage = boundText(response?.error, 2_000).trim();

  const messages = {
    NOT_CONFIGURED: "No AI model is configured. Open AI365 Settings in an Office app and choose a model.",
    MODEL_NOT_CONFIGURED: "No AI model is configured. Open AI365 Settings in an Office app and choose a model.",
    VISION_NOT_SUPPORTED: "The selected model cannot read screenshots. Clear the screenshot or choose a vision model in AI365 Settings.",
    CONTEXT_TOO_LARGE: "The attached page context is too large. Clear one or more attachments and try again.",
    PROMPT_TOO_LARGE: "The message exceeds AI365's 16,000-character limit.",
    BUSY: "AI365 is busy with another request. Wait a moment and try again.",
    RATE_LIMITED: "The AI provider is rate-limiting requests. Wait a moment and try again.",
    AUTHENTICATION_FAILED: "AI365 could not authenticate with the selected provider. Check AI365 Settings and sign in again.",
    UNAUTHORIZED_ORIGIN: "This extension is not authorized to use the installed AI365 browser bridge. Reinstall matching versions."
  };

  return messages[code] || hostMessage || `AI365 could not complete the request${code ? ` (${code})` : ""}.`;
}

function describePageAccessError(error) {
  const message = error instanceof Error ? error.message : String(error || "");
  const lower = message.toLowerCase();

  if (
    lower.includes("cannot access contents") ||
    lower.includes("missing host permission") ||
    lower.includes("cannot be scripted") ||
    lower.includes("the extensions gallery cannot be scripted")
  ) {
    return "This page does not allow extensions to read its text. Try a regular webpage, or attach a visible screenshot instead.";
  }
  if (lower.includes("active tab") || lower.includes("permission")) {
    return "Click the AI365 toolbar button while viewing this tab to grant temporary access, then try again.";
  }

  return message || "AI365 could not access the current page.";
}

function contextAttachedMessage(kind) {
  if (kind === "selection") {
    return "Selection attached until you clear it.";
  }
  if (kind === "page") {
    return "Readable page text attached until you clear it.";
  }
  if (kind === "screenshot") {
    return "Visible screenshot attached until you clear it.";
  }
  return "Page context attached until you clear it.";
}

function emptyContext() {
  return {
    title: "",
    url: "",
    selection: "",
    pageText: "",
    screenshotDataUrl: ""
  };
}

function boundText(value, limit) {
  return typeof value === "string" ? value.slice(0, limit) : "";
}

function isSupportedScreenshot(value) {
  return typeof value === "string" && /^data:image\/(?:jpeg|png|webp);base64,[a-z0-9+/]+=*$/i.test(value);
}

function dataUrlByteLength(dataUrl) {
  if (typeof dataUrl !== "string") {
    return 0;
  }

  const commaIndex = dataUrl.indexOf(",");
  if (commaIndex < 0) {
    return 0;
  }

  const base64 = dataUrl.slice(commaIndex + 1);
  const padding = base64.endsWith("==") ? 2 : base64.endsWith("=") ? 1 : 0;
  return Math.max(0, Math.floor(base64.length * 3 / 4) - padding);
}

function formatBytes(bytes) {
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString();
}

function createRequestId() {
  if (typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

class NativeResponseError extends Error {
  constructor(message, errorCode) {
    super(message);
    this.name = "NativeResponseError";
    this.errorCode = errorCode || "";
  }
}
