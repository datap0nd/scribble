const { test, expect } = require("@playwright/test");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

function eventSlot() {
  const slot = { listener: null };
  slot.api = {
    addListener(listener) { slot.listener = listener; },
    removeListener() {}
  };
  return slot;
}

function createHarness() {
  const connect = eventSlot();
  const detach = eventSlot();
  const portMessage = eventSlot();
  const portDisconnect = eventSlot();
  const posted = [];
  const detachCalls = [];
  const never = new Promise(() => {});
  const chrome = {
    runtime: {
      onInstalled: eventSlot().api,
      onStartup: eventSlot().api,
      onConnect: connect.api,
      getURL: (value) => `chrome-extension://fixture/${value}`
    },
    action: { onClicked: eventSlot().api },
    contextMenus: {
      onClicked: eventSlot().api,
      removeAll: async () => {},
      create: () => {}
    },
    sidePanel: { open: async () => {} },
    tabs: { get: async () => ({ url: "https://example.test/" }) },
    debugger: {
      onDetach: detach.api,
      attach: async () => {},
      sendCommand: async () => never,
      detach: async ({ tabId }) => { detachCalls.push(tabId); }
    }
  };
  const source = fs.readFileSync(
    path.resolve(__dirname, "../../src/Scribble.BrowserExtension/background.js"),
    "utf8"
  );
  const sandbox = { chrome, console, Set, Map, Promise, setTimeout };
  vm.createContext(sandbox);
  vm.runInContext(source, sandbox);
  const port = {
    name: "scribble-browser-operator",
    sender: { url: "chrome-extension://fixture/sidepanel.html" },
    onMessage: portMessage.api,
    onDisconnect: portDisconnect.api,
    postMessage: (message) => posted.push(message),
    disconnect: () => {}
  };
  connect.listener(port);
  return {
    detach,
    portMessage,
    portDisconnect,
    posted,
    detachCalls,
    validateKey(params) {
      sandbox.__scribbleKeyParams = params;
      return vm.runInContext(
        'validateCdpParams("Input.dispatchKeyEvent", __scribbleKeyParams)',
        sandbox
      );
    }
  };
}

async function registerAndStart(harness) {
  await harness.portMessage.listener({
    type: "registerWorkTabs",
    requestId: "register",
    chatId: "chat",
    tabIds: [42]
  });
  void harness.portMessage.listener({
    type: "cdpAction",
    requestId: "action",
    tabId: 42,
    commands: [{
      command: "Input.dispatchMouseEvent",
      params: { type: "mouseMoved", x: 10, y: 10 }
    }]
  });
  await new Promise((resolve) => setTimeout(resolve, 0));
}

test("canceled_by_user debugger detach becomes a clean operator stop", async () => {
  const harness = createHarness();
  await registerAndStart(harness);
  harness.detach.listener({ tabId: 42 }, "canceled_by_user");
  expect(harness.posted).toContainEqual({
    type: "operatorDetached",
    tabId: 42,
    reason: "canceled_by_user"
  });
});

test("Stop or Clear-style detachAll detaches the live work tab", async () => {
  const harness = createHarness();
  await registerAndStart(harness);
  await harness.portMessage.listener({ type: "detachAll", requestId: "stop" });
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(harness.detachCalls).toContain(42);
  expect(harness.posted).toContainEqual({ type: "detachedAll", requestId: "stop" });
});

test("side-panel port disconnect detaches the live work tab", async () => {
  const harness = createHarness();
  await registerAndStart(harness);
  harness.portDisconnect.listener();
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(harness.detachCalls).toContain(42);
});

test("Enter dispatch includes browser-compatible keyboard metadata", () => {
  const harness = createHarness();
  const validated = harness.validateKey({ type: "keyDown", key: "Enter" });
  expect(validated.key).toBe("Enter");
  expect(validated.code).toBe("Enter");
  expect(validated.windowsVirtualKeyCode).toBe(13);
  expect(validated.text).toBe('\r');
  expect(harness.validateKey({type:'keyUp',key:'Enter',text:'arbitrary caller text'}).text).toBe('');
  expect(harness.validateKey({type:'keyDown',key:' ',text:'arbitrary caller text'}).text).toBe(' ');
});
