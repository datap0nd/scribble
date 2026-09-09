const { test, expect } = require('@playwright/test');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const source = fs.readFileSync(path.resolve(__dirname, '../../src/Scribble.BrowserExtension/sidepanel.js'), 'utf8');

function controller() {
  const replies = [], prompts = [], activities = [];
  let command = null;
  const sandbox = {
    URL, URLSearchParams, location: { search: '?suite=test-suite&token=test-case' },
    crypto: { randomUUID: () => 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' },
    setInterval: () => 1, setTimeout, Event: class {}, PING_TIMEOUT_MS: 1000,
    connection: { connected: true, configured: true }, isSending: false, activeRecovery: null,
    elements: { prompt: { value: '', dispatchEvent() {} }, composer: { requestSubmit() { sandbox.isSending = false; } } },
    chrome: { tabs: { query: async () => [{ id: 123 }], get: async id => ({ id, url: 'http://127.0.0.1:9999/operations.html' }) } },
    clearChat: async () => {}, renderCurrentTab: async () => {}, setActivity: text => activities.push(text),
    sendChatMessage: async () => { prompts.push(sandbox.elements.prompt.value); },
    sendNativeMessage: async message => {
      const data = JSON.parse(message.taskData);
      if (data.action === 'complete') { replies.push(data); return { ok: true, content: 'null' }; }
      return { ok: true, content: JSON.stringify(command) };
    }
  };
  vm.createContext(sandbox);
  vm.runInContext(source.slice(source.indexOf('// Headed suite controller.')), sandbox);
  return { sandbox, replies, prompts, activities, setCommand(value) { command = value; },
    execute(value) { sandbox.command = value; return vm.runInContext('executeSuiteCommand(command)', sandbox); },
    poll() { return vm.runInContext('pollSuite()', sandbox); }
  };
}

test('suite loads the synthetic source and submits through the normal chat entry point', async () => {
  const c = controller();
  await c.execute({ id: 'load', action: 'load', runId: 'run', sourceUrl: 'http://127.0.0.1:9999/operations.html' });
  expect(c.sandbox.scribbleSuiteSourceTab).toBe(123);
  expect(c.sandbox.scribbleLabRun).toBe('run');
  await c.execute({ id: 'send', action: 'submit', prompt: 'Exact known case prompt' });
  expect(c.prompts).toEqual(['Exact known case prompt']);
  expect(c.replies.map(r => r.error)).toEqual(['', '']);
});

test('suite rejects unrelated source pages and unconfigured models without submitting', async () => {
  const c = controller();
  await c.execute({ id: 'bad-source', action: 'load', sourceUrl: 'https://example.com/operations.html' });
  expect(c.replies[0].error).toContain('Invalid suite source');
  c.sandbox.connection.configured = false;
  await c.execute({ id: 'offline', action: 'submit', prompt: 'Do not send' });
  expect(c.replies[1].error).toContain('not ready');
  expect(c.prompts).toEqual([]);
});

test('suite polling consumes a command once and returns full model errors', async () => {
  const c = controller();
  c.setCommand({ id: 'once', action: 'submit', prompt: 'Run exactly once' });
  await c.poll(); await new Promise(resolve => setImmediate(resolve));
  await c.poll(); await new Promise(resolve => setImmediate(resolve));
  expect(c.prompts).toHaveLength(1);
  c.sandbox.sendChatMessage = async () => { c.sandbox.activeRecovery = { blocker: 'Model failed: diagnostic tail END_OF_ERROR' }; };
  await c.execute({ id: 'failed', action: 'submit', prompt: 'Fail' });
  expect(c.replies.at(-1).error).toContain('END_OF_ERROR');
});

test('suite Stop uses the normal stop action and acknowledges after the request ends', async () => {
  const c = controller();
  c.sandbox.isSending = true;
  await c.execute({ id: 'stop', action: 'stop' });
  expect(c.sandbox.isSending).toBe(false);
  expect(c.replies[0].error).toBe('');
});

test('headed controller reads the fixture tab rather than its extension page', async () => {
  const c = controller();
  c.sandbox.scribbleSuiteSourceTab = 123;
  vm.runInContext(source.slice(source.indexOf('async function getActiveTab()'), source.indexOf('async function renderCurrentTab()')), c.sandbox);
  const tab = await vm.runInContext('getActiveTab()', c.sandbox);
  expect(tab.id).toBe(123);
  expect(tab.url).toContain('/operations.html');
});
