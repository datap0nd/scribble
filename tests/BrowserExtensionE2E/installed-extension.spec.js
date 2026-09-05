const { test, expect, chromium } = require('@playwright/test');
const http = require('http');
const path = require('path');

test('installed extension observes generic controls and uses its real trusted-input broker', async ({}, testInfo) => {
  test.setTimeout(60000);
  const server = http.createServer((request, response) => {
    response.setHeader('Content-Type', 'text/html; charset=utf-8');
    response.end(`<!doctype html><title>Generic controls fixture</title>
      <style>label{display:inline-block;padding:12px;border:1px solid}input{position:absolute;opacity:0;width:18px;height:18px}</style>
      <label><input type="checkbox"> Include supporting details</label>
      <div id="listener" style="width:180px;height:40px">View delivery options</div>
      <canvas width="150" height="50" aria-label="Diagram"></canvas>
      <script>document.querySelector('#listener').addEventListener('click', event => event.currentTarget.textContent='Delivery options loaded');</script>`);
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const url = `http://127.0.0.1:${server.address().port}/fixture`;
  const extensionPath = path.resolve(__dirname, '../../src/Scribble.BrowserExtension');
  let context;
  try {
    context = await chromium.launchPersistentContext('', {
      channel: 'chromium', headless: true, viewport: {width:1000,height:800},
      args: [`--disable-extensions-except=${extensionPath}`, `--load-extension=${extensionPath}`]
    });
    await context.tracing.start({screenshots:true,snapshots:true,sources:true});
    const worker = context.serviceWorkers()[0] || await context.waitForEvent('serviceworker');
    const extensionId = worker.url().split('/')[2];
    const panel = await context.newPage();
    await panel.goto(`chrome-extension://${extensionId}/sidepanel.html`);
    const opened = context.waitForEvent('page');
    const tabId = await panel.evaluate(async url => {
      const tab = await chrome.tabs.create({url,active:false});
      workTabIds = [tab.id,null,null,null,null];
      lastWorkSlot = 1;
      return tab.id;
    }, url);
    const workPage = await opened;
    await workPage.waitForLoadState('load');
    const snapshot = await panel.evaluate(tabId => inspectWorkTab(tabId), tabId);
    expect(snapshot.perceptionLimitation).toBeFalsy();
    expect(snapshot.accessibility.methods).toContain('DOMSnapshot');
    const checkbox = snapshot.controls.find(c => c.name === 'Include supporting details');
    expect(checkbox?.proxy).toBe(true);
    expect(snapshot.controls.some(c => c.name === 'View delivery options')).toBe(true);
    expect(await panel.evaluate(tabId => snapshotImagesByTab.get(tabId), tabId)).toMatch(/^data:image\/jpeg;base64,/);
    // Exercise the installed service worker's bounded input protocol. Native
    // host policy and model selection have their own tests; no native endpoint
    // is mocked or claimed to have executed in this Chromium fixture.
    await panel.evaluate(async ({tabId,control}) => {
      const resolved = await runPageAgent(tabId,'actionability',{ref:control.ref,revision:control.revision});
      if (!resolved.receivesEvents) throw new Error('The visible proxy is obstructed.');
      await dispatchCdpBatch(tabId,[
        {command:'Input.dispatchMouseEvent',params:{type:'mouseMoved',x:resolved.x,y:resolved.y}},
        {command:'Input.dispatchMouseEvent',params:{type:'mousePressed',x:resolved.x,y:resolved.y}},
        {command:'Input.dispatchMouseEvent',params:{type:'mouseReleased',x:resolved.x,y:resolved.y}}
      ]);
    }, {tabId,control:checkbox});
    const after = await panel.evaluate(tabId => inspectWorkTab(tabId),tabId);
    expect(after.controls.find(c => c.name === 'Include supporting details').selected).toBe(true);
    await expect(panel.evaluate(() => sendOperatorMessage({type:'inspectSemantics',tabId:999999}))).rejects.toThrow(/registered work tab/);
  } finally {
    try {
      if (context) {
        if (testInfo.status !== testInfo.expectedStatus) await context.tracing.stop({path:testInfo.outputPath('installed-extension-trace.zip')});
        else await context.tracing.stop();
      }
    } finally {
      try { if (context) await context.close(); }
      finally { await new Promise(resolve => server.close(resolve)); }
    }
  }
});
