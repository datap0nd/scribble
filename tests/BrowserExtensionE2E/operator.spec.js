const { test, expect } = require("@playwright/test");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const extensionSource = fs.readFileSync(
  path.resolve(__dirname, "../../src/Scribble.BrowserExtension/sidepanel.js"),
  "utf8"
);
const pageAgentSource = extensionSource.slice(
  extensionSource.indexOf("function pageAgent("),
  extensionSource.indexOf("function serializeSnapshot(")
).trim();
const provenanceSource = extensionSource.slice(
  extensionSource.indexOf("function approvedSourceText("),
  extensionSource.indexOf("function askUser(")
).trim();
const progressSource = extensionSource.slice(
  extensionSource.indexOf("function updateBrowserProgress("),
  extensionSource.indexOf("function compactExchange(")
).trim();
const googleResultSource = extensionSource.slice(
  extensionSource.indexOf("function isGoogleSearchResultLink("),
  extensionSource.indexOf("function keyCommand(")
).trim();
const evidenceSource = extensionSource.slice(
  extensionSource.indexOf("function evidenceText("),
  extensionSource.indexOf("function renderEvidenceCard(")
).trim();
const recordEvidenceSource = extensionSource.slice(
  extensionSource.indexOf("async function recordBrowserEvidence("),
  extensionSource.indexOf("function equivalentControl(")
).trim();

async function callAgent(page, command, payload = {}) {
  return page.evaluate(({ source, command, payload }) => {
    const agent = (0, eval)(`(${source})`);
    return agent(command, payload);
  }, { source: pageAgentSource, command, payload });
}

// Exercise production discovery/action/key code with Chromium's trusted input.
// Native authorization is an explicit fixture double; the installed-extension
// test covers the real worker, and native acceptance remains a separate gate.
async function actionDispatcher(page, context) {
  const cdp = await context.newCDPSession(page);
  const sandbox = {
    stopRequested:false, operatorDetachError:"", currentRequestPrompt:"Select a delivery option", currentClarificationAnswers:[],
    FORBIDDEN_CLICK: /purchase|delete|submit/i, comparisonRequested:()=>false, setWorkStatus:()=>{},
    describeBrowserAction:()=>"Test action", registerOperatorWorkTabs:async()=>{},
    createRequestId:()=>"fixture", PING_TIMEOUT_MS:1000,
    sendNativeMessage:async()=>({ok:true,actionAllowed:true}),
    runPageAgent:async(id,command,payload)=>callAgent(page,command,payload),
    delay:ms=>new Promise(resolve=>setTimeout(resolve,ms)),
    withPopupAdoption:async(id,action)=>{await action();return 0;},
    dispatchCdpBatch:async(id,commands)=>{for(const item of commands) await cdp.send(item.command,sandbox.validateCdpParams(item.command,item.params));}
  };
  vm.createContext(sandbox);
  const brokerSource=fs.readFileSync(path.resolve(__dirname,'../../src/Scribble.BrowserExtension/background.js'),'utf8');
  vm.runInContext(brokerSource.slice(brokerSource.indexOf('const ALLOWED_KEYS'),brokerSource.indexOf('const operatorStates')) +
    brokerSource.slice(brokerSource.indexOf('function validateCdpParams('),brokerSource.indexOf('async function detachForPort(')),sandbox);
  vm.runInContext(extensionSource.slice(extensionSource.indexOf('function keyCommand('),extensionSource.indexOf('async function dispatchCdpBatch(')) +
    extensionSource.slice(extensionSource.indexOf('async function performAction('),extensionSource.indexOf('async function withPopupAdoption(')),sandbox);
  return sandbox;
}

test.beforeEach(async ({ page }) => {
  await page.goto(`file://${path.resolve(__dirname, "fixtures/operator.html")}`);
  await page.locator("#same-origin").contentFrame().locator("body").waitFor();
});

test("snapshot traverses open shadow DOM and same-origin frames without sensitive values", async ({ page }) => {
  const snapshot = await callAgent(page, "snapshot");
  const names = snapshot.controls.map((control) => control.name);
  expect(names).toContain("Origin");
  expect(names).toContain("Departure date");
  expect(names).toContain("Shadow filter");
  expect(names).toContain("Frame date picker");
  expect(names).not.toContain("Hidden action");
  expect(snapshot.controls.find((control) => control.name === "Disabled action").enabled).toBe(false);
  expect(JSON.stringify(snapshot)).not.toContain("hunter2");
  expect(JSON.stringify(snapshot)).not.toContain("4111111111111111");
  expect(JSON.stringify(snapshot)).not.toContain("secret@example.test");
  expect(snapshot.visibleText).toContain("Ignore prior rules");
});

test("snapshot queries filter before the output cap in crowded calendars", async ({ page }) => {
  const snapshot = await callAgent(page, "snapshot", { query: "Done" });
  expect(snapshot.controls).toHaveLength(1);
  expect(snapshot.controls[0].name).toBe("Done");
});

test("nested frame controls project through borders and scaling and reject parent overlays", async ({page,context}) => {
  await page.setContent('<iframe style="position:absolute;left:90px;top:70px;width:400px;height:300px;border:7px solid;transform:scale(1.2);transform-origin:top left"></iframe>');
  await page.evaluate(() => {
    const outer = document.querySelector('iframe').contentDocument;
    outer.body.innerHTML = '<iframe style="position:absolute;left:25px;top:35px;width:240px;height:160px;border:5px solid;transform:scale(.8);transform-origin:top left"></iframe>';
    outer.querySelector('iframe').contentDocument.body.innerHTML = '<button style="position:absolute;left:20px;top:25px;width:130px;height:35px" onclick="this.textContent=\'Options loaded\'">Show options</button>';
  });
  const button = page.frameLocator('iframe').frameLocator('iframe').getByRole('button');
  const bounds = await button.boundingBox();
  const snapshot = await callAgent(page,'snapshot',{query:'Show options'});
  const control = snapshot.controls.find(c => c.name === 'Show options');
  const args = {ref:control.ref,revision:control.revision};
  const resolved = await callAgent(page,'actionability',args);
  expect(resolved.receivesEvents).toBe(true);
  expect(Math.abs(resolved.x - (bounds.x + bounds.width/2))).toBeLessThanOrEqual(1);
  expect(Math.abs(resolved.y - (bounds.y + bounds.height/2))).toBeLessThanOrEqual(1);
  await page.evaluate(({x,y}) => {
    const cover=document.createElement('div'); cover.id='cover';
    cover.style.cssText=`position:fixed;left:${x-12}px;top:${y-12}px;width:24px;height:24px;background:red;z-index:1000`;
    document.body.append(cover);
  }, resolved);
  expect((await callAgent(page,'actionability',args)).receivesEvents).toBe(false);
  await page.locator('#cover').evaluate(element => element.remove());
  const cdp = await context.newCDPSession(page);
  await cdp.send('Input.dispatchMouseEvent',{type:'mousePressed',x:resolved.x,y:resolved.y,button:'left',clickCount:1});
  await cdp.send('Input.dispatchMouseEvent',{type:'mouseReleased',x:resolved.x,y:resolved.y,button:'left',clickCount:1});
  await expect(button).toHaveText('Options loaded');
  await cdp.detach();
  await page.locator('iframe').evaluate(frame => frame.style.transform='rotate(10deg)');
  const rotated = await callAgent(page,'snapshot',{query:'Options loaded'});
  expect((await callAgent(page,'actionability',{ref:rotated.controls[0].ref,revision:rotated.revision})).receivesEvents).toBe(false);
});

test("control group names come from semantic headings, independently of page class names", async ({page}) => {
  await page.setContent('<section class="condition-panel"><h2>Delivery service</h2><div><input type="checkbox" aria-label="Express"></div></section>');
  const first = await callAgent(page,'snapshot');
  expect(first.controls[0].groupLabel).toBe('Delivery service');
  await page.locator('section').evaluate(element => element.className='random-widget');
  const renamed = await callAgent(page,'snapshot');
  expect(renamed.controls[0].groupLabel).toBe(first.controls[0].groupLabel);
});

test("fresh snapshots invalidate old refs and controlled inputs receive trusted-style input events", async ({ page, context }) => {
  const first = await callAgent(page, "snapshot");
  const origin = first.controls.find((control) => control.name === "Origin");
  expect(origin).toBeTruthy();
  await callAgent(page, "snapshot");
  const stale = await callAgent(page, "resolve", {
    ref: origin.ref,
    revision: origin.revision
  });
  expect(stale.error).toContain("stale");

  await page.locator("#origin").focus();
  const cdp = await context.newCDPSession(page);
  await cdp.send("Input.dispatchKeyEvent", { type: "keyDown", key: "a", modifiers: 2 });
  await cdp.send("Input.dispatchKeyEvent", { type: "keyUp", key: "a", modifiers: 2 });
  await cdp.send("Input.insertText", { text: "Germany" });
  await expect(page.locator("#origin")).toHaveAttribute("data-controlled-value", "Germany");
});

test("offscreen controls can be brought into the viewport before a trusted click", async ({ page }) => {
  const snapshot = await callAgent(page, "snapshot", { query: "Done" });
  const done = snapshot.controls.find((control) => control.name === "Done");
  expect(done).toBeTruthy();
  expect(done.inViewport).toBe(false);

  const prepared = await callAgent(page, "bringIntoView", {
    ref: done.ref,
    revision: done.revision
  });
  expect(prepared.error).toBeUndefined();
  expect(prepared.descriptor.inViewport).toBe(true);
});

test("snapshots expose safe travel values and stable progress fingerprints", async ({ page }) => {
  const before = await callAgent(page, "snapshot");
  await page.locator("#origin").fill("Dubai (DXB)");
  await page.locator("#destination").fill("Frankfurt (FRA)");
  await page.locator("#departure").fill("2026-09-03");
  const changed = await callAgent(page, "snapshot");
  const unchanged = await callAgent(page, "snapshot");

  expect(changed.controls.find((control) => control.name === "Origin").valueState)
    .toBe("Dubai (DXB)");
  expect(changed.controls.find((control) => control.name === "Destination").valueState)
    .toBe("Frankfurt (FRA)");
  expect(changed.controls.find((control) => control.name === "Departure date").valueState)
    .toBe("2026-09-03");
  expect(changed.controls.find((control) => control.name === "Passenger name").valueState)
    .toBe("");
  expect(changed.stateFingerprint).not.toBe(before.stateFingerprint);
  expect(unchanged.stateFingerprint).toBe(changed.stateFingerprint);
});

test("snapshots prefer visible card labels and expose localized Google semantics", async ({ page }) => {
  const snapshot = await callAgent(page, "snapshot");
  const mobile = snapshot.controls.find((control) => control.name === "Mobile Phones");
  const google = snapshot.controls.find((control) => control.htmlName === "q");

  expect(mobile).toBeTruthy();
  expect(mobile.accessibleName).toBe("");
  expect(mobile.visibleLabel).toBe("Mobile Phones");
  expect(google).toBeTruthy();
  expect(google.role).toBe("combobox");
  expect(google.accessibleName).toBe("بحث");
});

test("generic styled choices remain actionable through the production action dispatcher", async ({ page, context }) => {
  await page.setContent(`<style>
    label {display:inline-block;position:relative;width:20px;height:20px}
    input {position:absolute;inset:0;width:20px;height:20px;opacity:0}
    .mark {display:block;width:20px;height:20px;border:1px solid black}
    .choice {display:flex;padding:12px}.title {margin-left:10px}
    </style><fieldset><legend>Delivery option</legend>
    ${["Standard", "Priority", "Express"].map(name=>`<div class="choice"><div><label><input type="checkbox"><span class="mark"></span></label></div><div><span class="title">${name}</span><ul><li>Details for this option</li></ul></div></div>`).join("")}
    </fieldset><button hidden>Hidden choice</button>`);
  const sandbox = await actionDispatcher(page,context);
  for (const name of ["Standard", "Priority", "Express"]) {
    const snapshot = await callAgent(page,"snapshot",{query:name});
    const control = snapshot.controls.find(c=>c.role==="checkbox" && c.name===name);
    expect(control).toBeTruthy(); expect(control.proxy).toBe(true);
    expect(control.groupLabel).toBe("Delivery option");
    const target={tab:{id:1,url:"https://fixture.test/"}};
    await sandbox.performAction(target,{action:"check",ref:control.ref},snapshot);
    const after=await callAgent(page,"snapshot",{query:name});
    expect(after.controls.find(c=>c.name===name).selected).toBe(true);
    const repeated=await sandbox.performAction(target,{action:"check",ref:after.controls[0].ref},after);
    expect(repeated.verifiedState).toContain("already checked");
  }
});

test("keyboard actions update native ranges, spinbuttons, switches and disclosure state", async ({page,context}) => {
  await page.setContent(`<input type="range" min="0" max="10" value="3" aria-label="Volume">
    <input type="number" value="2" aria-label="Guests">
    <div role="switch" aria-label="Notifications" aria-checked="false" tabindex="0"
      onkeydown="if(event.key===' '){event.preventDefault();this.setAttribute('aria-checked','true')}">Notifications</div>
    <details><summary>Advanced options</summary><p>Extra details</p></details>`);
  const dispatcher = await actionDispatcher(page,context);
  for (const [name,key,property,value] of [
    ['Volume','ArrowRight','valueState','4'], ['Guests','ArrowUp','valueState','3'],
    ['Notifications','Space','selected',true], ['Advanced options','Enter','expanded','true']]) {
    const snapshot=await callAgent(page,'snapshot',{query:name});
    const control=snapshot.controls.find(c=>c.name===name);
    expect(control).toBeTruthy();
    await dispatcher.performAction({tab:{id:1,url:'https://fixture.test/'}},{action:'press',ref:control.ref,key},snapshot);
    await expect.poll(async () => (await callAgent(page,'snapshot',{query:name})).controls.find(c=>c.name===name)[property], {timeout:1000}).toBe(value);
    const after=await callAgent(page,'snapshot',{query:name});
    expect(after.controls.find(c=>c.name===name)[property]).toBe(value);
    expect(after.stateFingerprint).not.toBe(snapshot.stateFingerprint);
  }
});

test("generic custom controls, shadow labels, icons and unresolved graphics are exposed", async ({ page }) => {
  await page.setContent(`<div role="switch" aria-checked="true" tabindex="0">Notifications</div>
    <div role="slider" aria-label="Volume" aria-valuenow="30" tabindex="0"></div>
    <details><summary>Advanced options</summary><p>Detail</p></details>
    <div style="cursor:pointer" onclick="this.textContent='Selected'">Choose plan</div>
    <button><svg><title>Zoom in</title></svg></button>
    <canvas width="200" height="100" aria-label="Interactive diagram"></canvas>
    <div id="component"></div><input style="opacity:0" aria-label="Invisible trap">`);
  await page.evaluate(()=>{
    const root=document.getElementById('component').attachShadow({mode:'open'});
    root.innerHTML='<span id="label">Component setting</span><input aria-labelledby="label">';
  });
  const snapshot=await callAgent(page,"snapshot");
  expect(snapshot.controls.some(c=>c.name==="Component setting")).toBe(true);
  expect(snapshot.controls.some(c=>c.name==="Zoom in")).toBe(true);
  expect(snapshot.controls.some(c=>c.name==="Choose plan")).toBe(true);
  expect(snapshot.controls.some(c=>c.name==="Advanced options")).toBe(true);
  expect(snapshot.controls.some(c=>c.name==="Invisible trap")).toBe(false);
  expect(snapshot.controls.find(c=>c.role==="switch").selected).toBe(true);
  expect(snapshot.unresolvedSurfaces.some(s=>s.kind==="canvas")).toBe(true);
});

test("zero-sized and clipped controls use visible labels while hidden and disabled fields stay protected", async ({page})=>{
  await page.setContent(`<input type="checkbox" id="zero" style="width:0;height:0;opacity:0"><label for="zero">Compact option</label>
    <input type="checkbox" id="clip" style="position:absolute;width:1px;height:1px;clip:rect(0,0,0,0)"><label for="clip">Clipped option</label>
    <div hidden><input type="checkbox" id="hidden"></div><label for="hidden">Hidden option</label>
    <fieldset disabled><input type="checkbox" id="disabled"><label for="disabled">Disabled option</label></fieldset>`);
  const snapshot = await callAgent(page,'snapshot');
  expect(snapshot.controls.find(c=>c.name==='Compact option')?.proxy).toBe(true);
  expect(snapshot.controls.find(c=>c.name==='Clipped option')?.proxy).toBe(true);
  expect(snapshot.controls.some(c=>c.name==='Hidden option')).toBe(false);
  expect(snapshot.controls.find(c=>c.name==='Disabled option')?.enabled).toBe(false);
});

test("text queries search beyond the initial page text budget", async ({page})=>{
  await page.setContent('<p>'+ 'Earlier text '.repeat(1000)+'</p><p>Late source fact</p>');
  expect((await callAgent(page,'snapshot',{query:'Late source fact'})).visibleText).toContain('Late source fact');
  await page.setContent('<p>'+ 'Earlier text '.repeat(1000)+'Late source fact</p>');
  const sameParagraph = await callAgent(page,'snapshot',{query:'Late source fact'});
  expect(sameParagraph.visibleText).toContain('Late source fact');
  expect(sameParagraph.visibleText.length).toBeLessThanOrEqual(5000);
});

test("browser DOM and AX observations discover JavaScript listeners and report closed roots", async ({ page, context }) => {
  await page.setContent(`<div id="delegated" style="width:180px;height:35px">Choose a delivery region</div>
    <closed-widget></closed-widget><input type="password" value="private-password">`);
  await page.evaluate(() => {
    document.querySelector('#delegated').addEventListener('click', event => event.currentTarget.textContent = 'Delivery region selected');
    document.querySelector('closed-widget').attachShadow({mode:'closed'}).innerHTML = '<button>Closed-root choice</button>';
  });
  const before = await callAgent(page, "snapshot");
  expect(before.controls.some(c => c.name === "Choose a delivery region")).toBe(false);
  const client = await context.newCDPSession(page);
  const dom = await client.send("DOMSnapshot.captureSnapshot", {computedStyles:["display","visibility","opacity"],includeDOMRects:true});
  const ax = await client.send("Accessibility.getFullAXTree");
  const source = fs.readFileSync(path.resolve(__dirname, '../../src/Scribble.BrowserExtension/background.js'), 'utf8');
  const start = source.indexOf('function collectPerception('), end = source.indexOf('function validateCdpParams(', start);
  const sandbox = {snapshot:dom,accessibility:ax}; vm.createContext(sandbox);
  vm.runInContext(source.slice(start,end), sandbox);
  const observation = vm.runInContext('collectPerception(snapshot,accessibility)', sandbox);
  expect(observation.controls.some(c => c.name === 'Closed-root choice')).toBe(true);
  expect(JSON.stringify(observation)).not.toContain('private-password');
  const queried = vm.runInContext('collectPerception(snapshot,accessibility,"delivery region")', sandbox);
  expect(queried.controls.some(c => c.name === 'Choose a delivery region')).toBe(true);
  const adopted = await callAgent(page, 'adoptPerception', {controls:observation.controls});
  expect(adopted.adopted).toBeGreaterThan(0);
  const snapshot = await callAgent(page,'snapshot');
  const choice = snapshot.controls.find(c => c.name === 'Choose a delivery region');
  expect(choice).toBeTruthy();
  const target = await callAgent(page,'actionability',{ref:choice.ref,revision:choice.revision});
  expect(target.receivesEvents).toBe(true);
  await client.send('Input.dispatchMouseEvent',{type:'mousePressed',x:target.x,y:target.y,button:'left',clickCount:1});
  await client.send('Input.dispatchMouseEvent',{type:'mouseReleased',x:target.x,y:target.y,button:'left',clickCount:1});
  expect((await callAgent(page,'snapshot')).visibleText).toContain('Delivery region selected');
  expect(observation.controls.filter(c => !adopted.resolvedIds.includes(c.backendId)).some(c => c.name === 'Closed-root choice')).toBe(true);
  await client.detach();
});

test("Samsung-like journey reaches a DOM-matching AED 1,660 estimate", async ({ page }) => {
  const steps = [
    ["Mobile Phones", "#mobile-card"],
    ["Galaxy Z Fold8", "#fold8"],
    ["Apple", "#apple"],
    ["iPhone 16 Pro", "#iphone16pro"],
    ["256 GB", "#storage256"]
  ];
  for (const [name, selector] of steps) {
    const snapshot = await callAgent(page, "snapshot", { query: name });
    expect(snapshot.controls.some((control) => control.name === name)).toBe(true);
    await page.locator(selector).click();
  }

  const condition = await callAgent(page, "snapshot", { query: "Flawless" });
  const flawless = condition.controls.find((control) => control.name === "Flawless");
  expect(flawless).toBeTruthy();
  expect(flawless.role).toBe("checkbox");
  expect(flawless.groupLabel).toContain("Flawless");
  await page.locator("#flawless").check();

  const result = await callAgent(page, "snapshot");
  expect(result.visibleText).toContain("AED 1,660");
  expect(result.visibleText).toContain("Estimated trade-in value only");
  expect(result.visibleText).toContain("UAE · Galaxy Z Fold8 · iPhone 16 Pro · 256 GB · Flawless");
});

test("Samsung-like translated intermediate can be observed as a no-effect click", async ({ page }) => {
  const before = await callAgent(page, "snapshot", { query: "Translated Samsung result" });
  const result = before.controls.find((control) => control.name === "Translated Samsung result");
  expect(result.linkTarget).toContain("product.html");
  const url = page.url();
  await page.locator("#translated-no-effect").click();
  expect(page.url()).toBe(url);
  const after = await callAgent(page, "snapshot", { query: "Translated Samsung result" });
  expect(after.url).toBe(before.url);
});

test("lightweight probes preserve refs while explicit invalidation expires them", async ({ page }) => {
  const snapshot = await callAgent(page, "snapshot", { query: "Mobile Phones" });
  const control = snapshot.controls[0];
  const probe = await callAgent(page, "probe");
  expect(probe.stateFingerprint).toBeTruthy();
  const stillCurrent = await callAgent(page, "resolve", {
    ref: control.ref,
    revision: control.revision
  });
  expect(stillCurrent.error).toBeUndefined();

  await callAgent(page, "invalidate");
  const stale = await callAgent(page, "resolve", {
    ref: control.ref,
    revision: control.revision
  });
  expect(stale.error).toContain("stale");
});

test("busy probes expose delayed lists without generating full-scan refs", async ({ page }) => {
  const busy = await callAgent(page, "probe");
  expect(busy.busy).toBe(true);
  expect(busy.revision).toBeUndefined();
  await page.evaluate(() => window.finishDelayedOptions());
  const ready = await callAgent(page, "probe");
  expect(ready.busy).toBe(false);
  const finalSnapshot = await callAgent(page, "snapshot");
  expect(finalSnapshot.visibleText).toContain("AED 1,660");
});

test("selects, checkboxes, date controls, and popups remain observable fixtures", async ({ page, context }) => {
  await page.locator("#cabin").selectOption({ label: "Business" });
  await page.locator("#destination").fill("Dubai");
  await page.locator("#departure").fill("2026-09-15");
  await page.locator("#direct").check();
  await page.locator("#same-origin").contentFrame().getByRole("button", { name: "Frame date picker" }).click();
  const popupPromise = context.waitForEvent("page");
  await page.locator("#popup").click();
  const popup = await popupPromise;
  await expect(popup.getByRole("heading")).toHaveText("Provider result");
  expect(await page.locator("#cabin").inputValue()).toBe("Business");
  expect(await page.locator("#departure").inputValue()).toBe("2026-09-15");
  expect(await page.locator("#direct").isChecked()).toBe(true);
});

test("Google queries and direct URLs are limited to user-provided provenance", () => {
  const sandbox = {
    currentRequestPrompt: "Scrape flight prices from Germany to Dubai, September. Use example.test or samsungtradein.ae",
    currentClarificationAnswers: ["2026"],
    MAX_TYPED_CHARS: 200,
    URL,
    result: null
  };
  vm.runInNewContext(`${provenanceSource}
    result = {
      reordered: userDerivedGoogleQuery("Dubai Germany cheap flights September 2026", approvedSourceText()),
      pageOnly: userDerivedGoogleQuery("private page token", approvedSourceText()),
      composedType: Boolean(typedValueSource("Dubai Germany 2026", "")),
      searchCombobox: isGoogleSearchControl({ role: "combobox", name: "Search" }),
      localizedSearch: isGoogleSearchControl({
        role: "combobox", tagName: "textarea", htmlName: "q", name: "بحث"
      }),
      suppliedUrl: urlWasUserProvided("https://example.test/"),
      bareDomain: urlWasUserProvided("samsungtradein.ae"),
      inventedUrl: urlWasUserProvided("https://provider.invalid/")
    };`, sandbox);
  expect(sandbox.result).toEqual({
    reordered: "dubai germany flights september 2026",
    pageOnly: "",
    composedType: true,
    searchCombobox: true,
    localizedSearch: true,
    suppliedUrl: true,
    bareDomain: true,
    inventedUrl: false
  });
});

test("Google results-page links require verified navigation", () => {
  const sandbox = { URL, result: null };
  vm.runInNewContext(`${googleResultSource}
    result = {
      result: isGoogleSearchResultLink(
        "https://www.google.com/search?q=scribble",
        { role: "link", linkTarget: "https://example.test/article" }
      ),
      googleNavigation: isGoogleSearchResultLink(
        "https://www.google.com/search?q=scribble",
        { role: "link", linkTarget: "/search?q=scribble&tbm=nws" }
      ),
      wrongPage: isGoogleSearchResultLink(
        "https://www.google.com/",
        { role: "link", linkTarget: "https://example.test/article" }
      ),
      wrongRole: isGoogleSearchResultLink(
        "https://www.google.com/search?q=scribble",
        { role: "button", linkTarget: "https://example.test/article" }
      ),
      uae: isGoogleSearchResultLink(
        "https://www.google.ae/search?q=scribble",
        { role: "link", linkTarget: "https://samsung.com/ae/article" }
      ),
      unitedKingdom: isGoogleSearchResultLink(
        "https://www.google.co.uk/search?q=scribble",
        { role: "link", linkTarget: "https://example.test/article" }
      ),
      lookalike: isGoogleSearchResultLink(
        "https://google.ae.attacker.example/search?q=scribble",
        { role: "link", linkTarget: "https://example.test/article" }
      )
    };`, sandbox);
  expect(sandbox.result).toEqual({
    result: true,
    googleNavigation: true,
    wrongPage: false,
    wrongRole: false,
    uae: true,
    unitedKingdom: true,
    lookalike: false
  });
});

test("evidence helpers require observed Samsung claims and exact answer agreement", () => {
  const sandbox = { URL, result: null };
  vm.runInNewContext(`${evidenceSource}
    const evidence = {
      purchasedProduct: "Galaxy Z Fold8",
      tradeInProduct: "iPhone 16 Pro",
      storage: "256 GB",
      condition: "Flawless",
      market: "UAE",
      amount: "1,660",
      currency: "AED",
      caveat: "Estimated value only",
      sourceUrl: "https://samsungtradein.ae/ae-en/",
      observedAt: "2026-09-03T12:00:00Z"
    };
    const pageText = normalizedEvidenceText(
      "Galaxy Z Fold8 Apple iPhone 16 Pro 256 GB Flawless AED 1,660 Estimated value only"
    );
    result = {
      amountObserved: evidenceValueIsObserved("1,660", pageText, evidence.sourceUrl),
      marketObserved: evidenceValueIsObserved("UAE", pageText, evidence.sourceUrl),
      inventedRejected: evidenceValueIsObserved("2,999", pageText, evidence.sourceUrl),
      answerMatches: answerMatchesEvidence(canonicalEvidenceAnswer(evidence), evidence),
      alteredRejected: answerMatchesEvidence("I found AED 2,999.", evidence)
    };`, sandbox);
  expect(sandbox.result).toEqual({
    amountObserved: true,
    marketObserved: true,
    inventedRejected: false,
    answerMatches: true,
    alteredRejected: false
  });
});

test("extension evidence validation rejects stale, cross-tab, and unsupported claims", async () => {
  const sourceUrl = "https://samsungtradein.ae/ae-en/";
  const snapshot = {
    revision: "r-final",
    stateFingerprint: "state-final",
    workTabId: 10,
    title: "Samsung Trade-in UAE",
    url: sourceUrl,
    visibleText: "Galaxy Z Fold8 iPhone 16 Pro 256 GB Flawless AED 1,660 Estimated trade-in value only",
    controls: []
  };
  const baseArgs = {
    tab: 1,
    revision: "r-final",
    purchased_product: "Galaxy Z Fold8",
    trade_in_product: "iPhone 16 Pro",
    storage: "256 GB",
    condition: "Flawless",
    market: "UAE",
    amount: "1,660",
    currency: "AED",
    caveat: "Estimated trade-in value only",
    source_url: sourceUrl,
    supporting_excerpts: ["AED 1,660"]
  };
  const sandbox = {
    URL,
    MAX_URL_CHARS: 2_000,
    currentTurnId: "turn-1",
    latestValidatedEvidence: null,
    expectedConditions: new Set(),
    validatedEvidenceByCondition: new Map(),
    lastSnapshotBySlot: new Map([[1, snapshot]]),
    actionReceiptsBySlot: new Map([[1, []]]),
    resolveWorkTab: async () => ({ slot: 1, tab: { id: 10, url: sourceUrl } }),
    renderEvidenceCard: () => {},
    result: null
  };
  vm.runInNewContext(`${evidenceSource}\n${recordEvidenceSource}`, sandbox);
  const invoke = (args) => sandbox.recordBrowserEvidence({
    arguments: JSON.stringify(args)
  });

  const verified = await invoke(baseArgs);
  expect(verified).toContain("[VERIFIED_BROWSER_EVIDENCE]");
  const rejectionMessage = async (promise) => {
    try {
      await promise;
      return "";
    } catch (error) {
      return String(error?.message || error);
    }
  };
  expect(await rejectionMessage(invoke({ ...baseArgs, revision: "r-stale" })))
    .toMatch(/stale/i);
  expect(await rejectionMessage(invoke({
    ...baseArgs, source_url: "https://other.example/"
  }))).toMatch(/current non-Google HTTPS work tab/i);
  expect(await rejectionMessage(invoke({ ...baseArgs, amount: "2,999" })))
    .toMatch(/amount/i);
  snapshot.workTabId = 11;
  expect(await rejectionMessage(invoke(baseArgs))).toMatch(/stale/i);
});

test("browser progress resets on changed state and accumulates on unchanged calls", () => {
  const sandbox = { result: null };
  vm.runInNewContext(`${progressSource}
    const requests = [{ id: "one", name: "browser_snapshot" }];
    result = {
      unchanged: updateBrowserProgress(requests, [
        { id: "one", content: "Progress marker: unchanged; state=a" }
      ], 19),
      changed: updateBrowserProgress(requests, [
        { id: "one", content: "Progress marker: changed; state=b" }
      ], 19),
      nonBrowser: updateBrowserProgress([
        { id: "two", name: "ask_user" }
      ], [{ id: "two", content: "2026" }], 7)
    };`, sandbox);
  expect(sandbox.result).toEqual({ unchanged: 20, changed: 0, nonBrowser: 7 });
});

test("all 2300 controls are reachable by pages without a total option cap", async ({page}) => {
  await page.evaluate(() => {
    document.body.replaceChildren();
    for (let i=0;i<2300;i++) { const button=document.createElement("button"); button.textContent=`Choice ${i}`; document.body.append(button); }
  });
  let offset=0; const names=new Set();
  do {
    const snapshot=await callAgent(page,"snapshot",{offset});
    expect(snapshot.totalControls).toBe(2300);
    snapshot.controls.forEach(control=>names.add(control.name));
    offset=snapshot.nextOffset;
  } while(offset!==null);
  expect(names.size).toBe(2300);
  expect(names.has("Choice 2299")).toBe(true);
});

test("native select pages retain a fresh select ref through the final option", async ({page}) => {
  await page.evaluate(() => {
    document.body.innerHTML='<label>Condition<select id="conditions"></select></label>';
    const select=document.getElementById("conditions");
    for(let i=0;i<240;i++) select.add(new Option(`Condition ${i}`,String(i)));
  });
  let snapshot=await callAgent(page,"snapshot");
  let ref=snapshot.controls.find(c=>c.optionCount===240).ref;
  let offset=0; const names=[];
  do {
    snapshot=await callAgent(page,"snapshot",{options_ref:ref,offset});
    names.push(...snapshot.controls.map(c=>c.name)); ref=snapshot.selectRef; offset=snapshot.nextOffset;
  } while(offset!==null);
  expect(names).toHaveLength(240);
  const plan=await callAgent(page,"selectPlan",{ref,revision:snapshot.revision,value:"Condition 239"});
  expect(plan.index).toBe(239);
});

test("actionability detects overlays instead of clicking through them", async ({page}) => {
  const snapshot=await callAgent(page,"snapshot",{query:"Origin"});
  const control=snapshot.controls[0];
  await page.evaluate(() => {const overlay=document.createElement("div");overlay.id="overlay";Object.assign(overlay.style,{position:"fixed",inset:"0",zIndex:"99999"});document.body.append(overlay);});
  const blocked=await callAgent(page,"actionability",{ref:control.ref,revision:control.revision});
  expect(blocked.receivesEvents).toBe(false);
  await page.locator("#overlay").evaluate(node=>node.remove());
  const ready=await callAgent(page,"actionability",{ref:control.ref,revision:control.revision});
  expect(ready.receivesEvents).toBe(true);
});

test("cross-origin frame discovery routes fresh refs and translates click coordinates", async ({page}) => {
  const http=require("http");
  const child=http.createServer((req,res)=>res.end('<button id="choice">Frame condition</button>'));
  await new Promise(resolve=>child.listen(0,"127.0.0.1",resolve));
  const childUrl=`http://127.0.0.1:${child.address().port}/frame`;
  const parent=http.createServer((req,res)=>res.end(`<iframe style="margin:40px" src="${childUrl}"></iframe>`));
  await new Promise(resolve=>parent.listen(0,"127.0.0.1",resolve));
  try {
    await page.goto(`http://127.0.0.1:${parent.address().port}`);
    await page.frameLocator("iframe").locator("button").waitFor();
    const runSource=extensionSource.slice(extensionSource.indexOf("async function runPageAgent("),extensionSource.indexOf("function pageAgent("));
    const sandbox={pageAgent:null,chrome:{scripting:{executeScript:async({target,func,args=[]})=>{
      const frames=page.frames(); const ids=target.allFrames?frames.map((_,index)=>index):(target.frameIds||[0]);
      return Promise.all(ids.map(async frameId=>({frameId,result:await frames[frameId].evaluate(({source,args})=>(0,eval)(`(${source})`)(...args),{source:func.toString(),args})})));
    }}}};
    vm.runInNewContext(`${pageAgentSource}\n${runSource}`,sandbox);
    const top=await sandbox.runPageAgent(1,"snapshot",{});
    expect(top.frames[0].url).toBe(childUrl);
    const nested=await sandbox.runPageAgent(1,"snapshot",{frame:1});
    const ref=nested.controls[0].ref;
    expect(ref).toMatch(/^f1@/);
    const resolved=await sandbox.runPageAgent(1,"resolve",{ref,revision:nested.revision});
    expect(resolved.x).toBeGreaterThan(40);
    const ready=await sandbox.runPageAgent(1,"actionability",{ref,revision:nested.revision});
    expect(ready.receivesEvents).toBe(true);
    await page.evaluate(()=>{const overlay=document.createElement("div");Object.assign(overlay.style,{position:"fixed",inset:"0",zIndex:"9999"});document.body.append(overlay);});
    const blocked=await sandbox.runPageAgent(1,"actionability",{ref,revision:nested.revision});
    expect(blocked.receivesEvents).toBe(false);
  } finally { await Promise.all([new Promise(resolve=>parent.close(resolve)),new Promise(resolve=>child.close(resolve))]); }
});

test("comparison coverage requires every condition page and retains prior decisions", () => {
  const helper=extensionSource.slice(extensionSource.indexOf("function recordConditionDiscovery("),extensionSource.indexOf("async function runPageAgent("));
  const comparison=extensionSource.match(/function comparisonRequested\(\) \{[^\n]+/)[0];
  const sandbox={expectedConditions:new Set(),conditionPageCoverage:new Map(),conditionScopesWithChoices:new Set(),conditionEnumerationComplete:false,
    normalizedEvidenceText:value=>value.toLowerCase(),currentRequestPrompt:"Check trade-in quotes",currentClarificationAnswers:["Compare all"]};
  vm.runInNewContext(`${helper}\n${comparison}`,sandbox);
  const control=name=>({name,role:"radio",groupLabel:"Device condition"});
  sandbox.recordConditionDiscovery({revision:"r1",offset:0,nextOffset:2,controls:[control("Flawless"),control("Good")]});
  expect(sandbox.conditionEnumerationComplete).toBe(false);
  sandbox.recordConditionDiscovery({revision:"r2",offset:2,nextOffset:null,controls:[control("Broken")]});
  expect(sandbox.conditionEnumerationComplete).toBe(true);
  expect([...sandbox.expectedConditions]).toEqual(["flawless","good","broken"]);
  expect(sandbox.comparisonRequested()).toBe(true);
});
