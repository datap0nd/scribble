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
