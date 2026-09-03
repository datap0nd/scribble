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
    currentRequestPrompt: "Scrape flight prices from Germany to Dubai, September. Use example.test",
    currentClarificationAnswers: ["2026"],
    URL,
    result: null
  };
  vm.runInNewContext(`${provenanceSource}
    result = {
      ordered: isOrderedUserTokenSubset("flight prices Germany Dubai September 2026", approvedSourceText()),
      reordered: isOrderedUserTokenSubset("Dubai Germany flights", approvedSourceText()),
      suppliedUrl: urlWasUserProvided("https://example.test/"),
      inventedUrl: urlWasUserProvided("https://provider.invalid/")
    };`, sandbox);
  expect(sandbox.result).toEqual({
    ordered: true,
    reordered: false,
    suppliedUrl: true,
    inventedUrl: false
  });
});
