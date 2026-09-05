const { test, expect } = require("@playwright/test");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const extensionRoot = path.resolve(
  __dirname,
  "../../src/Scribble.BrowserExtension"
);
const source = fs.readFileSync(
  path.join(extensionRoot, "sidepanel.js"),
  "utf8"
);
const html = fs.readFileSync(
  path.join(extensionRoot, "sidepanel.html"),
  "utf8"
);
const manifest = JSON.parse(fs.readFileSync(
  path.join(extensionRoot, "manifest.json"),
  "utf8"
));

test("extension version status distinguishes current and stale installs", () => {
  const compareSource = source.slice(
    source.indexOf("function compareVersions("),
    source.indexOf("function setActivity(")
  );
  const sandbox = { result: null };
  vm.runInNewContext(`${compareSource}
    result = {
      stale: compareVersions("1.2.0", "1.2.1") < 0,
      current: compareVersions("1.2.1", "1.2.1") === 0,
      newer: compareVersions("1.3", "1.2.9") > 0
    };`, sandbox);
  expect(sandbox.result).toEqual({
    stale: true,
    current: true,
    newer: true
  });
  expect(html).toContain('id="reloadExtension"');
  expect(manifest.version).toBe("1.7.0");
});

test("browser activity stays in Pixel Pal and public aliases stay bounded", () => {
  const provenanceSource = source.slice(
    source.indexOf("function normalizedTokens("),
    source.indexOf("function isGoogleSearchControl(")
  );
  const sandbox = { result: null };
  vm.runInNewContext(`${provenanceSource}
    result = {
      dubaiAirport: isSafePublicInference(
        "Dubai International DXB",
        "Find flights to Dubai in September"
      ),
      septemberNumber: isSafePublicInference(
        "09",
        "Find flights in September"
      ),
      wrongAirport: isSafePublicInference(
        "SHJ",
        "Find flights to Dubai"
      ),
      pageOnly: isSafePublicInference(
        "private page token",
        "Find flights to Dubai"
      )
    };`, sandbox);
  expect(sandbox.result).toEqual({
    dubaiAirport: true,
    septemberNumber: true,
    wrongAirport: false,
    pageOnly: false
  });
  expect(source).not.toContain('appendMessage("audit"');
  expect(source).not.toContain('role === "audit"');
});

test("Scribble's live browser narration uses first person", () => {
  expect(source).toContain("I'm thinking about what I found");
  expect(source).toContain("I'm clicking ${label} in ${site}");
  expect(source).toContain("I'm reviewing Google’s results");
  expect(source).not.toContain("Thinking about what it found");
  expect(source).not.toContain("Clicking ${label} in ${site}");
  expect(source).not.toMatch(/Scribble (?:asks|keeps|stops|cannot)\b/);
  expect(source).not.toMatch(
    /(?:setActivity|setWorkStatus|throw new Error)\(\s*["`](?:The|That|A|An|No|Google)\b/
  );
});

test("snapshot and evidence boundaries are explicit and complete-record based", () => {
  expect(source).toContain("MAX_SNAPSHOT_CHARS = 24_000");
  expect(source).toContain("MAX_VISIBLE_TEXT_CHARS = 5_000");
  expect(source).toContain("controls omitted by snapshot budget");
  expect(source).toContain("for (const line of controlLines)");
  expect(source).toContain('runPageAgent(target.tab.id, "invalidate"');
  expect(source).toContain("openObservedHttpsLink");
  expect(source).toContain("browser_record_evidence");
  expect(source).toContain("Open evidence tab");
});

test("browser actions resolve the supplied ref before any fresh snapshot", async () => {
  const helperSource = source.slice(
    source.indexOf("async function resolveBeforeActionSnapshot("),
    source.indexOf("async function runVerifiedAction(")
  );
  const calls = [];
  const sandbox = {
    calls,
    chrome: {
      tabs: {
        get: async () => ({
          id: 7,
          title: "Google results",
          url: "https://www.google.ae/search?q=samsung"
        })
      }
    },
    runPageAgent: async (_tabId, command) => {
      calls.push(command);
      if (command === "resolve") {
        return {
          revision: "r1",
          descriptor: {
            role: "link",
            name: "Samsung trade-in",
            linkTarget: "https://samsungtradein.ae/ae-en/"
          }
        };
      }
      return { stateFingerprint: "stable-google-results" };
    },
    inspectWorkTab: async () => {
      calls.push("snapshot");
      throw new Error("A ref-scoped action must not snapshot first.");
    },
    result: null
  };
  vm.createContext(sandbox);
  await vm.runInContext(`${helperSource}
    resolveBeforeActionSnapshot(
      { tab: { id: 7, url: "https://www.google.ae/search?q=samsung" } },
      { action: "click", ref: "r1:e4" },
      null).then(value => { result = value; });`, sandbox);

  expect(calls).toEqual(["resolve", "probe"]);
  expect(sandbox.result.controls[0].linkTarget)
    .toBe("https://samsungtradein.ae/ae-en/");
});
