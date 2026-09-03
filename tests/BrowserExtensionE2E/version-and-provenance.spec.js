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
