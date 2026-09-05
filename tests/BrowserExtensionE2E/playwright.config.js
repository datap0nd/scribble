const { defineConfig } = require("@playwright/test");

module.exports = defineConfig({
  testDir: ".",
  testMatch: "*.spec.js",
  timeout: 30_000,
  fullyParallel: false,
  use: {
    trace: "retain-on-failure",
    screenshot: "only-on-failure",
    headless: true,
    viewport: { width: 1000, height: 800 }
  }
});
