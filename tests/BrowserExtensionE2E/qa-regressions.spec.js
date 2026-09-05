const { test, expect } = require('@playwright/test');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

test('inspection failures distinguish restrictions from navigation and unknown failures', () => {
  const source = fs.readFileSync(path.resolve(__dirname, '../../src/Scribble.BrowserExtension/sidepanel.js'), 'utf8');
  const helper = source.slice(source.indexOf('function inspectionFailure('), source.indexOf('async function inspectWorkTab('));
  const context = { URL };
  vm.createContext(context);
  vm.runInContext(helper, context);
  const classify = (message, url = 'https://example.test/private?token=secret') => context.inspectionFailure({ message }, url);
  expect(classify('Execution context was destroyed').retry).toBe(true);
  expect(classify('Cannot access contents of url').message).toContain('BROWSER_ACCESS_DENIED');
  expect(classify('Unknown failure', 'chrome://settings').message).toContain('BROWSER_RESTRICTED_SCHEME');
  expect(classify('No tab with id 2').message).toContain('BROWSER_TAB_UNAVAILABLE');
  expect(classify('Unexpected error secret').message).toContain('BROWSER_INSPECTION_FAILED');
  expect(classify('Unexpected error secret').message).not.toMatch(/private|token|secret/);
});

test('narrow Office tables expose every labeled cell and ordered lists preserve starts', async ({ page }) => {
  await page.setViewportSize({ width: 340, height: 900 });
  await page.addInitScript(() => {});
  const html = fs.readFileSync(path.resolve(__dirname, '../../src/Scribble/UI/ChatPaneWeb.html'), 'utf8');
  await page.setContent(html.replace('<script>', `<script>window.chrome = { webview: { postMessage() {}, addEventListener(name, handler) { window.deliver = handler; } } };`));
  await page.evaluate(() => window.deliver({ data: { type: 'assistant', md:
    '| Market | Filename | Rows | Columns | Issue | Source |\n|---|---|---|---|---|---|\n| JO | JOStockInformation.xls | 12 | 3 | Missing description | Product Extract SELV-JO |\n| EG | EGStockInformation.xls | 8 | 3 | Duplicate SKU | Product Extract SEEG |\n\n1. First option\n\n2. Second option' } }));
  await expect(page.locator('.md td')).toHaveCount(12);
  await expect(page.locator('.md ol').nth(1)).toHaveAttribute('start', '2');
  const layout = await page.locator('.md .tblwrap').evaluate(node => ({ width: node.clientWidth, scroll: node.scrollWidth,
    labels: [...node.querySelectorAll('td')].map(cell => getComputedStyle(cell, '::before').content) }));
  expect(layout.scroll).toBeLessThanOrEqual(layout.width + 1);
  expect(layout.labels[5]).toContain('Source');
  await page.setViewportSize({ width: 900, height: 900 });
  await expect(page.locator('.md table')).toHaveCSS('display', 'table');
});

test('production inspector retries a destroyed document once and stops other failures', async () => {
  const source = fs.readFileSync(path.resolve(__dirname, '../../src/Scribble.BrowserExtension/sidepanel.js'), 'utf8');
  const inspector = source.slice(source.indexOf('function inspectionFailure('), source.indexOf('async function settleAndInspectWorkTab('));
  for (const transient of [true, false]) {
    let calls = 0;
    const context = { URL, snapshotImagesByTab: new Map(), stopRequested: false, operatorDetachError: '',
      boundText: text => text, delay: async () => {}, chrome: { tabs: { get: async () => ({ url: 'https://example.test/' }) } },
      runPageAgent: async () => { calls++; if (calls === 1 || !transient) throw new Error(transient ? 'Execution context destroyed' : 'Cannot access contents of url');
        return { title: 'Product', visibleText: 'Available models', controls: [] }; } };
    vm.createContext(context);
    vm.runInContext(inspector, context);
    if (transient) { await context.inspectWorkTab(1, '', 0, 1); expect(calls).toBe(2); expect(context.stopRequested).toBe(false); }
    else {
      let failure = '';
      try { await context.inspectWorkTab(1, '', 0, 1); } catch (error) { failure = error.message; }
      expect(failure).toContain('BROWSER_ACCESS_DENIED'); expect(calls).toBe(1); expect(context.stopRequested).toBe(true);
    }
  }
});
