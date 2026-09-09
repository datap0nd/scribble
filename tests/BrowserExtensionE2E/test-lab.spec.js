const {test,expect}=require('@playwright/test');
const fs=require('fs');
const path=require('path');
const source=fs.readFileSync(path.resolve(__dirname,'../../src/Scribble.BrowserExtension/sidepanel.js'),'utf8');
const lab=source.slice(source.indexOf('async function refreshTestLab()'),source.indexOf('void refreshTestLab();'));

test('starting a Chrome case fills the exact prompt once without submitting or overwriting edits',async({page})=>{
  await page.setContent('<body><textarea id="prompt"></textarea></body>');
  await page.evaluate(code=>{
    window.PING_TIMEOUT_MS=10000;window.SETTINGS_TIMEOUT_MS=900000;window.isSending=false;
    window.elements={prompt:document.querySelector('#prompt')};window.clears=0;window.nativeCalls=[];
    window.clearChat=async()=>{clears++};
    window.nativeStatus={enabled:true,runId:'run1',host:'Chrome',prompt:'Read the operations page.',captureState:'capturing'};
    window.sendNativeMessage=async m=>{nativeCalls.push(m.type);return {ok:true,content:JSON.stringify(nativeStatus)}};
    window.eval(code);
  },lab);
  await page.evaluate(()=>refreshTestLab());
  await expect(page.locator('#prompt')).toHaveValue('Read the operations page.');
  await page.locator('#prompt').fill('Operator edit');
  await page.evaluate(()=>refreshTestLab());
  await expect(page.locator('#prompt')).toHaveValue('Operator edit');
  expect(await page.evaluate(()=>clears)).toBe(1);
  expect(await page.evaluate(()=>nativeCalls.every(t=>t==='testLabStatus'))).toBe(true);
  await page.evaluate(()=>{nativeStatus={...nativeStatus,runId:'run2',host:'Excel'};return refreshTestLab()});
  await expect(page.locator('#prompt')).toHaveValue('Operator edit');
});

test('test lab is hidden by default, reflects native capture failure, and disappears on disable',async({page})=>{
  await page.setContent('<body><p>Ordinary Scribble pane</p></body>');
  await page.evaluate(code=>{
    window.PING_TIMEOUT_MS=10000;window.SETTINGS_TIMEOUT_MS=900000;
    window.nativeStatus={enabled:false};window.nativeCalls=[];
    window.sendNativeMessage=async message=>{nativeCalls.push(message.type);return {ok:true,content:JSON.stringify(nativeStatus)}};
    window.eval(code);
  },lab);
  await page.evaluate(()=>refreshTestLab());
  await expect(page.locator('#testLabButton')).toHaveCount(0);
  await page.evaluate(()=>{nativeStatus={enabled:true,runId:'test',captureState:'capturing'};return refreshTestLab()});
  await expect(page.locator('#testLabButton')).toHaveText('Test Lab • capturing');
  await page.locator('#testLabButton').click();
  expect(await page.evaluate(()=>nativeCalls)).toContain('openTestLab');
  await page.evaluate(()=>{nativeStatus.captureState='incomplete';return refreshTestLab()});
  await expect(page.locator('#testLabButton')).toHaveText('Test Lab • incomplete');
  await page.evaluate(()=>{nativeStatus={enabled:false};return refreshTestLab()});
  await expect(page.locator('#testLabButton')).toBeHidden();
});

test('lost native connection hides the operator entry rather than retaining stale enabled state',async({page})=>{
  await page.setContent('<body><button id="testLabButton">Test Lab</button></body>');
  await page.evaluate(code=>{window.PING_TIMEOUT_MS=10000;window.sendNativeMessage=async()=>{throw new Error('Disconnected')};window.eval(code)},lab);
  await page.evaluate(()=>refreshTestLab());
  await expect(page.locator('#testLabButton')).toBeHidden();
});
