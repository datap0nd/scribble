const {test,expect}=require('@playwright/test');
const fs=require('fs');
const path=require('path');
const source=fs.readFileSync(path.resolve(__dirname,'../../src/Scribble.BrowserExtension/sidepanel.js'),'utf8');
const lab=source.slice(source.indexOf('async function refreshTestLab()'),source.indexOf('void refreshTestLab();'));

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
  expect(await page.evaluate(()=>nativeCalls.at(-1))).toBe('openTestLab');
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
