import fs from 'node:fs';
import assert from 'node:assert/strict';
import { chromium, expect } from '@playwright/test';
const directory = '.runtime/native-test';
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
async function excel(action, fields = {}) {
  const id = crypto.randomUUID();
  fs.writeFileSync(`${directory}/command.tmp`, JSON.stringify({ id, action, ...fields }));
  for (let attempt = 0; ; attempt++) {
    try { fs.renameSync(`${directory}/command.tmp`, `${directory}/command.json`); break; }
    catch (error) { if (!['EPERM', 'EBUSY'].includes(error.code) || attempt >= 40) throw error; await delay(10); }
  }
  for (let attempt = 0; attempt < 150; attempt++) {
    await delay(100);
    try {
      const response = JSON.parse(fs.readFileSync(`${directory}/response.json`, 'utf8').replace(/^\uFEFF/, ''));
      if (response.id === id) { if (response.error) throw new Error(response.error); return response.result; }
    } catch (error) { if (!['ENOENT', 'EBUSY', 'EPERM'].includes(error.code) && !(error instanceof SyntaxError)) throw error; }
  }
  throw new Error('Excel harness command timed out: ' + action);
}
const read = address => excel('get', { sheet: 'Model', address });
const set = (address, value) => excel('set', { sheet: 'Model', address, value });
const browser = await chromium.connectOverCDP('http://127.0.0.1:9223');
const page = browser.contexts()[0].pages().find(p => p.url().startsWith('http://127.0.0.1:17843'));
assert(page, 'Native WebView2 research page must exist');
const checks = [], errors = [];
let failure;
const measurements = {};
page.on('pageerror', error => errors.push(error.message));
async function check(name, action) { const start = performance.now(); await action(); checks.push({ name, status: 'PASS', elapsedMs: Math.round(performance.now() - start) }); console.log('PASS ' + name); }
async function preview() {
  await page.locator('#preview-refresh').click();
  await expect(page.locator('#preview-refresh')).toBeEnabled({ timeout: 35000 });
  await expect(page.locator('#apply-refresh')).toBeVisible({ timeout: 35000 });
}
async function apply() { await page.locator('#apply-refresh').click(); await expect(page.locator('#notice')).toContainText('reported values updated', { timeout: 35000 }); }
try {
  await check('real WebView2 connects to Excel and shows sourced data', async () => {
    await expect(page.locator('#host-state')).toHaveText('Connected to Excel');
    await expect(page.locator('#revenue-value')).toHaveText('$281.7B');
    await page.screenshot({ path: 'artifacts/validation/native-overview.png', fullPage: true });
  });
  await check('preview and apply preserve formulas and analyst inputs', async () => {
    const formula = (await read('F10')).formula;
    await page.locator('[data-tab="refresh"]').click(); await preview();
    await expect(page.locator('#refresh-plan')).toContainText('9 reported values ready');
    await page.screenshot({ path: 'artifacts/validation/native-refresh.png', fullPage: true });
    await apply();
    assert.equal((await read('E10')).value, 281724); assert.equal((await read('F10')).formula, formula); assert.equal((await read('F5')).value, .12);
    assert(Math.abs((await read('F10')).value - 315530.88) < .00001);
    assert.equal((await excel('get', { sheet: 'Sources', address: 'K92' })).value, 'MSFT-Revenue-FY2025');
    assert.equal((await excel('get', { sheet: 'Sources', address: 'M92' })).value, 281724000000);
    assert.match((await excel('get', { sheet: 'Sources', address: 'L92' })).value, /^https:\/\/www\.sec\.gov\/Archives\//);
  });
  await check('rollback restores historical values and preserves later input edits', async () => {
    await set('F5', .2); await page.locator('#rollback-refresh').click();
    await expect(page.locator('#notice')).toContainText('9 values restored', { timeout: 35000 });
    assert.equal((await read('E10')).value, null); assert.equal((await read('F5')).value, .2); assert.equal((await read('F10')).hasFormula, true);
    assert.equal((await excel('get', { sheet: 'Sources', address: 'K92' })).value, null);
    await set('F5', .12);
  });
  await check('edits after preview abort the entire transaction', async () => {
    await preview(); await set('E10', 123); await page.locator('#apply-refresh').click();
    await expect(page.locator('#notice')).toContainText('Cells edited since preview: Model!E10', { timeout: 35000 });
    assert.equal((await read('E10')).value, 123); assert.equal((await read('E11')).value, null); await set('E10', null);
  });
  await check('changing the company after preview is a conflict', async () => {
    await preview(); await set('B4', 'AAPL'); await page.locator('#apply-refresh').click();
    await expect(page.locator('#notice')).toContainText('Model!B4', { timeout: 35000 });
    assert.equal((await read('E10')).value, null); await set('B4', 'MSFT');
  });
  await check('company mismatch pauses forecasts until a new reviewed import', async () => {
    await preview(); await apply(); await set('B4', 'AAPL');
    assert.equal((await read('F10')).value, ''); assert.match((await read('A21')).value, /Review required/);
    await set('B4', 'MSFT'); assert(Math.abs((await read('F10')).value - 315530.88) < .00001);
  });
  await check('rollback refuses to overwrite a later historical edit', async () => {
    await set('E10', 999); await page.locator('#rollback-refresh').click();
    await expect(page.locator('#notice')).toContainText('Undo would overwrite later edits', { timeout: 35000 });
    assert.equal((await read('E10')).value, 999); assert.equal((await read('E11')).value, 193893);
    await set('E10', 281724);
  });
  await check('formula insertion uses the selected empty cell and refuses overwrite', async () => {
    await excel('select', { sheet: 'Model', address: 'I10' }); await page.locator('[data-tab="overview"]').click();
    await page.locator('#insert-formula').click(); await expect(page.locator('#notice')).toContainText('Formula inserted', { timeout: 35000 });
    assert.match((await read('I10')).formula, /LL.METRIC/);
    let inserted;
    for(let i=0;i<60;i++){inserted=(await read('I10')).value;if(inserted===281724)break;await delay(100);await excel('calculate')}
    assert.equal(inserted,281724,'Inserted asynchronous formula must resolve before the workbook is saved');
    await page.locator('#insert-formula').click(); await expect(page.locator('#notice')).toContainText('Existing content is never overwritten', { timeout: 35000 });
  });
  await check('calculated research saves claims and citations into a new sheet', async () => {
    await page.locator('[data-tab="research"]').click(); await page.locator('#use-ai').uncheck(); await page.locator('#run-research').click();
    await expect(page.locator('.answer-card')).toBeVisible(); await page.getByRole('button', { name: 'Save research to workbook' }).click();
    await expect(page.locator('#notice')).toContainText('saved to a new worksheet', { timeout: 35000 });
    assert.equal((await excel('status')).sheetCount, 6);
    await page.screenshot({ path: 'artifacts/validation/native-research.png', fullPage: true });
  });
  await check('Excel application settings are restored', async () => {
    const status = await excel('status'); assert.equal(status.updating, true); assert.equal(status.events, true); assert.equal(status.calculation, -4105); assert.equal(status.ready, true);
    assert.deepEqual(errors, []); measurements.workbook = await excel('save');
  });
  await check('malformed formula requests show an actionable status', async()=>{
    await excel('set',{sheet:'Model',address:'J4',formula:'=LL.STATUS("MSFT","INVALID_METRIC","FY2025")'});
    let value;
    for(let i=0;i<50;i++){value=(await read('J4')).value;if(typeof value==='string'&&value.startsWith('Unavailable:'))break;await delay(100);await excel('calculate')}
    assert.match(value,/Unknown metric.*Revenue/);await set('J4',null);
  });
  const endpoint = JSON.parse(fs.readFileSync('.runtime/endpoint.json', 'utf8').replace(/^\uFEFF/, ''));
  async function service(path, body) {
    const response = await fetch(endpoint.BaseUrl + '/api' + path, { method: body ? 'POST' : 'GET', headers: { Authorization: 'Bearer ' + endpoint.Token, 'Content-Type': 'application/json' }, body: body ? JSON.stringify(body) : undefined });
    assert.equal(response.status, 200); return response.json();
  }
  await check('2000 repeated asynchronous formulas share one provider load', async () => {
    await service('/connection', { mode: 'slow' });
    const before = await service('/diagnostics');
    const start = performance.now();
    measurements.bulk = await excel('bulkFormula', { formula: '=LL.METRIC("MSFT","Revenue","FY2025",314159)' });
    const responsiveStart = performance.now();
    assert.equal((await excel('status')).ready, true);
    measurements.responsivenessMs = performance.now() - responsiveStart;
    let last;
    for (let attempt = 0; attempt < 100; attempt++) {
      last = await excel('get', { sheet: 'Performance', address: 'A2000' });
      if (last.value === 281724) break;
      await delay(100); await excel('calculate');
    }
    assert.equal(last.value, 281724); assert.equal((await excel('get', { sheet: 'Performance', address: 'A1' })).value, 281724);
    measurements.allCellsResolvedMs = performance.now() - start;
    const after = await service('/diagnostics');
    measurements.providerLoads = after.providerCalls - before.providerCalls;
    assert.equal(measurements.providerLoads, 1); assert(measurements.responsivenessMs < 1000);
    await service('/connection', { mode: 'online' });
  });
  await check('repeated workbook and pane lifecycle remains functional', async () => {
    measurements.lifecycle = [];
    for (let cycle = 0; cycle < 4; cycle++) {
      await excel('cycle'); await delay(1500);
      const nextPage = browser.contexts()[0].pages().find(p => p.url().startsWith(endpoint.BaseUrl));
      assert(nextPage); await expect(nextPage.locator('#host-state')).toHaveText('Connected to Excel');
      const sample = await excel('memory'); const health = await service('/diagnostics');
      measurements.lifecycle.push({ cycle: cycle + 1, ...sample, streamSubscribers: health.streamSubscribers });
    }
  });
  await check('closing a workbook cancels its pending formula waiter', async () => {
    await service('/connection', { mode: 'slow', delayMs: 5000 });
    const before = await excel('cancellations');
    await excel('openPending'); await delay(300); await excel('closePending');
    let after;
    for (let attempt = 0; attempt < 30; attempt++) { after = await excel('cancellations'); if (after > before) break; await delay(100); }
    assert(after > before, 'Excel-DNA should signal cancellation when the requesting workbook closes');
    measurements.canceledWaiters = after - before;
    await service('/connection', { mode: 'online' });
  });
} catch(error) { failure=error.message; throw error; }
finally {
  fs.writeFileSync('artifacts/validation/native-results.json', JSON.stringify({ checkedAt: new Date().toISOString(), status:failure?'FAIL':'PASS', failure, checks, errors, measurements }, null, 2));
  const current=JSON.parse(fs.readFileSync('.runtime/endpoint.json','utf8').replace(/^\uFEFF/,''));
  await fetch(current.BaseUrl+'/api/connection',{method:'POST',headers:{Authorization:'Bearer '+current.Token,'Content-Type':'application/json'},body:JSON.stringify({mode:'online'})}).catch(()=>{});
  await browser.close();
}
