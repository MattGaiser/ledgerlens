import test from 'node:test';
import assert from 'node:assert/strict';
import { createOfficeBridge } from '../web/office-bridge.js';
function host({ value = '', formula = '', merged = false, protectedSheet = false, readOnly = false } = {}) {
  const cell = { values: [[value]], formulas: [[formula]], rowCount: 1, columnCount: 1, address: 'Sheet1!A1', load() {}, getMergedAreasOrNullObject: () => ({ isNullObject: !merged }), worksheet: { protection: { protected: protectedSheet, load() {} } } };
  const range = { format: { fill: {}, font: {}, autofitRows() {} } };
  const sheet = { getRangeByIndexes: () => range, getRange: () => range, activate() {} };
  let syncs = 0, calls = 0, created = 0;
  const context = { workbook: { getSelectedRange: () => cell, worksheets: { add() { created++; return sheet; } } }, async sync() { syncs++; } };
  const office = { context: { document: { mode: readOnly ? 'readOnly' : 'readWrite' } }, DocumentMode: { ReadOnly: 'readOnly' } };
  const excel = { run: callback => callback(context) };
  const bridge = createOfficeBridge(office, excel, async () => { calls++; return { fact: { value: 281724, metric: 'Revenue', sourceId: 'MSFT-Revenue-FY2025' } }; });
  return { bridge, cell, range, stats: () => ({ syncs, calls, created }) };
}
const payload = { ticker: 'MSFT', metric: 'Revenue', period: 'FY2025' };
test('Office.js inserts a fetched number only into an empty cell', async () => {
  const h = host(); const result = await h.bridge.execute('insertFormula', payload);
  assert.equal(result.mode, 'snapshot'); assert.equal(h.cell.values[0][0], 281724); assert.equal(h.stats().syncs, 2);
});
for (const [name, settings] of [['existing value', { value: 5 }], ['formula', { formula: '=1+1' }], ['merged range', { merged: true }], ['protected sheet', { protectedSheet: true }]]) {
  test('Office.js rejects ' + name + ' without writing', async () => {
    const h = host(settings); await assert.rejects(h.bridge.execute('insertFormula', payload)); assert.deepEqual(h.cell.values, [[settings.value ?? '']]); assert.equal(h.stats().syncs, 1);
  });
}
test('read-only Office document fails before fetching', async () => {
  const h = host({ readOnly: true }); await assert.rejects(h.bridge.execute('insertFormula', payload), /read-only/); assert.equal(h.stats().calls, 0);
});
test('native-only Office commands are explicitly unsupported', async () => {
  await assert.rejects(host().bridge.execute('applyRefresh', {}), /Windows native/);
});
test('Office research export escapes formula-like text', async () => {
  const h = host();
  await h.bridge.execute('saveResearch', { answer: { headline: '=HYPERLINK("untrusted")', summary: '+1+1', provider: 'Fixture', model: 'Fixture', claims: [{ text: '@example', sourceIds: ['evidence'] }], caveats: ['Historical'], sources: [{ sourceId: 'evidence', sourceUrl: 'https://www.sec.gov/Archives/edgar/data/1/index.html' }] } });
  assert.equal(h.stats().created, 1); assert(h.range.values.every(row => row.every(value => value.startsWith("'")))); assert.equal(h.range.numberFormat[0][0], '@');
});
test('Office research export rejects missing citations before writing', async () => {
  const h = host(); await assert.rejects(h.bridge.execute('saveResearch', { answer: { claims: [{ text: 'Claim', sourceIds: ['missing'] }], sources: [] } }), /citation/); assert.equal(h.stats().created, 0);
});
