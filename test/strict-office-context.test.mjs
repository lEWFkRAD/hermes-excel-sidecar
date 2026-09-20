import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const source = fs.readFileSync(new URL('../taskpane.js', import.meta.url), 'utf8');

function strictRange({ address = 'Sheet1!A1', values = [[42]], formulas = [['=42']], rowCount = 1, columnCount = 1, worksheet, pendingLoads }) {
  const loaded = new Set();
  const need = (field) => {
    if (!loaded.has(field)) throw new Error(`Office mock: ${field} was read before load() + sync()`);
  };
  return {
    load(fields) {
      const fieldsToLoad = Array.isArray(fields) ? fields : String(fields).split(',');
      pendingLoads.push(() => fieldsToLoad.forEach((field) => loaded.add(field.trim())));
    },
    get address() { need('address'); return address; },
    get values() { need('values'); return values; },
    get formulas() { need('formulas'); return formulas; },
    get rowCount() { need('rowCount'); return rowCount; },
    get columnCount() { need('columnCount'); return columnCount; },
    worksheet,
    getCell() { return this; },
    getResizedRange() { return this; },
  };
}

function strictPane() {
  const pendingLoads = [];
  const sheet = { id: 'sheet-1', name: 'Sheet1', load(fields) { this.loaded = fields; } };
  const selected = strictRange({ worksheet: sheet, pendingLoads });
  const used = strictRange({ address: 'Sheet1!A1:A1', worksheet: sheet, pendingLoads });
  const sheets = {
    items: [sheet],
    load() {},
    getActiveWorksheet() { return sheet; },
  };
  sheet.getUsedRangeOrNullObject = () => used;
  const context = {
    workbook: { worksheets: sheets, getSelectedRange: () => selected },
    sync: async () => { while (pendingLoads.length) pendingLoads.shift()(); },
  };
  const sandbox = {
    window: { location: { origin: 'https://localhost:8788' } },
    document: { querySelector: () => null, getElementById: () => ({}) },
    Office: { onReady() {} },
    Excel: { run: async (fn) => fn(context) },
    crypto: { getRandomValues: (bytes) => bytes.fill(1) }, Uint8Array, Blob,
  };
  vm.createContext(sandbox);
  vm.runInContext(source, sandbox);
  vm.runInContext('state.workbookId = "book-strict"; setStatus = () => {};', sandbox);
  return sandbox;
}

test('real pane context reader loads Office.js properties before reading them', async () => {
  const sandbox = strictPane();
  const result = await sandbox.readWorkbookContext();
  assert.equal(result.selection.address, 'Sheet1!A1');
  assert.deepEqual(JSON.parse(JSON.stringify(result.selection.values)), [[42]]);
  assert.equal(result.workbook.activeSheet, 'Sheet1');
});
