import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const source = fs.readFileSync(new URL('../taskpane.js', import.meta.url), 'utf8');
const clone = (value) => JSON.parse(JSON.stringify(value));
function pane({ rows = 1, columns = 1 } = {}) {
  const messages = [];
  const loads = [];
  const originalFormat = { fill: 'yellow', font: 'Calibri', border: 'double', alignment: 'right', locked: false, columnWidth: 18 };
  let formulas = [[10]], numberFormat = [['0.00']];
  let syncCount = 0;
  let failSync = 0;
  const sheet = { id: 'sheet-original', name: 'Sheet1', load() {}, getRange: () => range,
    getUsedRangeOrNullObject: () => ({ isNullObject: true, load() {} }) };
  const range = {
    address: 'Sheet1!A1', rowCount: rows, columnCount: columns, worksheet: sheet,
    format: clone(originalFormat), load(fields) { loads.push({ fields: [...fields], rows: this.rowCount, columns: this.columnCount }); },
    get formulas() { return clone(formulas); }, set formulas(value) { formulas = clone(value); },
    get values() { return clone(formulas); }, set values(value) { formulas = clone(value); },
    get numberFormat() { return clone(numberFormat); }, set numberFormat(value) { numberFormat = clone(value); },
    getCell() { return this; },
    getResizedRange(r, c) {
      if (rows === 1 && columns === 1) return this;
      return { rowCount: r + 1, columnCount: c + 1, values: [['sample']], formulas: [['sample']],
        load(fields) { loads.push({ fields: [...fields], rows: r + 1, columns: c + 1 }); } };
    },
    clear() { throw new Error('Undo must not clear existing formats'); },
  };
  const sheets = { items: [sheet], load() {}, getItem: () => sheet, getActiveWorksheet: () => sheet };
  const context = { workbook: { worksheets: sheets, getSelectedRange: () => range },
    sync: async () => { syncCount++; if (syncCount === failSync) throw new Error('simulated Office failure'); } };
  const sandbox = {
    window: { location: { origin: 'https://localhost:8788' } },
    document: { querySelector: () => null, getElementById: () => ({}) },
    Office: { onReady() {} },
    Excel: { run: async (fn) => fn(context) },
    crypto: { getRandomValues: (bytes) => bytes.fill(1) }, Uint8Array,
  };
  vm.createContext(sandbox); vm.runInContext(source, sandbox);
  vm.runInContext('addMessage = (...args) => messages.push(args); saveChatHistory = () => {}; setStatus = () => {}; state.workbookId = "book-1";', Object.assign(sandbox, { messages }));
  return { sandbox, range, sheet, loads, originalFormat,
    state: vm.runInContext('state', sandbox),
    failNextSync() { failSync = syncCount + 1; } };
}
const action = { type: 'write_cells', start_cell: 'Sheet1!A1', values: [[20]], auto_format: true };

test('immediate write and Undo preserve all pre-existing formatting', async () => {
  const p = pane();
  await p.sandbox.writeCellsAction(action);
  assert.deepEqual(p.range.values, [[20]]);
  assert.deepEqual(p.range.format, p.originalFormat);
  await p.sandbox.undoLast();
  assert.deepEqual(p.range.formulas, [[10]]);
  assert.deepEqual(p.range.numberFormat, [['0.00']]);
  assert.deepEqual(p.range.format, p.originalFormat);
  assert.equal(p.state.undoStack.length, 0);
});

for (const scenario of ['later edit', 'replacement sheet', 'number-format edit']) {
  test(`immediate Undo refuses ${scenario} and retains the record`, async () => {
    const p = pane();
    await p.sandbox.writeCellsAction(action);
    if (scenario === 'later edit') p.range.formulas = [[999]];
    if (scenario === 'replacement sheet') p.sheet.id = 'replacement';
    if (scenario === 'number-format edit') p.range.numberFormat = [['0%']];
    const before = p.range.formulas;
    await p.sandbox.undoLast();
    assert.deepEqual(p.range.formulas, before);
    assert.equal(p.state.undoStack.length, 1);
  });
}

test('failed Undo retains a retryable record and releases the busy flag', async () => {
  const p = pane(); await p.sandbox.writeCellsAction(action);
  p.failNextSync(); await p.sandbox.undoLast();
  assert.equal(p.state.undoStack.length, 1);
  assert.equal(p.state.undoing, false);
  assert.deepEqual(p.range.formulas, [[20]]);
});

test('unverified write does not allow Undo to cascade into an earlier record', async () => {
  const p = pane(); await p.sandbox.writeCellsAction(action);
  // Fail after the before-image read, during the commit sync.
  const originalRun = p.sandbox.Excel.run;
  p.sandbox.Excel.run = fn => originalRun(async context => {
    let calls = 0; const sync = context.sync;
    context.sync = async () => { if (++calls === 2) throw new Error('commit failed'); await sync(); };
    return fn(context);
  });
  await assert.rejects(p.sandbox.writeCellsAction(action), /commit failed/);
  assert.equal(p.state.undoStack.at(-1).kind, 'non_undoable');
});

test('reviewed write and Undo preserve formatting', async () => {
  const p = pane();
  const proposal = await p.sandbox.bindProposalToWorkbook({ actions: [action] });
  await p.sandbox.applyReviewedProposal(proposal);
  assert.deepEqual(p.range.formulas, [[20]]);
  assert.deepEqual(p.range.format, p.originalFormat);
  await p.sandbox.undoLast();
  assert.deepEqual(p.range.formulas, [[10]]);
  assert.deepEqual(p.range.format, p.originalFormat);
});

test('reviewed Undo refuses subsequent content changes', async () => {
  const p = pane();
  await p.sandbox.applyReviewedProposal(await p.sandbox.bindProposalToWorkbook({ actions: [action] }));
  p.range.formulas = [[999]];
  await p.sandbox.undoLast();
  assert.deepEqual(p.range.formulas, [[999]]);
  assert.equal(p.state.undoStack.length, 1);
});

for (const [rows, columns] of [[1048576, 1], [1, 16384], [1048576, 16384]]) {
  test(`selection ${rows}x${columns} samples at most 100x16 cells`, async () => {
    const p = pane({ rows, columns });
    await p.sandbox.readWorkbookContext();
    for (const load of p.loads.filter(load => load.fields.includes('values') || load.fields.includes('formulas'))) {
      assert.ok(load.rows <= 100 && load.columns <= 16);
    }
    assert.equal(p.state.selection.rowCount, rows);
    assert.equal(p.state.selection.columnCount, columns);
    assert.equal(p.state.selection.truncated, true);
    assert.deepEqual(clone(p.state.selection.values), [['sample']]);
  });
}

test('small selection retains its data without truncation', async () => {
  const p = pane(); await p.sandbox.readWorkbookContext();
  assert.equal(p.state.selection.truncated, false);
  assert.deepEqual(clone(p.state.selection.values), [[10]]);
});
