import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { normalizeWorkbook } from '../broker/server.mjs';

const source = fs.readFileSync(new URL('../taskpane.js', import.meta.url), 'utf8');
const clone = (value) => JSON.parse(JSON.stringify(value));
function pane({ rows = 1, columns = 1, api = 7 } = {}) {
  const messages = [];
  const elements = new Map();
  const element = id => {
    if (!elements.has(id)) elements.set(id, { value: "", events: {}, addEventListener(name, fn) { this.events[name] = fn; } });
    return elements.get(id);
  };
  const loads = [];
  const originalFormat = { fill: 'yellow', font: 'Calibri', border: 'double', alignment: 'right', locked: false, columnWidth: 18 };
  let formulas = [[10]], numberFormat = [['0.00']];
  let syncCount = 0;
  let failSync = 0;
  const sheet = { id: 'sheet-original', name: 'Sheet1', load() {}, getRange: () => range,
    getUsedRangeOrNullObject: () => ({ isNullObject: true, load() {} }),
    getUsedRange: () => ({ address: 'Sheet1!A1', rowCount: 1, columnCount: 1, load() {} }) };
  const range = {
    address: 'Sheet1!A1', rowCount: rows, columnCount: columns, worksheet: sheet,
    format: clone(originalFormat), load(fields) { loads.push({ fields: [...fields], rows: this.rowCount, columns: this.columnCount }); },
    get formulas() { return clone(formulas); }, set formulas(value) { formulas = clone(value); },
    get values() { return clone(formulas); }, set values(value) { formulas = clone(value); },
    get numberFormat() { return clone(numberFormat); }, set numberFormat(value) { numberFormat = clone(value); },
    getCell(r, c) {
      if (r === 0 && c === 0) return this;
      return { endRow: r, endColumn: c };
    },
    getBoundingRect(end) {
      if (rows === 1 && columns === 1) return this;
      const r = end.endRow || 0, c = end.endColumn || 0;
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
    document: { querySelector: () => null, getElementById: element },
    Office: { onReady() {}, context: { requirements: { isSetSupported: (name, version) => name === 'ExcelApi' && Number(version.split('.')[1]) <= api } } },
    Excel: { run: async (fn) => fn(context) },
    crypto: { getRandomValues: (bytes) => bytes.fill(1) }, Uint8Array, Blob, AbortController, setTimeout, clearTimeout,
  };
  vm.createContext(sandbox); vm.runInContext(source, sandbox);
  const persistHistory = sandbox.saveChatHistory;
  vm.runInContext('addMessage = (...args) => messages.push(args); saveChatHistory = () => {}; setStatus = () => {}; state.workbookId = "book-1";', Object.assign(sandbox, { messages }));
  return { sandbox, range, sheet, loads, originalFormat, persistHistory, elements, messages,
    state: vm.runInContext('state', sandbox),
    failNextSync() { failSync = syncCount + 1; } };
}
const action = { type: 'write_cells', start_cell: 'Sheet1!A1', values: [[20]], auto_format: true };

test('oversized attachments are rejected before reading a workbook or encoding files', async () => {
  const p = pane();
  await assert.rejects(p.sandbox.askHermes('test', [{ size: 101 * 1024 * 1024 }]), /100 MB per-file/);
  assert.equal(p.loads.length, 0);
});

test('attachment batches are bounded while ordinary files remain accepted', () => {
  const p = pane();
  const file = (mb) => ({ size: mb * 1024 * 1024 });
  assert.throws(() => p.sandbox.validateAttachmentSizes([file(80), file(80)]), /150 MB total/);
  assert.throws(() => p.sandbox.validateAttachmentSizes(Array(13).fill(file(0))), /12 files/);
  assert.doesNotThrow(() => p.sandbox.validateAttachmentSizes([file(100), file(50)]));
  assert.doesNotThrow(() => p.sandbox.validateAttachmentSizes([file(29), file(71), file(29)]));
});

test('HTTP upload rejection stays actionable and is not retried as a connection error', async () => {
  const p = pane();
  let calls = 0;
  p.sandbox.fetch = async () => { calls++; return { ok: false, status: 413, json: async () => ({ error: 'Batch exceeds upload limit' }) }; };
  await assert.rejects(p.sandbox.postChat({}), /Batch exceeds upload limit/);
  assert.equal(calls, 1);
});

test('native file encoding preserves base64 payload without per-byte concatenation', async () => {
  const p = pane();
  p.sandbox.FileReader = class { readAsDataURL() { this.result = 'data:application/pdf;base64,AAEC/w=='; this.onload(); } };
  const result = await p.sandbox.fileToPayload({ name: 'synthetic.pdf', size: 4, type: 'application/pdf' });
  assert.equal(result.base64, 'AAEC/w==');
  assert.equal(result.size, 4);
});

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

for (const api of [1, 2, 4, 6, 7, 12, 20]) {
  test(`ExcelApi 1.${api}: basic read, direct write and Undo work`, async () => {
    const p = pane({ api });
    if (api < 4) p.sheet.getUsedRangeOrNullObject = () => { throw new Error('1.4 API used on old host'); };
    const context = await p.sandbox.readWorkbookContext();
    assert.ok(context.workbook.excelApi.includes('1.1'));
    assert.equal(context.workbook.excelApi.includes('1.7'), api >= 7);
    await p.sandbox.writeCellsAction(action);
    assert.deepEqual(p.range.formulas, [[20]]);
    await p.sandbox.undoLast();
    assert.deepEqual(p.range.formulas, [[10]]);
  });
}

for (const [type, minimum] of [
  ['merge_cells', 2], ['unmerge_cells', 2], ['sort_range', 2], ['autofit', 2],
  ['set_column_width', 2], ['set_row_height', 2], ['conditional_format', 6],
  ['freeze_panes', 7], ['unfreeze_panes', 7],
]) {
  test(`${type} rejects the whole response before writes on an older host`, async () => {
    const p = pane({ api: minimum - 1 });
    const result = { actions: [action, { type }] };
    await assert.rejects(p.sandbox.runWorkbookActions(result), /No changes were applied/);
    assert.equal(p.loads.length, 0);
    assert.deepEqual(p.range.formulas, [[10]]);
    const supported = pane({ api: minimum });
    assert.doesNotThrow(() => supported.sandbox.assertCompatibleActions(result.actions));
  });
}

test('unknown capability API fails closed instead of assuming modern Excel', async () => {
  const p = pane();
  delete p.sandbox.Office.context;
  await assert.rejects(p.sandbox.runWorkbookActions({ actions: [action] }), /ExcelApi 1.1/);
  assert.equal(p.loads.length, 0);
});

test('explicit size/autofit requires 1.2, basic formatting remains available in 1.1', () => {
  const p = pane({ api: 1 });
  for (const extra of [{ auto_fit: true }, { column_width: 20 }, { row_height: 15 }]) {
    assert.throws(() => p.sandbox.assertCompatibleActions([{ type: 'format_cells', ...extra }]), /ExcelApi 1.2/);
  }
  assert.doesNotThrow(() => p.sandbox.assertCompatibleActions([{ type: 'format_cells', bold: true }]));
});

test('1.1 full-sheet selection stays bounded and blank used range is conservatively A1', async () => {
  const p = pane({ api: 1, rows: 1048576, columns: 16384 });
  p.sheet.getUsedRangeOrNullObject = () => { throw new Error('unsupported'); };
  await p.sandbox.readWorkbookContext();
  assert.equal(p.state.workbook.sheets[0].usedRange, 'Sheet1!A1');
  for (const load of p.loads.filter(load => load.fields.includes('values') || load.fields.includes('formulas'))) {
    assert.ok(load.rows <= 100 && load.columns <= 16);
  }
});

test('absolute range sizing uses 1.1 cells and bounding rectangle at non-A1 origins', () => {
  const p = pane({ api: 1 });
  const cell = (row, column) => ({ row, column,
    getCell: (r, c) => cell(row + r, column + c),
    getBoundingRect: end => ({ top: row, left: column, bottom: end.row, right: end.column }),
  });
  assert.deepEqual(p.sandbox.sizedRange(cell(22, 7), 3, 4), { top: 22, left: 7, bottom: 24, right: 10 });
  assert.throws(() => p.sandbox.sizedRange(cell(0, 0), 0, 2), /positive integers/);
});

for (const api of [1, 2, 6, 7]) {
  test(`new-sheet presentation on 1.${api} calls only available optional APIs`, () => {
    const p = pane({ api });
    let autofits = 0, freezes = 0;
    const format = { font: {}, fill: {}, borders: { getItem: () => ({}) } };
    if (api >= 2) {
      format.autofitColumns = () => autofits++;
      format.autofitRows = () => autofits++;
    }
    const target = { format, getCell() { return this; }, getBoundingRect() { return this; } };
    const sheet = api >= 7 ? { freezePanes: { freezeRows: () => freezes++ } } : {};
    p.sandbox.Excel.BorderIndex = {};
    p.sandbox.Excel.BorderLineStyle = { continuous: 'Continuous' };
    p.sandbox.Excel.BorderWeight = {};
    p.sandbox.Excel.HorizontalAlignment = {};
    p.sandbox.Excel.VerticalAlignment = {};
    p.sandbox.applyProfessionalTableFormat(sheet, target, [['Heading', 'Value'], ['Item', 3]]);
    assert.equal(autofits > 0, api >= 2);
    assert.equal(freezes, api >= 7 ? 1 : 0);
  });
}

test('failed formatting protects earlier Undo records', async () => {
  const p = pane();
  await p.sandbox.writeCellsAction(action);
  p.sandbox.formatCellsAction = async () => { throw new Error('partial sync failure'); };
  await p.sandbox.runWorkbookActions({ actions: [{ type: 'format_cells', range: 'Sheet1!A1' }] });
  assert.equal(p.state.undoStack.at(-1).kind, 'non_undoable');
  await p.sandbox.undoLast();
  assert.deepEqual(p.range.formulas, [[20]]);
});

test('failed create_sheet prevents Undo from reaching an unrelated prior write', async () => {
  const p = pane();
  await p.sandbox.writeCellsAction(action);
  p.sandbox.createSheetAction = async () => { throw new Error('unverified sheet creation'); };
  await p.sandbox.runWorkbookActions({ actions: [{ type: 'create_sheet' }] });
  assert.equal(p.state.undoStack.at(-1).kind, 'non_undoable');
});

test('capability hints are bounded and preserved by workbook normalization', () => {
  const workbook = normalizeWorkbook({ excelApi: ['1.1', '1.7', '1.7', 'invented', { text: 'untrusted' }] });
  assert.deepEqual(workbook.excelApi, ['1.1', '1.7']);
  assert.equal('excelApi' in normalizeWorkbook({}), false);
});

test('activity fetch supports browsers without AbortSignal.timeout and clears timers', async () => {
  const p = pane();
  let cleared = 0;
  p.sandbox.clearTimeout = timer => { cleared++; clearTimeout(timer); };
  p.sandbox.fetch = async (url, options) => {
    assert.ok(options.signal instanceof AbortSignal);
    return { ok: true, json: async () => ({ ok: true }) };
  };
  assert.deepEqual(await p.sandbox.fetchJsonWithDeadline('/api/activity', {}, 3000), { ok: true });
  assert.equal(cleared, 1);
  p.sandbox.fetch = async () => { throw new Error('network failure'); };
  await assert.rejects(p.sandbox.fetchJsonWithDeadline('/api/activity', {}, 3000), /network failure/);
  assert.equal(cleared, 2);
});

test('activity deadline also aborts a stalled response body', async () => {
  const p = pane();
  p.sandbox.fetch = async (url, { signal }) => ({ ok: true,
    json: () => new Promise((resolve, reject) => signal.addEventListener('abort', () => reject(new Error('aborted')))),
  });
  await assert.rejects(p.sandbox.fetchJsonWithDeadline('/api/activity', {}, 5), /aborted/);
});

test('sheet deletion lookup works without 1.4 null-object APIs', async () => {
  const p = pane({ api: 1 });
  let deleted = false;
  p.sheet.delete = () => { deleted = true; };
  assert.match(await p.sandbox.deleteSheetAction({ name: 'missing' }), /not found/);
  assert.equal(deleted, false);
  assert.match(await p.sandbox.deleteSheetAction({ name: 'sheet1' }), /Deleted/);
  assert.equal(deleted, true);
});

test('unsaved workbooks never read or write a shared persisted chat', () => {
  const p = pane();
  let reads = 0, writes = 0;
  p.sandbox.localStorage = {
    getItem() { reads++; return JSON.stringify({ history: [{ content: 'another workbook' }], messages: [] }); },
    setItem() { writes++; },
  };
  for (const key of ['', 'default']) {
    p.state.workbookKey = key;
    p.sandbox.loadChatHistory();
    p.persistHistory();
    assert.equal(p.state.history.length, 0);
  }
  assert.equal(reads, 0);
  assert.equal(writes, 0);
  p.state.workbookKey = 'file:///synthetic.xlsx';
  p.sandbox.loadChatHistory();
  p.persistHistory();
  assert.equal(reads, 1);
  assert.equal(writes, 1);
});

for (const api of [1, 0]) {
  test(`chat submission applies document actions directly with ExcelApi support=${api > 0}`, async () => {
    const p = pane({ api });
    // A legacy persisted preference must never stop a direct write for approval.
    p.sandbox.localStorage = { getItem: () => '1', setItem() {} };
    p.state.files = [{ name: 'synthetic.csv', size: 12 }];
    const captured = [];
    p.sandbox.askHermes = async (prompt, files) => {
      captured.push({ prompt, files });
      return { message: 'Import the attached table.', actions: [action], files: [] };
    };
    p.sandbox.checkBridgeHealth = async () => true;
    p.sandbox.startWorkIndicator = () => {};
    p.sandbox.setWorkStage = () => {};
    p.sandbox.stopWorkIndicator = () => {};
    p.sandbox.clearFiles = () => { p.state.files = []; };
    p.sandbox.verifyWrittenRange = async () => ({ values: p.range.values });
    p.sandbox.wireActions();
    p.elements.get('prompt').value = 'Put this document table into the workbook';
    await p.elements.get('chatForm').events.submit({ preventDefault() {} });
    assert.equal(captured.length, 1);
    assert.equal(captured[0].files[0].name, 'synthetic.csv');
    assert.equal(p.state.sending, false);
    assert.equal(p.elements.get('sendButton').disabled, false);
    assert.ok(!p.messages.some(message => /Waiting for review|Review these changes/.test(message[1])));
    if (api > 0) {
      assert.deepEqual(p.range.values, [[20]]);
      assert.equal(p.state.files.length, 0);
      assert.equal(p.state.lastSentFiles[0].name, 'synthetic.csv');
      await p.sandbox.undoLast();
      assert.deepEqual(p.range.values, [[10]]);
    } else {
      assert.deepEqual(p.range.values, [[10]]);
      assert.equal(p.state.files.length, 1);
      assert.ok(p.messages.some(message => /No changes were applied/.test(message[1])));
    }
  });
}
