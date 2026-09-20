#!/usr/bin/env node

import { readFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

// Each schema action must deliberately name its fixed execution or pre-application route.
// Labels and unrelated conditionals cannot satisfy this contract.
const PANE_ROUTES = Object.freeze({
  write_cells: 'writeCellsAction', create_sheet: 'createSheetAction', format_cells: 'formatCellsAction', conditional_format: 'conditionalFormatAction',
  merge_cells: 'mergeCellsAction', unmerge_cells: 'unmergeCellsAction', insert_rows: 'insertCellsAction', insert_columns: 'insertCellsAction',
  delete_rows: 'deleteCellsAction', delete_columns: 'deleteCellsAction', set_column_width: 'setSizeAction', set_row_height: 'setSizeAction',
  freeze_panes: 'freezePanesAction', unfreeze_panes: 'unfreezePanesAction', autofit: 'autofitAction', rename_sheet: 'renameSheetAction',
  delete_sheet: 'deleteSheetAction', sort_range: 'sortRangeAction', clear_range: 'clearRangeAction', create_table: 'createTableAction',
  delete_table: 'deleteTableAction', create_chart: 'createChartAction', auto_filter: 'autoFilterAction', remove_filter: 'removeFilterAction',
  export: 'exportAction', read_range: 'executeReadRange', request_external_access: 'requestExternalAccessApproval',
  cynteka_search: 'performCyntekaSearch', cynteka_query: 'performCyntekaQuery',
});

// These are source-level dispatch bindings, not handler-name hints. Shared handlers
// retain their action-specific argument (rows vs columns) in the required binding.
const PANE_BINDINGS = Object.freeze({
  write_cells: 'action.type === "write_cells" ? await writeCellsAction(action) : await createSheetAction(action)',
  create_sheet: 'action.type === "write_cells" ? await writeCellsAction(action) : await createSheetAction(action)',
  format_cells: 'action.type === "format_cells") {\n        statusLines.push(await formatCellsAction(action))',
  conditional_format: 'action.type === "conditional_format") {\n        statusLines.push(await conditionalFormatAction(action))',
  merge_cells: 'merge_cells: (action) => mergeCellsAction(action)',
  unmerge_cells: 'unmerge_cells: (action) => unmergeCellsAction(action)',
  insert_rows: 'insert_rows: (action) => insertCellsAction(action, "rows")',
  insert_columns: 'insert_columns: (action) => insertCellsAction(action, "columns")',
  delete_rows: 'delete_rows: (action) => deleteCellsAction(action, "rows")',
  delete_columns: 'delete_columns: (action) => deleteCellsAction(action, "columns")',
  set_column_width: 'set_column_width: (action) => setSizeAction(action)',
  set_row_height: 'set_row_height: (action) => setSizeAction(action)',
  freeze_panes: 'freeze_panes: (action) => freezePanesAction(action)',
  unfreeze_panes: 'unfreeze_panes: (action) => unfreezePanesAction(action)',
  autofit: 'autofit: (action) => autofitAction(action)',
  rename_sheet: 'rename_sheet: (action) => renameSheetAction(action)',
  delete_sheet: 'delete_sheet: (action) => deleteSheetAction(action)',
  sort_range: 'sort_range: (action) => sortRangeAction(action)',
  clear_range: 'clear_range: (action) => clearRangeAction(action)',
  create_table: 'create_table: (action) => createTableAction(action)',
  delete_table: 'delete_table: (action) => deleteTableAction(action)',
  create_chart: 'create_chart: (action) => createChartAction(action)',
  auto_filter: 'auto_filter: (action) => autoFilterAction(action)',
  remove_filter: 'remove_filter: (action) => removeFilterAction(action)',
  export: 'if (action.type === "export") statusLines.push(await exportAction(action))',
  read_range: 'readActions.map((action) => executeReadRange(action.range))',
  request_external_access: 'await requestExternalAccessApproval(action)',
  cynteka_search: 'await performCyntekaSearch(action)',
  cynteka_query: 'await performCyntekaQuery(action)',
});

function protocolActions(root) {
  const code = [
    'import importlib.util, importlib, sys, pathlib, json',
    'root = pathlib.Path(sys.argv[1])',
    'spec = importlib.util.spec_from_file_location("sidecar", root / "__init__.py", submodule_search_locations=[str(root)])',
    'package = importlib.util.module_from_spec(spec)',
    'sys.modules["sidecar"] = package',
    'spec.loader.exec_module(package)',
    'print(json.dumps(sorted(importlib.import_module("sidecar.excel_tool").ACTION_TYPES)))',
  ].join('; ');
  const result = spawnSync('python', ['-c', code, root], { encoding: 'utf8' });
  if (result.status !== 0) throw new Error(`Could not load protocol actions: ${result.stderr.trim()}`);
  return new Set(JSON.parse(result.stdout));
}

export function stripJsComments(source) {
  let output = "";
  let quote = "";
  for (let i = 0; i < source.length; i += 1) {
    const char = source[i];
    const next = source[i + 1];
    if (quote) {
      output += char;
      if (char === "\\") { output += next || ""; i += 1; }
      else if (char === quote) quote = "";
      continue;
    }
    if (char === "\"" || char === "'" || char === "`") { quote = char; output += char; continue; }
    if (char === "/" && next === "/") {
      while (i < source.length && source[i] !== "\n") i += 1;
      if (i < source.length) output += "\n";
      continue;
    }
    if (char === "/" && next === "*") {
      i += 2;
      while (i < source.length && !(source[i] === "*" && source[i + 1] === "/")) {
        if (source[i] === "\n") output += "\n";
        i += 1;
      }
      i += 1;
      continue;
    }
    output += char;
  }
  return output;
}

function functionBody(source, functionName, nextFunctionName) {
  const start = source.indexOf(`function ${functionName}(`);
  const end = source.indexOf(`function ${nextFunctionName}(`, start + 1);
  return start >= 0 && end >= 0 ? source.slice(start, end) : '';
}

export function actionCoverage(root = ROOT) {
  const broker = readFileSync(path.join(root, 'broker/server.mjs'), 'utf8');
  const pane = stripJsComments(readFileSync(path.join(root, 'taskpane.js'), 'utf8'));
  const protocol = protocolActions(root);
  const normalizer = functionBody(broker, 'normalizeAction', 'normalizeActions');
  const brokerActions = new Set([...normalizer.matchAll(/type\s*===\s*"([a-z_]+)"/g)].map((match) => match[1]));

  return {
    protocolActions: [...protocol].sort(),
    missingBroker: [...protocol].filter((type) => !brokerActions.has(type)).sort(),
    missingPane: [...protocol].filter((type) => !PANE_ROUTES[type] || !pane.includes(PANE_ROUTES[type])).sort(),
    missingBindings: [...protocol].filter((type) => !PANE_BINDINGS[type] || !pane.includes(PANE_BINDINGS[type])).sort(),
    unknownRoutes: Object.keys(PANE_ROUTES).filter((type) => !protocol.has(type)).sort(),
  };
}

if (import.meta.url === pathToFileUrl(process.argv[1])) {
  const coverage = actionCoverage();
  const failures = [...coverage.missingBroker, ...coverage.missingPane, ...coverage.missingBindings, ...coverage.unknownRoutes];
  if (failures.length) { console.error(JSON.stringify(coverage, null, 2)); process.exitCode = 1; }
  else console.log(`Action coverage passed for ${coverage.protocolActions.length} typed actions.`);
}

function pathToFileUrl(value) { return new URL(`file://${path.resolve(value)}`).href; }
