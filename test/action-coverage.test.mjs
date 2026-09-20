import test from 'node:test';
import assert from 'node:assert/strict';
import { actionCoverage, stripJsComments } from '../scripts/check-action-coverage.mjs';

test('action coverage ignores dispatch-shaped JavaScript comments', () => {
  const source = 'const live = 1; // delete_columns: (action) => deleteCellsAction(action, "columns")\n/* insert_rows: (action) => insertCellsAction(action, "rows") */';
  assert.ok(!stripJsComments(source).includes('delete_columns:'));
  assert.ok(!stripJsComments(source).includes('insert_rows:'));
});

test('every typed protocol action has broker normalization and pane handling or an explicit interception', () => {
  const coverage = actionCoverage();
  assert.equal(coverage.protocolActions.length, 29);
  assert.deepEqual(coverage.protocolActions.filter((type) => type.includes('_rows') || type.includes('_columns')), [
    'delete_columns', 'delete_rows', 'insert_columns', 'insert_rows',
  ]);
  assert.deepEqual(coverage.missingBroker, []);
  assert.deepEqual(coverage.missingPane, []);
  assert.deepEqual(coverage.missingBindings, []);
  assert.deepEqual(coverage.unknownRoutes, []);
});
