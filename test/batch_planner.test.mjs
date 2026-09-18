import test from "node:test";
import assert from "node:assert/strict";
import { isComplete, nextCursor, planBatch } from "../jobs/batch_planner.mjs";

test("plans a bounded 4,000-row job without skipping rows", () => {
  const first = planBatch({ startRow: 10, endRow: 4009, batchSize: 25, cursor: 10 });
  assert.deepEqual(first.rows, [{ startRow: 10, endRow: 34 }]);
  assert.equal(nextCursor(first), 35);
  const last = planBatch({ startRow: 10, endRow: 4009, batchSize: 25, cursor: 4000 });
  assert.deepEqual(last.rows, [{ startRow: 4000, endRow: 4009 }]);
  assert.equal(nextCursor(last), 4010);
});

test("cursor restart is deterministic and completion is explicit", () => {
  assert.deepEqual(planBatch({ startRow: 1, endRow: 3, batchSize: 2, cursor: 4 }), { startRow: 1, endRow: 3, cursor: 4, rows: [], done: true });
  assert.equal(isComplete({ cursor: 4, endRow: 3 }), true);
  assert.equal(isComplete({ cursor: 3, endRow: 3 }), false);
});

test("rejects invalid bounds and oversized batches", () => {
  assert.throws(() => planBatch({ startRow: 0, endRow: 2 }));
  assert.throws(() => planBatch({ startRow: 2, endRow: 1 }));
  assert.throws(() => planBatch({ startRow: 1, endRow: 2, batchSize: 251 }));
  assert.throws(() => planBatch({ startRow: 1, endRow: 2, cursor: 4 }));
});
