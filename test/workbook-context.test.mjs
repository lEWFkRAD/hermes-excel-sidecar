import test from "node:test";
import assert from "node:assert/strict";
import { describeSheet, paginateRange, snapshotWorkbook } from "../src/workbook_context.mjs";

test("workbook snapshots include bounded sheet metadata and exclude hidden sheets by default", () => {
  const result = snapshotWorkbook({ workbookId: "wb", activeSheet: "Input", sheets: [{ name: "Input", usedRange: "A1:Z4009", rowCount: 4009, columnCount: 26, tables: [{ name: "Offers", range: "A1:Z10" }], namedRanges: ["Materials"] }, { name: "Secrets", hidden: true }] });
  assert.equal(result.snapshot.selectionIsTargetHint, true);
  assert.equal(result.snapshot.sheets[0].tables[0].name, "Offers");
  assert.equal(result.snapshot.sheets[1].excluded, true);
});

test("snapshot truncates under a declared budget and provides a cursor", () => {
  const result = snapshotWorkbook({ sheets: Array.from({ length: 10 }, (_, i) => ({ name: `Sheet-${i}`, usedRange: "A1:Z4000" })) }, { budget: 200 });
  assert.equal(result.truncated, true);
  assert.ok(Number.isInteger(result.nextCursor));
});

test("range pagination is bounded and resumable", () => {
  assert.deepEqual(paginateRange("Sheet1!A1:Z4009", 0, 8), { address: "Sheet1!A", cursor: 0, nextCursor: 8, truncated: true });
  assert.equal(paginateRange("A1", 0, 10).nextCursor, null);
});
