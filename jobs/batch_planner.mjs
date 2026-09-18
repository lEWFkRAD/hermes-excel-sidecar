export function planBatch({ startRow, endRow, batchSize = 25, cursor = startRow }) {
  if (!Number.isInteger(startRow) || !Number.isInteger(endRow) || startRow < 1 || endRow < startRow) throw new Error("Invalid row bounds");
  if (!Number.isInteger(batchSize) || batchSize < 1 || batchSize > 250) throw new Error("batchSize must be 1..250");
  const normalizedCursor = Number.isInteger(cursor) ? cursor : startRow;
  if (normalizedCursor < startRow || normalizedCursor > endRow + 1) throw new Error("cursor outside row bounds");
  if (normalizedCursor > endRow) return { startRow, endRow, cursor: normalizedCursor, rows: [], done: true };
  const batchEnd = Math.min(endRow, normalizedCursor + batchSize - 1);
  return { startRow, endRow, cursor: normalizedCursor, rows: [{ startRow: normalizedCursor, endRow: batchEnd }], done: false };
}

export function nextCursor(result) {
  if (!result || result.done) return result?.endRow + 1;
  const row = result.rows?.at(-1);
  if (!row) return result.cursor;
  return row.endRow + 1;
}

export function isComplete(job) {
  return Number.isInteger(job?.cursor) && Number.isInteger(job?.endRow) && job.cursor > job.endRow;
}
