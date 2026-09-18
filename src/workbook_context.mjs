const DEFAULT_BUDGET = 12000;
const MAX_SHEETS = 50;

function boundedText(value, max = 160) { return String(value ?? "").slice(0, max); }

export function describeSheet(sheetMetadata = {}, policy = {}) {
  const hidden = Boolean(sheetMetadata.hidden);
  if (hidden && policy.includeHidden !== true) return { name: boundedText(sheetMetadata.name), hidden: true, excluded: true };
  return {
    name: boundedText(sheetMetadata.name),
    hidden,
    excluded: false,
    usedRange: boundedText(sheetMetadata.usedRange, 80),
    rowCount: Number.isInteger(sheetMetadata.rowCount) ? sheetMetadata.rowCount : null,
    columnCount: Number.isInteger(sheetMetadata.columnCount) ? sheetMetadata.columnCount : null,
    tables: Array.isArray(sheetMetadata.tables) ? sheetMetadata.tables.slice(0, 50).map((table) => ({ name: boundedText(table?.name), range: boundedText(table?.range, 80) })) : [],
    namedRanges: Array.isArray(sheetMetadata.namedRanges) ? sheetMetadata.namedRanges.slice(0, 50).map((name) => boundedText(name, 120)) : [],
  };
}

export function snapshotWorkbook(metadata = {}, policy = {}) {
  const budget = Number.isInteger(policy.budget) && policy.budget > 0 ? policy.budget : DEFAULT_BUDGET;
  const sheets = Array.isArray(metadata.sheets) ? metadata.sheets.slice(0, MAX_SHEETS).map((sheet) => describeSheet(sheet, policy)) : [];
  const snapshot = { workbookId: boundedText(metadata.workbookId, 160), activeSheet: boundedText(metadata.activeSheet), sheets, selectionIsTargetHint: true };
  let serialized = JSON.stringify(snapshot);
  if (serialized.length <= budget) return { snapshot, truncated: false, nextCursor: null };
  const kept = [];
  for (const sheet of sheets) {
    const candidate = { ...snapshot, sheets: [...kept, sheet] };
    if (JSON.stringify(candidate).length > budget) break;
    kept.push(sheet);
  }
  return { snapshot: { ...snapshot, sheets: kept }, truncated: kept.length < sheets.length, nextCursor: kept.length < sheets.length ? kept.length : null };
}

export function paginateRange(address, cursor = 0, budget = 100) {
  const text = boundedText(address, 160);
  if (!Number.isInteger(cursor) || cursor < 0) throw new Error("cursor must be non-negative");
  if (!Number.isInteger(budget) || budget < 1) throw new Error("budget must be positive");
  const end = Math.min(text.length, cursor + budget);
  return { address: text.slice(cursor, end), cursor, nextCursor: end < text.length ? end : null, truncated: end < text.length };
}
