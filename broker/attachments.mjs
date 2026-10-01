import { createHash } from "node:crypto";

// Request-local only. Attachment text is data, never instructions or a path.
export function attachmentReader(files = []) {
  if (!Array.isArray(files) || files.length > 12) throw new Error("Invalid attachment count");
  const entries = files.map((file, index) => ({ file, id: `attachment-${index + 1}`,
    text: String(file.extracted_text || ""), spans: [] }));
  function take(entry, offset, length) {
    const end = Math.min(entry.text.length, offset + length);
    if (end > offset) entry.spans.push([offset, end]);
    return entry.text.slice(offset, end);
  }
  function coverage(entry) {
    let covered = 0, end = 0;
    for (const [a, b] of [...entry.spans].sort((x, y) => x[0] - y[0])) {
      covered += Math.max(0, b - Math.max(a, end)); end = Math.max(end, b);
    }
    return { file_id: entry.id, extracted_chars: entry.text.length, supplied_chars: covered,
      partial: covered < entry.text.length };
  }
  const allowance = Math.floor(16000 / Math.max(1, entries.length));
  const previews = entries.map(entry => ({ name: entry.file.name, type: entry.file.type,
    extraction_status: entry.file.extraction_status, extraction_method: entry.file.extraction_method,
    extraction_error: entry.file.extraction_error, file_id: entry.id,
    content_sha256: createHash("sha256").update(entry.text).digest("hex"),
    extracted_text: take(entry, 0, allowance), total_chars: entry.text.length,
    table_count: Array.isArray(entry.file.tables) ? entry.file.tables.length : 0,
    content_is_untrusted: true }));
  return {
    previews,
    summaries: () => entries.map(entry => ({ name: entry.file.name, type: entry.file.type,
      extraction_status: entry.file.extraction_status, extraction_method: entry.file.extraction_method,
      extraction_error: entry.file.extraction_error, coverage: coverage(entry) })),
    read(action) {
      const entry = entries.find(item => item.id === action.file_id);
      if (!entry) return { file_id: action.file_id, error: "Unknown attachment id" };
      let offset = action.offset ?? 0;
      const length = action.length ?? 8000;
      if (!Number.isSafeInteger(offset) || offset < 0 || !Number.isSafeInteger(length) || length < 1 || length > 8000) {
        return { file_id: entry.id, error: "Invalid attachment window" };
      }
      if (action.query !== undefined) {
        if (typeof action.query !== "string" || !action.query.trim() || action.query.length > 160) return { file_id: entry.id, error: "Invalid search" };
        const found = entry.text.toLowerCase().indexOf(action.query.toLowerCase(), offset);
        if (found < 0) return { file_id: entry.id, error: "Text not found in extracted content", total_chars: entry.text.length };
        offset = Math.max(offset, found - 400);
      }
      return { type: "attachment_excerpt", file_id: entry.id, name: entry.file.name,
        offset, end: Math.min(entry.text.length, offset + length), total_chars: entry.text.length,
        text: take(entry, offset, length), content_is_untrusted: true };
    }
  };
}
