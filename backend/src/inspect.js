// Read-only summary of a data source: used by `npm run check:source` to show what the backend will do.
// Counts only; never returns addresses, subjects or message text.

import { hasContent, rowToJob } from './rows.js';

export function summarizeRows(rows, cfg) {
  const byLabel = new Map(Object.entries(cfg.statusLabels).map(([k, v]) => [v.toLowerCase(), k]));
  const out = { total: rows.length, blank: 0, ready: 0, incomplete: 0, invalid: [], pending: 0, sent: 0, failed: 0, needs_review: 0, other: 0 };
  for (const r of rows) {
    const cell = String(r.status ?? '').trim();
    if (cell) { out[byLabel.get(cell.toLowerCase()) || 'other'] += 1; continue; }
    if (!hasContent(r)) { out.blank += 1; continue; }
    const m = rowToJob(r, { maxCc: cfg.maxCcRecipients });
    if (m.ok) out.ready += 1;
    else if (m.incomplete) out.incomplete += 1;
    else out.invalid.push({ row: r.rowNumber, reason: m.error.replace(/:.*$/, '') });
  }
  return out;
}
