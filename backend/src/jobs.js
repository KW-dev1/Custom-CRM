// Job service: claims the next unprocessed row and records results in the status column.
//
// Status column lifecycle:  (empty) --claim--> pending --extension--> sent | failed | needs_review
// Safety rules:
//  * Only rows whose status cell is EMPTY are ever claimed; anything else is left alone.
//  * Rows with no address / subject / message yet are skipped untouched (still being drafted);
//    rows with malformed data (e.g. a broken email address) are marked `failed` with a reason.
//  * Claiming writes `pending` BEFORE the job is returned; a pending row is never auto-retried
//    (clear the cell by hand to requeue, after checking Outlook's Sent Items if unsure).
//  * Final statuses are terminal: a repeat of the same status is idempotent, a different one is a conflict.
//  * Every operation runs under one in-process lock. Run exactly ONE backend instance per sheet.

import { hasContent, rowToJob, rowHash, parseJobId } from './rows.js';

export class JobError extends Error {
  constructor(code, message, { status = 400, retryAfterSec = null } = {}) {
    super(message);
    this.name = 'JobError';
    this.code = code;
    this.status = status;
    this.retryAfterSec = retryAfterSec;
  }
}

export const FINAL = ['sent', 'failed', 'needs_review'];
export const REPORTABLE = ['processing', ...FINAL];
const MAX_INVALID_PER_CLAIM = 25;

export function createJobService({ sheet, cfg, now = Date.now, log = () => {} }) {
  const labels = cfg.statusLabels;
  const byLabel = new Map(Object.entries(labels).map(([k, v]) => [v.toLowerCase(), k]));
  const canonical = (cell) => {
    const v = String(cell ?? '').trim();
    return v === '' ? 'empty' : byLabel.get(v.toLowerCase()) || 'other';
  };
  const note = (text) => `${new Date(now()).toISOString()} ${String(text).replace(/\s+/g, ' ').slice(0, 300)}`;

  let chain = Promise.resolve();
  const locked = (fn) => {
    const run = chain.then(fn);
    chain = run.catch(() => {});
    return run;
  };
  let lastClaimAt = 0;

  async function locate({ rowNumber, hash }) {
    const row = await sheet.readRow(rowNumber);
    if (rowHash(row) === hash) return row;
    // The row was edited, deleted or moved (e.g. the sheet was sorted). Find it by content.
    const matches = (await sheet.readRows()).filter((r) => rowHash(r) === hash);
    const pending = matches.filter((r) => canonical(r.status) === 'pending');
    const pick = pending.length === 1 ? pending[0] : matches.length === 1 ? matches[0] : null;
    if (!pick) throw new JobError('NOT_FOUND', 'The row for this job no longer exists or its content changed.', { status: 404 });
    log('info', 'row_relocated', { from: rowNumber, to: pick.rowNumber });
    return pick;
  }

  return {
    /** @returns the next job (and marks its row pending), or null when nothing is waiting. */
    claimNext() {
      return locked(async () => {
        const min = cfg.minClaimIntervalSec;
        if (lastClaimAt && min > 0) {
          const waited = (now() - lastClaimAt) / 1000;
          if (waited < min) throw new JobError('RATE_LIMITED', 'Jobs are released at a limited rate.', { status: 429, retryAfterSec: Math.ceil(min - waited) });
        }
        const rows = await sheet.readRows();
        let invalid = 0;
        let skipped = 0;
        for (const r of rows) {
          if (canonical(r.status) !== 'empty' || !hasContent(r)) continue;
          const mapped = rowToJob(r, { maxCc: cfg.maxCcRecipients });
          if (!mapped.ok && mapped.incomplete) { skipped += 1; continue; } // still being drafted: leave untouched
          if (!mapped.ok) {
            if (invalid >= MAX_INVALID_PER_CLAIM) break;
            invalid += 1;
            await sheet.writeStatus(r.rowNumber, labels.failed, note(`invalid row: ${mapped.error}`));
            log('warn', 'row_invalid', { row: r.rowNumber, reason: mapped.error });
            continue;
          }
          await sheet.writeStatus(r.rowNumber, labels.pending, note('claimed'));
          lastClaimAt = now();
          log('info', 'claimed', { row: r.rowNumber, jobId: mapped.job.id, cc: mapped.job.cc.length });
          return mapped.job;
        }
        if (skipped) log('info', 'rows_incomplete', { count: skipped });
        return null;
      });
    },

    /** Record a result reported by the extension. Idempotent for repeated identical reports. */
    report(id, status, message = '') {
      return locked(async () => {
        const parsed = parseJobId(id);
        if (!parsed) throw new JobError('BAD_JOB_ID', 'Malformed job id.', { status: 400 });
        if (!REPORTABLE.includes(status)) throw new JobError('BAD_STATUS', `Status must be one of: ${REPORTABLE.join(', ')}.`, { status: 400 });
        const row = await locate(parsed);
        const current = canonical(row.status);

        if (status === 'processing') {
          if (current === 'empty') throw new JobError('NOT_CLAIMED', 'This job has not been claimed.', { status: 409 });
          return { status: current, applied: false, row: row.rowNumber };
        }
        if (current === status) return { status, applied: false, row: row.rowNumber };
        if (current !== 'pending' && current !== 'empty') {
          log('warn', 'status_conflict', { row: row.rowNumber, current, incoming: status });
          throw new JobError('CONFLICT', `Row is already "${row.status}"; refusing to change it to "${labels[status]}".`, { status: 409 });
        }
        await sheet.writeStatus(row.rowNumber, labels[status], note(`${status}${message ? `: ${message}` : ''}`));
        log('info', 'status_written', { row: row.rowNumber, status });
        return { status, applied: true, row: row.rowNumber };
      });
    },

    /** Current backend-side state of a job: empty (unclaimed) | pending | sent | failed | needs_review | other. */
    getState(id) {
      return locked(async () => {
        const parsed = parseJobId(id);
        if (!parsed) throw new JobError('BAD_JOB_ID', 'Malformed job id.', { status: 400 });
        const row = await locate(parsed);
        return { status: canonical(row.status), row: row.rowNumber };
      });
    },
  };
}
