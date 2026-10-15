// Pure helpers: parse the email-list cell and map a sheet row to an email job.

import { createHash } from 'node:crypto';

const EMAIL_RE = /^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?)+$/;
const MAX_EMAIL_LEN = 254;
export const MAX_SUBJECT = 998;
export const MAX_BODY = 100000;

/**
 * Parse the "array of email addresses" cell. Accepts a JSON array (["a@x.com","b@y.com"]),
 * or any list separated by commas, semicolons, spaces or newlines, with optional
 * brackets/quotes or "Name <addr>" forms.
 * @returns {{emails: string[], invalid: string[]}} emails are de-duplicated (case-insensitive), in order.
 */
export function parseEmailList(cell) {
  let text = String(cell ?? '').trim();
  if (!text) return { emails: [], invalid: [] };
  if (text.startsWith('[')) {
    try {
      const arr = JSON.parse(text);
      if (Array.isArray(arr)) text = arr.map((x) => String(x ?? '')).join('\n');
    } catch { /* fall through: treat as loose text */ }
  }
  const tokens = text.split(/[\s,;<>"'()[\]]+/).filter((t) => t.includes('@'));
  const emails = [];
  const invalid = [];
  const seen = new Set();
  for (const raw of tokens) {
    const t = raw.replace(/^[.:]+|[.:]+$/g, '');
    if (!EMAIL_RE.test(t) || t.length > MAX_EMAIL_LEN) { invalid.push(raw); continue; }
    const key = t.toLowerCase();
    if (!seen.has(key)) { seen.add(key); emails.push(t); }
  }
  return { emails, invalid };
}

export const hasContent = (r) => [r.emails, r.subject, r.body].some((v) => String(v ?? '').trim() !== '');

/** Stable fingerprint of the row's payload; lets us detect edits/moves before writing a status. */
export function rowHash(r) {
  const h = createHash('sha256');
  h.update(JSON.stringify([String(r.emails ?? ''), String(r.subject ?? ''), String(r.body ?? '')]));
  return h.digest('hex').slice(0, 12);
}
export const makeJobId = (rowNumber, hash) => `r${rowNumber}-${hash}`;
export function parseJobId(id) {
  const m = /^r(\d{1,9})-([0-9a-f]{12})$/.exec(String(id ?? ''));
  return m ? { rowNumber: Number(m[1]), hash: m[2] } : null;
}

/**
 * Map a sheet row to a job. First valid address is TO, the rest are CC.
 * Any malformed address, or a missing subject/body, makes the row invalid so we never send to a
 * partial or silently altered recipient list.
 * @returns {{ok:true, job:object} | {ok:false, error:string}}
 */
export function rowToJob(r, { maxCc = 20 } = {}) {
  const { emails, invalid } = parseEmailList(r.emails);
  if (invalid.length) return { ok: false, error: `Invalid email address in column: ${invalid.slice(0, 3).join(', ')}` };
  if (!emails.length) return { ok: false, error: 'No email address found.' };
  if (emails.length - 1 > maxCc) return { ok: false, error: `Too many CC recipients (${emails.length - 1} > ${maxCc}).` };

  const subject = String(r.subject ?? '').replace(/\s*[\r\n]+\s*/g, ' ').trim();
  if (!subject) return { ok: false, error: 'Subject is empty.' };
  if (subject.length > MAX_SUBJECT) return { ok: false, error: 'Subject is too long.' };

  const body = String(r.body ?? '').replace(/\r\n?/g, '\n').replace(/\s+$/, '');
  if (!body.trim()) return { ok: false, error: 'Message content is empty.' };
  if (body.length > MAX_BODY) return { ok: false, error: 'Message content is too long.' };

  return {
    ok: true,
    job: {
      id: makeJobId(r.rowNumber, rowHash(r)),
      to: emails[0], cc: emails.slice(1), subject, body,
      metadata: { row: r.rowNumber },
    },
  };
}
