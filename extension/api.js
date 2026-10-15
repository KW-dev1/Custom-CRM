// Backend API client. Request construction and response parsing are pure and unit-tested;
// network calls take an injectable fetch. Contract details are PLACEHOLDERS (see README).

import { normalizeSettings } from './config.js';

export class ApiError extends Error {
  /** kind: network | auth | rate_limit | server | client | invalid_response */
  constructor(kind, message, { status = null, retryAfterMs = null } = {}) {
    super(message);
    this.name = 'ApiError';
    this.kind = kind;
    this.status = status;
    this.retryAfterMs = retryAfterMs;
  }
}

const EMAIL_RE = /^[^\s@,;<>()"]+@[^\s@,;<>()"]+\.[^\s@,;<>()"]{2,}$/;
export const LIMITS = { id: 200, subject: 998, body: 200000, email: 254, cc: 20 };

export function buildUrl(settings, path, params = {}) {
  const s = normalizeSettings(settings);
  let p = path;
  for (const [k, v] of Object.entries(params)) p = p.replaceAll(`{${k}}`, encodeURIComponent(String(v)));
  return s.apiBaseUrl + p;
}

export function buildHeaders(settings, withBody = false) {
  const s = normalizeSettings(settings);
  const h = { Accept: 'application/json' };
  if (withBody) h['Content-Type'] = 'application/json';
  h[s.authHeader] = s.authScheme ? `${s.authScheme} ${s.apiKey}` : s.apiKey;
  return h;
}

export function buildClaimRequest(settings) {
  const s = normalizeSettings(settings);
  return { url: buildUrl(s, s.claimPath), init: { method: s.claimMethod, headers: buildHeaders(s, s.claimMethod !== 'GET') , ...(s.claimMethod === 'POST' ? { body: '{}' } : {}) } };
}

/** report: { jobId, status (canonical), message?, timestamp } */
export function buildStatusRequest(settings, report) {
  const s = normalizeSettings(settings);
  const payload = {
    jobId: report.jobId,
    status: s.statusValues[report.status] ?? report.status,
    timestamp: report.timestamp || new Date().toISOString(),
  };
  if (report.message) payload.message = String(report.message).slice(0, 1000);
  return {
    url: buildUrl(s, s.statusPath, { id: report.jobId }),
    init: { method: s.statusMethod, headers: buildHeaders(s, true), body: JSON.stringify(payload) },
  };
}

export function buildJobStateRequest(settings, jobId) {
  const s = normalizeSettings(settings);
  return { url: buildUrl(s, s.jobStatePath, { id: jobId }), init: { method: 'GET', headers: buildHeaders(s) } };
}

const isObj = (v) => v !== null && typeof v === 'object' && !Array.isArray(v);

/** Accepts `{job: {...}}`, `{job: null}` or a bare job object. Returns a job or null (no work). */
export function parseJobResponse(body) {
  if (body === null || body === undefined || body === '') return null;
  if (!isObj(body)) throw new ApiError('invalid_response', 'Backend returned a non-object job response.');
  const job = 'job' in body ? body.job : body;
  if (job === null || job === undefined) return null;
  if (!isObj(job)) throw new ApiError('invalid_response', 'Backend returned a malformed job.');
  if (Object.keys(job).length === 0) return null;
  return {
    id: job.id === undefined || job.id === null ? job.id : String(job.id),
    to: job.to,
    cc: job.cc === undefined || job.cc === null ? [] : job.cc,
    subject: job.subject,
    body: job.body,
    metadata: isObj(job.metadata) ? job.metadata : undefined,
  };
}

export function validateJob(job) {
  const errors = [];
  if (!isObj(job)) return { ok: false, errors: ['Job is not an object.'] };
  if (typeof job.id !== 'string' || !job.id.trim() || job.id.length > LIMITS.id) errors.push('Missing or invalid job id.');
  if (typeof job.to !== 'string' || !job.to.trim()) errors.push('Missing recipient (to).');
  else if (job.to.length > LIMITS.email || !EMAIL_RE.test(job.to.trim())) errors.push('Recipient must be a single valid email address.');
  const cc = job.cc === undefined ? [] : job.cc;
  if (!Array.isArray(cc)) errors.push('cc must be an array of email addresses.');
  else if (cc.length > LIMITS.cc) errors.push(`Too many CC recipients (max ${LIMITS.cc}).`);
  else {
    const bad = cc.filter((a) => typeof a !== 'string' || a.length > LIMITS.email || !EMAIL_RE.test(a.trim()));
    if (bad.length) errors.push('Every CC entry must be a single valid email address.');
    const all = [job.to, ...cc].filter((a) => typeof a === 'string').map((a) => a.trim().toLowerCase());
    if (new Set(all).size !== all.length) errors.push('Recipients must be unique (To and CC overlap or repeat).');
  }
  if (typeof job.subject !== 'string' || !job.subject.trim()) errors.push('Missing subject.');
  else if (/[\r\n]/.test(job.subject) || job.subject.length > LIMITS.subject) errors.push('Subject must be a single line of at most 998 characters.');
  if (typeof job.body !== 'string' || !job.body.trim()) errors.push('Missing body.');
  else if (job.body.length > LIMITS.body) errors.push('Body is too large.');
  return { ok: errors.length === 0, errors };
}

export function classifyHttpError(status, headers) {
  const ra = headers && typeof headers.get === 'function' ? Number(headers.get('retry-after')) : NaN;
  const retryAfterMs = Number.isFinite(ra) && ra >= 0 ? ra * 1000 : null;
  if (status === 401 || status === 403) return new ApiError('auth', `Backend rejected the credentials (HTTP ${status}). Check the API key.`, { status });
  if (status === 429) return new ApiError('rate_limit', 'Backend rate limit reached (HTTP 429).', { status, retryAfterMs });
  if (status >= 500) return new ApiError('server', `Backend server error (HTTP ${status}).`, { status });
  return new ApiError('client', `Backend rejected the request (HTTP ${status}).`, { status });
}

async function call({ url, init }, { fetchImpl = globalThis.fetch, timeoutMs = 15000 } = {}) {
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), timeoutMs);
  let res;
  try {
    res = await fetchImpl(url, { ...init, signal: ctl.signal, credentials: 'omit', cache: 'no-store' });
  } catch {
    throw new ApiError('network', 'Could not reach the backend (network error or timeout).');
  } finally {
    clearTimeout(timer);
  }
  if (!res.ok) throw classifyHttpError(res.status, res.headers);
  if (res.status === 204) return null;
  const text = await res.text();
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch {
    throw new ApiError('invalid_response', 'Backend returned invalid JSON.');
  }
}

/** Claim the next job. NOT retried: a retried claim could reserve a second job. */
export async function claimJob(settings, opts) {
  return parseJobResponse(await call(buildClaimRequest(settings), opts));
}

/** Optional pre-send state check. Returns the status string or null if unavailable. */
export async function getJobState(settings, jobId, opts) {
  const s = normalizeSettings(settings);
  if (!s.jobStatePath) return null;
  const body = await call(buildJobStateRequest(s, jobId), opts);
  const st = isObj(body) ? (body.status ?? body.job?.status) : null;
  return typeof st === 'string' ? st : null;
}

const RETRYABLE = new Set(['network', 'server', 'rate_limit']);
const defaultSleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * Report a job status. Status updates are keyed by job id and idempotent, so they are
 * retried with backoff. This never triggers a re-send of the email.
 */
export async function reportStatus(settings, report, { retries = 3, sleep = defaultSleep, ...opts } = {}) {
  const req = buildStatusRequest(settings, report);
  let attempt = 0;
  for (;;) {
    try {
      return await call(req, opts);
    } catch (e) {
      if (!(e instanceof ApiError) || !RETRYABLE.has(e.kind) || attempt >= retries) throw e;
      const wait = e.retryAfterMs ?? Math.min(1000 * 2 ** attempt, 8000);
      attempt += 1;
      await sleep(Math.min(wait, 30000));
    }
  }
}
