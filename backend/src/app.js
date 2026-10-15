// HTTP API consumed by the Chrome extension.
//   POST /api/email-jobs/claim          -> 200 {job} | 204
//   POST /api/email-jobs/:id/status     -> 200 {id,status,applied}
//   GET  /api/email-jobs/:id            -> 200 {id,status}
//   GET  /health                        -> 200 {ok:true}   (no auth, never touches the sheet)

import express from 'express';
import { createHash, timingSafeEqual } from 'node:crypto';
import { JobError, REPORTABLE } from './jobs.js';
import { SheetsError } from './sheets-client.js';
import { GoogleAuthError } from './google-auth.js';

const digest = (s) => createHash('sha256').update(String(s)).digest();

export function createApp({ service, cfg, log = () => {}, now = Date.now }) {
  const app = express();
  app.disable('x-powered-by');
  const keyDigests = cfg.apiKeys.map(digest);

  app.use((req, res, next) => {
    res.set({ 'cache-control': 'no-store', 'x-content-type-options': 'nosniff' });
    next();
  });

  app.get('/health', (req, res) => res.json({ ok: true }));

  // ---- authentication (+ a small failed-attempt limiter per client address) ----
  const failures = new Map();
  const WINDOW_MS = 60_000;
  const MAX_FAILS = 10;
  function authenticate(req, res, next) {
    const ip = req.ip;
    const rec = failures.get(ip);
    if (rec && now() - rec.since < WINDOW_MS && rec.count >= MAX_FAILS) {
      res.set('retry-after', String(Math.ceil((WINDOW_MS - (now() - rec.since)) / 1000)));
      return res.status(429).json({ error: { code: 'TOO_MANY_ATTEMPTS', message: 'Too many failed authentication attempts.' } });
    }
    const header = req.get('authorization') || '';
    const supplied = header.toLowerCase().startsWith('bearer ') ? header.slice(7).trim() : (req.get('x-api-key') || '').trim();
    const d = digest(supplied);
    const ok = supplied !== '' && keyDigests.some((k) => timingSafeEqual(k, d));
    if (!ok) {
      const fresh = !rec || now() - rec.since >= WINDOW_MS;
      failures.set(ip, { since: fresh ? now() : rec.since, count: fresh ? 1 : rec.count + 1 });
      if (failures.size > 1000) failures.clear();
      return res.status(401).json({ error: { code: 'UNAUTHORIZED', message: 'Missing or invalid API key.' } });
    }
    failures.delete(ip);
    return next();
  }

  const api = express.Router();
  api.use(authenticate);
  api.use(express.json({ limit: '20kb' }));
  const wrap = (fn) => (req, res, next) => Promise.resolve(fn(req, res)).catch(next);

  api.post('/email-jobs/claim', wrap(async (req, res) => {
    const job = await service.claimNext();
    if (!job) return res.status(204).end();
    return res.json({ job });
  }));

  api.post('/email-jobs/:id/status', wrap(async (req, res) => {
    const body = req.body && typeof req.body === 'object' ? req.body : {};
    if (body.jobId !== undefined && body.jobId !== req.params.id) {
      throw new JobError('ID_MISMATCH', 'jobId in the body does not match the URL.', { status: 400 });
    }
    if (typeof body.status !== 'string' || !REPORTABLE.includes(body.status)) {
      throw new JobError('BAD_STATUS', `status must be one of: ${REPORTABLE.join(', ')}.`, { status: 400 });
    }
    if (body.message !== undefined && typeof body.message !== 'string') {
      throw new JobError('BAD_MESSAGE', 'message must be a string.', { status: 400 });
    }
    const r = await service.report(req.params.id, body.status, (body.message || '').slice(0, 1000));
    res.json({ id: req.params.id, status: r.status, applied: r.applied });
  }));

  api.get('/email-jobs/:id', wrap(async (req, res) => {
    const s = await service.getState(req.params.id);
    res.json({ id: req.params.id, status: s.status });
  }));

  app.use('/api', api);
  app.use((req, res) => res.status(404).json({ error: { code: 'NOT_FOUND', message: 'No such endpoint.' } }));

  // eslint-disable-next-line no-unused-vars
  app.use((err, req, res, next) => {
    if (err instanceof JobError) {
      if (err.retryAfterSec) res.set('retry-after', String(err.retryAfterSec));
      return res.status(err.status).json({ error: { code: err.code, message: err.message } });
    }
    if (err instanceof SheetsError || err instanceof GoogleAuthError) {
      log('error', 'google_error', { kind: err.name, status: err.status ?? null, message: err.message });
      return res.status(502).json({ error: { code: 'SHEET_UNAVAILABLE', message: err.message } });
    }
    if (err && (err.type === 'entity.parse.failed' || err.type === 'entity.too.large')) {
      return res.status(err.status || 400).json({ error: { code: 'BAD_REQUEST', message: 'Invalid request body.' } });
    }
    log('error', 'unhandled', { message: err && err.message });
    return res.status(500).json({ error: { code: 'INTERNAL', message: 'Internal server error.' } });
  });
  return app;
}
