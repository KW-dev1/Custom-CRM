// Lightweight diagnostic logger: jobId, stage, ok/fail, category, timestamp.
// Never logs tokens, recipient addresses, subjects or bodies.

import { KEYS } from './store.js';

const MAX_LOGS = 100;
const MAX_ERRORS = 20;

export function redact(text, secrets = []) {
  let out = String(text ?? '');
  for (const s of secrets) if (s && s.length >= 4) out = out.split(s).join('[redacted]');
  return out.replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi, '[email]').slice(0, 500);
}

export function createLogger({ store, now = Date.now, console: con = globalThis.console, secrets = () => [] } = {}) {
  async function log({ jobId = null, stage, ok = true, category = null, message = '' }) {
    const entry = { at: now(), jobId, stage, ok, category, message: redact(message, secrets()) };
    con.log?.(`[OEA] ${new Date(entry.at).toISOString()} job=${jobId ?? '-'} stage=${stage} ${ok ? 'ok' : 'FAIL'}${category ? ` cat=${category}` : ''}${entry.message ? ` ${entry.message}` : ''}`);
    if (!store) return entry;
    try {
      const logs = [...(await store.getList(KEYS.logs)), entry].slice(-MAX_LOGS);
      await store.setList(KEYS.logs, logs);
      if (!ok) {
        const errs = [...(await store.getList(KEYS.errors)), entry].slice(-MAX_ERRORS);
        await store.setList(KEYS.errors, errs);
      }
    } catch { /* logging must never break processing */ }
    return entry;
  }
  return { log };
}
