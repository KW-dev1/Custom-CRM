import test from 'node:test';
import assert from 'node:assert/strict';
import { createLogger, redact } from '../logger.js';
import { createStore, KEYS } from '../store.js';
import { memoryArea } from './_chrome.js';

const quiet = { log() {} };

test('redact removes secrets and email addresses', () => {
  const t = redact('key=abcd1234 mail bob@example.com', ['abcd1234']);
  assert.ok(!t.includes('abcd1234') && !t.includes('bob@example.com'));
});
test('log entries carry job, stage, ok, category, timestamp; failures land in errors', async () => {
  const store = createStore(memoryArea());
  const { log } = createLogger({ store, now: () => 42, console: quiet });
  await log({ jobId: 'j1', stage: 'claim', ok: true });
  await log({ jobId: 'j1', stage: 'report', ok: false, category: 'network', message: 'down' });
  const logs = await store.getList(KEYS.logs);
  assert.equal(logs.length, 2);
  assert.deepEqual(Object.keys(logs[0]).sort(), ['at', 'category', 'jobId', 'message', 'ok', 'stage']);
  assert.equal((await store.getList(KEYS.errors)).length, 1);
});
test('log ring buffer is bounded', async () => {
  const store = createStore(memoryArea());
  const { log } = createLogger({ store, console: quiet });
  for (let i = 0; i < 130; i++) await log({ stage: 's' });
  assert.equal((await store.getList(KEYS.logs)).length, 100);
});
test('logging never throws if storage fails', async () => {
  const bad = { getList: async () => { throw new Error('x'); } };
  const { log } = createLogger({ store: bad, console: quiet });
  await log({ stage: 's' });
});
