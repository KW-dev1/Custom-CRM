import test, { before, after } from 'node:test';
import assert from 'node:assert/strict';
import { createApp } from '../src/app.js';
import { createJobService } from '../src/jobs.js';
import { SheetsError } from '../src/sheets-client.js';
import { FakeSheet, testConfig, lead } from './_fakes.js';

const KEY = 'abcdefghijklmnop1234';
let server; let base; let sheet; let clock;
const logs = [];

before(async () => {
  clock = { t: 5_000_000 };
  sheet = new FakeSheet([lead(1), lead(2, { emails: 'solo@x.com' }), lead(3)]);
  const cfg = testConfig();
  const service = createJobService({ sheet, cfg, now: () => clock.t, log: (...a) => logs.push(a) });
  const app = createApp({ service, cfg, now: () => clock.t, log: (...a) => logs.push(a) });
  await new Promise((r) => { server = app.listen(0, '127.0.0.1', r); });
  base = `http://127.0.0.1:${server.address().port}`;
});
after(() => server.close());

const call = (path, { method = 'GET', body, key = KEY, headers = {} } = {}) => fetch(base + path, {
  method, headers: { ...(key ? { authorization: `Bearer ${key}` } : {}), ...(body ? { 'content-type': 'application/json' } : {}), ...headers },
  body: body ? JSON.stringify(body) : undefined,
});

test('health needs no auth and never touches the sheet', async () => {
  const n = sheet.writes.length;
  const r = await call('/health', { key: null });
  assert.equal(r.status, 200);
  assert.deepEqual(await r.json(), { ok: true });
  assert.equal(sheet.writes.length, n);
});
test('API requires a valid key (Bearer or X-Api-Key)', async () => {
  assert.equal((await call('/api/email-jobs/claim', { method: 'POST', key: null })).status, 401);
  assert.equal((await call('/api/email-jobs/claim', { method: 'POST', key: 'wrong-wrong-wrong-wrong' })).status, 401);
  assert.equal(sheet.writes.length, 0, 'unauthenticated calls must not reach the sheet');
});
test('full lifecycle: claim -> processing -> sent -> state', async () => {
  const c = await call('/api/email-jobs/claim', { method: 'POST', body: {} });
  assert.equal(c.status, 200);
  const { job } = await c.json();
  assert.equal(job.to, 'lead1@x.com');
  assert.deepEqual(job.cc, ['cc1@y.org']);
  assert.equal(sheet.rows[0].status, 'pending');

  const p = await call(`/api/email-jobs/${job.id}/status`, { method: 'POST', body: { jobId: job.id, status: 'processing', timestamp: 'x' } });
  assert.equal(p.status, 200);
  const s = await call(`/api/email-jobs/${job.id}/status`, { method: 'POST', body: { jobId: job.id, status: 'sent', message: 'Verified: compose window closed after Send' } });
  assert.deepEqual(await s.json(), { id: job.id, status: 'sent', applied: true });
  assert.equal(sheet.rows[0].status, 'sent');
  const again = await call(`/api/email-jobs/${job.id}/status`, { method: 'POST', body: { status: 'sent' } });
  assert.equal((await again.json()).applied, false);
  assert.deepEqual(await (await call(`/api/email-jobs/${job.id}`)).json(), { id: job.id, status: 'sent' });
});
test('single-address row has an empty CC list', async () => {
  const { job } = await (await call('/api/email-jobs/claim', { method: 'POST', body: {} })).json();
  assert.equal(job.to, 'solo@x.com');
  assert.deepEqual(job.cc, []);
});
test('conflicting final status -> 409', async () => {
  const id = (await (await call('/api/email-jobs/claim', { method: 'POST', body: {} })).json()).job.id;
  await call(`/api/email-jobs/${id}/status`, { method: 'POST', body: { status: 'failed' } });
  const r = await call(`/api/email-jobs/${id}/status`, { method: 'POST', body: { status: 'sent' } });
  assert.equal(r.status, 409);
  assert.equal((await r.json()).error.code, 'CONFLICT');
});
test('empty queue -> 204', async () => {
  const r = await call('/api/email-jobs/claim', { method: 'POST', body: {} });
  assert.equal(r.status, 204);
});
test('input validation: bad id, bad status, id mismatch, bad JSON', async () => {
  assert.equal((await call('/api/email-jobs/nope/status', { method: 'POST', body: { status: 'sent' } })).status, 400);
  assert.equal((await call('/api/email-jobs/r2-aaaaaaaaaaaa/status', { method: 'POST', body: { status: 'delivered' } })).status, 400);
  assert.equal((await call('/api/email-jobs/r2-aaaaaaaaaaaa/status', { method: 'POST', body: { jobId: 'other', status: 'sent' } })).status, 400);
  const bad = await fetch(`${base}/api/email-jobs/r2-aaaaaaaaaaaa/status`, { method: 'POST', headers: { authorization: `Bearer ${KEY}`, 'content-type': 'application/json' }, body: '{oops' });
  assert.equal(bad.status, 400);
  assert.equal((await call('/api/email-jobs/r2-aaaaaaaaaaaa/status', { method: 'POST', body: { status: 'sent' } })).status, 404);
});
test('unknown endpoint -> JSON 404', async () => {
  const r = await call('/api/nothing');
  assert.equal(r.status, 404);
  assert.equal((await r.json()).error.code, 'NOT_FOUND');
});
test('Google/Sheets outage -> 502 without leaking internals', async () => {
  const bad = { claimNext: async () => { throw new SheetsError('Sheets API HTTP 503'); } };
  const app = createApp({ service: bad, cfg: testConfig(), log: () => {} });
  const s = await new Promise((r) => { const x = app.listen(0, '127.0.0.1', () => r(x)); });
  const r = await fetch(`http://127.0.0.1:${s.address().port}/api/email-jobs/claim`, { method: 'POST', headers: { authorization: `Bearer ${KEY}` } });
  assert.equal(r.status, 502);
  assert.equal((await r.json()).error.code, 'SHEET_UNAVAILABLE');
  s.close();
});
test('unexpected error -> generic 500', async () => {
  const app = createApp({ service: { claimNext: async () => { throw new Error('secret internals'); } }, cfg: testConfig(), log: () => {} });
  const s = await new Promise((r) => { const x = app.listen(0, '127.0.0.1', () => r(x)); });
  const r = await fetch(`http://127.0.0.1:${s.address().port}/api/email-jobs/claim`, { method: 'POST', headers: { authorization: `Bearer ${KEY}` } });
  assert.equal(r.status, 500);
  assert.ok(!JSON.stringify(await r.json()).includes('secret'));
  s.close();
});
test('rate floor surfaces as 429 + Retry-After', async () => {
  const sh = new FakeSheet([lead(1), lead(2)]);
  const cfg = testConfig({ MIN_CLAIM_INTERVAL_SEC: '600' });
  const app = createApp({ service: createJobService({ sheet: sh, cfg, now: () => 1000 }), cfg, log: () => {} });
  const s = await new Promise((r) => { const x = app.listen(0, '127.0.0.1', () => r(x)); });
  const u = `http://127.0.0.1:${s.address().port}/api/email-jobs/claim`;
  const h = { method: 'POST', headers: { authorization: `Bearer ${KEY}` } };
  assert.equal((await fetch(u, h)).status, 200);
  const r = await fetch(u, h);
  assert.equal(r.status, 429);
  assert.equal(r.headers.get('retry-after'), '600');
  s.close();
});
test('repeated bad keys are throttled', async () => {
  for (let i = 0; i < 10; i++) await call('/api/email-jobs/x', { key: 'bad-bad-bad-bad-bad' });
  const r = await call('/api/email-jobs/x', { key: 'bad-bad-bad-bad-bad' });
  assert.equal(r.status, 429);
  clock.t += 61_000;
  assert.equal((await call('/api/email-jobs/x', { key: 'bad-bad-bad-bad-bad' })).status, 401);
});
test('logs never contain email addresses or message bodies', () => {
  const text = JSON.stringify(logs);
  assert.ok(logs.length > 0, 'expected some log lines to have been captured');
  assert.ok(!text.includes('Body 1') && !text.includes('Subject 1'));
  assert.ok(!/lead\d@x\.com|cc\d@y\.org|solo@x\.com/.test(text), 'no email addresses in logs');
});

test('X-Client-Id: valid names are recorded, bad ones are rejected, other profiles cannot report', async () => {
  const sh = new FakeSheet([lead(1), lead(2)]);
  const cfg = testConfig();
  const app = createApp({ service: createJobService({ sheet: sh, cfg }), cfg, log: () => {} });
  const s = await new Promise((r) => { const x = app.listen(0, '127.0.0.1', () => r(x)); });
  const u = `http://127.0.0.1:${s.address().port}/api/email-jobs`;
  const h = (client) => ({ authorization: `Bearer ${KEY}`, 'content-type': 'application/json', ...(client ? { 'x-client-id': client } : {}) });
  const claim = await fetch(`${u}/claim`, { method: 'POST', headers: h('outlook-A'), body: '{}' });
  const { job } = await claim.json();
  assert.equal(sh.rows[0].sender, 'outlook-A');
  assert.equal((await fetch(`${u}/claim`, { method: 'POST', headers: h('bad@name.com'), body: '{}' })).status, 400);
  assert.equal((await fetch(`${u}/claim`, { method: 'POST', headers: h('x'.repeat(41)), body: '{}' })).status, 400);
  const wrong = await fetch(`${u}/${job.id}/status`, { method: 'POST', headers: h('outlook-B'), body: JSON.stringify({ status: 'sent' }) });
  assert.equal(wrong.status, 409);
  assert.equal((await wrong.json()).error.code, 'WRONG_CLIENT');
  const ok = await fetch(`${u}/${job.id}/status`, { method: 'POST', headers: h('outlook-A'), body: JSON.stringify({ status: 'sent' }) });
  assert.equal(ok.status, 200);
  s.close();
});
