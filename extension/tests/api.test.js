import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildClaimRequest, buildStatusRequest, buildJobStateRequest, parseJobResponse, validateJob,
  claimJob, reportStatus, getJobState, ApiError,
} from '../api.js';
import { goodSettings, goodJob, fakeResponse } from './_helpers.js';

test('claim request uses configured URL, method and auth header', () => {
  const r = buildClaimRequest(goodSettings());
  assert.equal(r.url, 'https://api.example.com/api/email-jobs/claim');
  assert.equal(r.init.method, 'POST');
  assert.equal(r.init.headers.Authorization, 'Bearer secret-key-123');
});
test('custom header and empty scheme send the raw key', () => {
  const r = buildClaimRequest({ ...goodSettings(), authHeader: 'X-Api-Key', authScheme: '' });
  assert.equal(r.init.headers['X-Api-Key'], 'secret-key-123');
  assert.equal(r.init.headers.Authorization, undefined);
});
test('status request maps the vocabulary and encodes the id', () => {
  const r = buildStatusRequest(
    { ...goodSettings(), statusValues: { sent: 'DELIVERED' } },
    { jobId: 'a/b c', status: 'sent', timestamp: 'T' },
  );
  assert.equal(r.url, 'https://api.example.com/api/email-jobs/a%2Fb%20c/status');
  assert.deepEqual(JSON.parse(r.init.body), { jobId: 'a/b c', status: 'DELIVERED', timestamp: 'T' });
});
test('status request never carries the API key in the body', () => {
  const r = buildStatusRequest(goodSettings(), { jobId: '1', status: 'failed', message: 'x' });
  assert.ok(!r.init.body.includes('secret-key-123'));
});
test('job state request', () => {
  const r = buildJobStateRequest({ ...goodSettings(), jobStatePath: '/jobs/{id}' }, '7');
  assert.equal(r.url, 'https://api.example.com/jobs/7');
});

test('parseJobResponse: wrapped, bare, empty', () => {
  assert.equal(parseJobResponse({ job: goodJob() }).id, 'job_1');
  assert.equal(parseJobResponse(goodJob()).to, 'a@example.com');
  assert.equal(parseJobResponse({ job: null }), null);
  assert.equal(parseJobResponse(null), null);
  assert.equal(parseJobResponse({}), null);
});
test('parseJobResponse rejects malformed shapes', () => {
  assert.throws(() => parseJobResponse('hi'), ApiError);
  assert.throws(() => parseJobResponse({ job: 'x' }), /malformed/);
  assert.throws(() => parseJobResponse([1]), ApiError);
});
test('validateJob accepts a good job', () => assert.equal(validateJob(goodJob()).ok, true));
test('validateJob rejects missing/invalid fields', () => {
  for (const bad of [
    { id: '' }, { to: '' }, { to: 'nope' }, { to: 'a@b.com, c@d.com' }, { to: 'Bob <a@b.com>' },
    { subject: '' }, { subject: 'a\nb' }, { body: '   ' }, { body: 5 },
  ]) {
    assert.equal(validateJob(goodJob(bad)).ok, false, JSON.stringify(bad));
  }
  assert.equal(validateJob(null).ok, false);
});

test('claimJob: 200, 204 and no-work', async () => {
  assert.equal((await claimJob(goodSettings(), { fetchImpl: async () => fakeResponse(200, { job: goodJob() }) })).id, 'job_1');
  assert.equal(await claimJob(goodSettings(), { fetchImpl: async () => fakeResponse(204) }), null);
});
test('claimJob maps HTTP and network errors, never retries', async () => {
  const cases = [[401, 'auth'], [403, 'auth'], [429, 'rate_limit'], [500, 'server'], [400, 'client']];
  for (const [status, kind] of cases) {
    let calls = 0;
    await assert.rejects(
      claimJob(goodSettings(), { fetchImpl: async () => { calls++; return fakeResponse(status, {}); } }),
      (e) => e.kind === kind,
    );
    assert.equal(calls, 1);
  }
  await assert.rejects(claimJob(goodSettings(), { fetchImpl: async () => { throw new TypeError('x'); } }), (e) => e.kind === 'network');
});
test('invalid JSON from the backend', async () => {
  await assert.rejects(claimJob(goodSettings(), { fetchImpl: async () => fakeResponse(200, '<html>') }), (e) => e.kind === 'invalid_response');
});
test('error messages never contain the API key', async () => {
  try { await claimJob(goodSettings(), { fetchImpl: async () => fakeResponse(401, {}) }); } catch (e) {
    assert.ok(!e.message.includes('secret-key-123'));
  }
});

test('reportStatus retries transient failures then succeeds', async () => {
  const seq = [500, 429, 200];
  const waits = [];
  await reportStatus(goodSettings(), { jobId: '1', status: 'sent' }, {
    fetchImpl: async () => fakeResponse(seq.shift(), {}, { 'retry-after': '2' }),
    sleep: async (ms) => waits.push(ms),
  });
  assert.equal(seq.length, 0);
  assert.equal(waits.length, 2);
});
test('reportStatus gives up after bounded retries', async () => {
  let calls = 0;
  await assert.rejects(reportStatus(goodSettings(), { jobId: '1', status: 'sent' }, {
    fetchImpl: async () => { calls++; return fakeResponse(503, {}); }, sleep: async () => {},
  }), (e) => e.kind === 'server');
  assert.equal(calls, 4);
});
test('reportStatus does not retry auth/client errors', async () => {
  let calls = 0;
  await assert.rejects(reportStatus(goodSettings(), { jobId: '1', status: 'sent' }, {
    fetchImpl: async () => { calls++; return fakeResponse(401, {}); }, sleep: async () => {},
  }), (e) => e.kind === 'auth');
  assert.equal(calls, 1);
});
test('getJobState returns null when not configured, status when configured', async () => {
  assert.equal(await getJobState(goodSettings(), '1', { fetchImpl: async () => { throw new Error('no'); } }), null);
  const s = { ...goodSettings(), jobStatePath: '/jobs/{id}' };
  assert.equal(await getJobState(s, '1', { fetchImpl: async () => fakeResponse(200, { status: 'sent' }) }), 'sent');
});

test('parseJobResponse keeps cc and defaults it to []', () => {
  assert.deepEqual(parseJobResponse({ job: goodJob({ cc: ['b@example.com'] }) }).cc, ['b@example.com']);
  assert.deepEqual(parseJobResponse({ job: goodJob() }).cc, []);
});
test('validateJob: CC rules', () => {
  assert.equal(validateJob(goodJob({ cc: ['b@example.com', 'c@example.com'] })).ok, true);
  for (const cc of [['nope'], 'b@example.com', [5], ['a@example.com'], ['b@example.com', 'B@EXAMPLE.com'], Array.from({ length: 21 }, (_, i) => `u${i}@example.com`)]) {
    assert.equal(validateJob(goodJob({ cc })).ok, false, JSON.stringify(cc).slice(0, 40));
  }
});

test('profile name is sent as X-Client-Id on every request type; absent when not set', () => {
  const s = { ...goodSettings(), profileName: 'outlook-A', jobStatePath: '/jobs/{id}' };
  assert.equal(buildClaimRequest(s).init.headers['X-Client-Id'], 'outlook-A');
  assert.equal(buildStatusRequest(s, { jobId: '1', status: 'sent' }).init.headers['X-Client-Id'], 'outlook-A');
  assert.equal(buildJobStateRequest(s, '1').init.headers['X-Client-Id'], 'outlook-A');
  assert.equal(buildClaimRequest(goodSettings()).init.headers['X-Client-Id'], undefined);
});
