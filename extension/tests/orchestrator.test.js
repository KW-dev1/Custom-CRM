import test from 'node:test';
import assert from 'node:assert/strict';
import { createOrchestrator } from '../orchestrator.js';
import { createStore } from '../store.js';
import { ApiError, validateJob } from '../api.js';
import { memoryArea } from './_chrome.js';
import { goodJob } from './_helpers.js';

const quiet = { log: async () => {} };

function setup({ rand = 0, job = goodJob(), sendResult, sendError, duringSend = null, state = null, enabled = true, running = true, stored = {} } = {}) {
  const area = memoryArea({
    settings: { apiBaseUrl: 'https://api.example.com', apiKey: 'k-12345', enabled, fetchMinMinutes: 10, fetchMaxMinutes: 20 },
    runtime: { running, consecutiveFailures: 0, lastFinishedAt: 0, nextClaimAt: 0 },
    ...stored,
  });
  const store = createStore(area);
  const calls = { claim: 0, send: [], reports: [], schedule: [] };
  let clock = 1_000_000;
  const env = {
    area, store, calls,
    advance: (ms) => { clock += ms; },
    reportFails: 0,
    claimError: null,
    tab: { id: 7 },
    ready: { state: 'ready' },
  };
  const api = {
    claimJob: async () => { calls.claim++; if (env.claimError) throw env.claimError; return job; },
    getJobState: async () => state,
    reportStatus: async (_s, r) => {
      if (env.reportFails > 0) { env.reportFails--; throw new ApiError('network', 'down'); }
      calls.reports.push(r);
    },
    validateJob,
  };
  const tabs = {
    findOutlookTab: async () => env.tab,
    send: async (_id, msg) => {
      if (msg.type === 'OEA_CHECK_READY') return env.ready;
      calls.send.push(msg);
      if (duringSend) await area.set({ inflight: duringSend });
      if (sendError) throw sendError;
      return sendResult;
    },
  };
  env.orch = createOrchestrator({
    store, api, tabs, logger: quiet, now: () => clock, scheduleNext: (s) => calls.schedule.push(s), random: () => rand,
    options: { staleMs: 120000 },
  });
  return env;
}
const statuses = (e) => e.calls.reports.map((r) => r.status);

test('not configured', async () => {
  const e = setup(); await e.area.set({ settings: {} });
  assert.equal((await e.orch.tick()).state, 'not_configured');
  assert.equal((await e.store.getStatus()).code, 'not_configured');
});
test('paused when disabled or stopped; never claims', async () => {
  for (const o of [{ enabled: false }, { running: false }]) {
    const e = setup(o);
    assert.equal((await e.orch.tick()).state, 'paused');
    assert.equal(e.calls.claim, 0);
    assert.equal((await e.store.getStatus()).code, 'paused');
  }
});
test('no Outlook tab -> waiting_outlook, nothing claimed', async () => {
  const e = setup(); e.tab = null;
  assert.equal((await e.orch.tick()).state, 'no_tab');
  assert.equal((await e.store.getStatus()).code, 'waiting_outlook');
  assert.equal(e.calls.claim, 0);
});
test('Outlook signed out / loading -> waiting_outlook, nothing claimed', async () => {
  for (const s of ['signed_out', 'loading', 'unexpected_page']) {
    const e = setup(); e.ready = { state: s };
    assert.equal((await e.orch.tick()).state, 'outlook_not_ready');
    assert.equal(e.calls.claim, 0);
  }
});
test('no job -> ready', async () => {
  const e = setup({ job: null });
  assert.equal((await e.orch.tick()).state, 'idle');
  assert.equal((await e.store.getStatus()).code, 'ready');
});

test('happy path: verified send reports processing then sent, records locally', async () => {
  const e = setup({ sendResult: { outcome: 'sent', sendClicked: true, evidence: 'compose window closed' } });
  const r = await e.orch.tick();
  assert.deepEqual(r, { state: 'processed', status: 'sent' });
  assert.deepEqual(statuses(e), ['processing', 'sent']);
  assert.equal(e.calls.send.filter((m) => m.type === 'OEA_SEND_JOB').length, 1);
  assert.equal((await e.store.getCompletedMap()).job_1.status, 'sent');
  assert.equal((await e.store.getStatus()).code, 'sent');
  assert.equal(await e.store.getRun(), null);
  assert.equal(await e.store.getPendingReport(), null);
});
test('uncertain outcome -> needs_review, automation paused, status is NOT sent', async () => {
  const e = setup({ sendResult: { outcome: 'uncertain', sendClicked: true, message: 'no confirmation' } });
  await e.orch.tick();
  assert.deepEqual(statuses(e), ['processing', 'needs_review']);
  assert.equal((await e.store.getStatus()).code, 'needs_review');
  assert.equal((await e.store.getRuntime()).running, false);
});
test('pre-send failure -> failed; pauses only after repeated failures', async () => {
  const e = setup({ sendResult: { outcome: 'failed', sendClicked: false, message: 'Recipient not accepted' } });
  await e.orch.tick();
  assert.deepEqual(statuses(e), ['processing', 'failed']);
  assert.equal((await e.store.getRuntime()).running, true);
  for (let i = 0; i < 2; i++) { e.advance(11 * 60000); await e.orch.tick(); }
  assert.equal((await e.store.getRuntime()).running, false);
});
test('a "failed" result with sendClicked=true is downgraded to needs_review', async () => {
  const e = setup({ sendResult: { outcome: 'failed', sendClicked: true, message: 'x' } });
  await e.orch.tick();
  assert.equal(statuses(e).at(-1), 'needs_review');
});
test('garbage result from the page -> needs_review', async () => {
  const e = setup({ sendResult: undefined });
  await e.orch.tick();
  assert.equal(statuses(e).at(-1), 'needs_review');
});

test('invalid job is reported failed and never sent', async () => {
  const e = setup({ job: goodJob({ to: 'nope' }) });
  await e.orch.tick();
  assert.equal(e.calls.send.filter((m) => m.type === 'OEA_SEND_JOB').length, 0);
  assert.equal(statuses(e).at(-1), 'failed');
  assert.match(e.calls.reports.at(-1).message, /Invalid job data/);
});
test('duplicate delivery of a locally completed job is not re-sent', async () => {
  const e = setup({ sendResult: { outcome: 'sent', evidence: 'x' }, stored: { completed: { job_1: { status: 'sent', at: 1 } } } });
  const r = await e.orch.tick();
  assert.equal(r.state, 'duplicate');
  assert.equal(e.calls.send.filter((m) => m.type === 'OEA_SEND_JOB').length, 0);
  assert.equal(e.calls.reports.at(-1).status, 'sent');
});
test('backend state check: already sent -> skipped, not sent again', async () => {
  const e = setup({ state: 'sent', sendResult: { outcome: 'sent' } });
  assert.equal((await e.orch.tick()).state, 'backend_final');
  assert.equal(e.calls.send.filter((m) => m.type === 'OEA_SEND_JOB').length, 0);
  assert.equal(e.calls.reports.length, 0);
});

test('channel closed after Send was clicked -> needs_review, never retried', async () => {
  const e = setup({ sendError: new Error('The message channel closed'), duringSend: { jobId: 'job_1', stage: 'verifying', at: 1 } });
  await e.orch.tick();
  assert.equal(e.calls.reports.at(-1).status, 'needs_review');
  assert.equal(e.calls.send.filter((m) => m.type === 'OEA_SEND_JOB').length, 1);
});
test('channel closed before Send was clicked -> failed (provably not sent)', async () => {
  const e = setup({ sendError: new Error('The message channel closed'), duringSend: { jobId: 'job_1', stage: 'filling', at: 1 } });
  await e.orch.tick();
  assert.equal(e.calls.reports.at(-1).status, 'failed');
});
test('no receiver and no progress record -> failed (provably not sent)', async () => {
  const e = setup({ sendError: new Error('Could not establish connection. Receiving end does not exist.') });
  await e.orch.tick();
  assert.equal(e.calls.reports.at(-1).status, 'failed');
});
test('unknown error with no progress record -> needs_review', async () => {
  const e = setup({ sendError: new Error('boom') });
  await e.orch.tick();
  assert.equal(e.calls.reports.at(-1).status, 'needs_review');
});

test('report failure keeps the result pending; next tick flushes before claiming', async () => {
  const e = setup({ sendResult: { outcome: 'sent', evidence: 'x' } });
  e.reportFails = 100;
  await e.orch.tick();
  const pending = await e.store.getPendingReport();
  assert.equal(pending.status, 'sent');
  assert.equal(e.calls.claim, 1);
  e.reportFails = 0; e.advance(5000);
  const r = await e.orch.tick();
  assert.equal(e.calls.claim, 1, 'must not claim a new job before the pending report is delivered');
  assert.equal(e.calls.reports.at(-1).status, 'sent');
  assert.equal(await e.store.getPendingReport(), null);
  assert.equal(r.state, 'waiting');
});
test('report failure never causes a second send', async () => {
  const e = setup({ sendResult: { outcome: 'sent', evidence: 'x' } });
  e.reportFails = 100;
  await e.orch.tick(); e.advance(11 * 60000); await e.orch.tick(); e.advance(11 * 60000); await e.orch.tick();
  assert.equal(e.calls.send.filter((m) => m.type === 'OEA_SEND_JOB').length, 1);
});

test('crash recovery: stale run after Send attempt -> needs_review, no resend', async () => {
  const e = setup({ sendResult: { outcome: 'sent' }, stored: {
    run: { jobId: 'job_1', phase: 'sending', startedAt: 1, updatedAt: 1 },
    inflight: { jobId: 'job_1', stage: 'verifying', at: 1 },
  } });
  const r = await e.orch.tick();
  assert.equal(r.state, 'recovered');
  assert.equal(e.calls.reports.at(-1).status, 'needs_review');
  assert.equal(e.calls.send.length, 0);
  assert.equal(e.calls.claim, 0);
});
test('crash recovery: persisted final result from the page is used', async () => {
  const e = setup({ stored: {
    run: { jobId: 'job_1', phase: 'sending', startedAt: 1, updatedAt: 1 },
    inflight: { jobId: 'job_1', stage: 'done', at: 1, result: { outcome: 'sent', evidence: 'compose window closed' } },
  } });
  await e.orch.tick();
  assert.equal(e.calls.reports.at(-1).status, 'sent');
});
test('a fresh run is left alone (still in progress)', async () => {
  const e = setup({ stored: { run: { jobId: 'job_1', phase: 'sending', startedAt: 1_000_000, updatedAt: 1_000_000 } } });
  assert.equal((await e.orch.tick()).state, 'in_progress');
  assert.equal(e.calls.claim, 0);
});

test('auth failure on claim pauses automation', async () => {
  const e = setup(); e.claimError = new ApiError('auth', 'Backend rejected the credentials (HTTP 401).');
  await e.orch.tick();
  assert.equal((await e.store.getRuntime()).running, false);
  assert.equal((await e.store.getStatus()).code, 'failed');
});
test('network error on claim keeps running and reschedules; no job is sent', async () => {
  const e = setup(); e.claimError = new ApiError('network', 'Could not reach the backend');
  await e.orch.tick();
  assert.equal((await e.store.getRuntime()).running, true);
  assert.ok(e.calls.schedule.length > 0);
});
test('rate limit honours Retry-After', async () => {
  const e = setup(); e.claimError = new ApiError('rate_limit', 'slow down', { retryAfterMs: 90000 });
  await e.orch.tick();
  assert.equal(e.calls.schedule.at(-1), 90);
});
test('no second fetch until the random gap has passed', async () => {
  const e = setup({ sendResult: { outcome: 'sent', evidence: 'x' } });
  await e.orch.tick();
  e.advance(5000);
  assert.equal((await e.orch.tick()).state, 'waiting');
  assert.equal(e.calls.claim, 1);
});
test('concurrent ticks do not double-process', async () => {
  const e = setup({ sendResult: { outcome: 'sent', evidence: 'x' } });
  const [a, b] = await Promise.all([e.orch.tick(), e.orch.tick()]);
  assert.ok([a.state, b.state].includes('busy'));
  assert.equal(e.calls.send.filter((m) => m.type === 'OEA_SEND_JOB').length, 1);
});

const MIN = 60000;
test('fetch gap is random within [min,max] minutes and persisted as nextClaimAt', async () => {
  for (const [rand, minutes] of [[0, 10], [1, 20], [0.5, 15]]) {
    const e = setup({ job: null, rand });
    await e.orch.tick();
    const rt = await e.store.getRuntime();
    assert.equal(rt.nextClaimAt - 1_000_000, minutes * MIN);
    assert.equal(e.calls.schedule.at(-1), minutes * 60);
  }
});
test('gap is applied after a fetch that returned a job too, and no earlier fetch happens', async () => {
  const e = setup({ rand: 1, sendResult: { outcome: 'sent', evidence: 'x' } });
  await e.orch.tick();
  e.advance(19 * MIN);
  assert.equal((await e.orch.tick()).state, 'waiting');
  assert.equal(e.calls.claim, 1);
  e.advance(2 * MIN);
  await e.orch.tick();
  assert.equal(e.calls.claim, 2);
});
test('ticks inside the gap never touch the backend claim endpoint (alarm storm safe)', async () => {
  const e = setup({ job: null });
  await e.orch.tick();
  for (let i = 0; i < 25; i++) { e.advance(20000); await e.orch.tick(); }
  assert.equal(e.calls.claim, 1);
});
test('Start resets the gap so the first fetch is immediate', async () => {
  const e = setup({ job: null, running: false });
  await e.store.patchRuntime({ nextClaimAt: 9_999_999_999 });
  await e.orch.start();
  assert.equal(e.calls.claim, 1);
});
test('a claim error retries soon (not after a full 10-20 min window)', async () => {
  const e = setup(); e.claimError = new ApiError('network', 'down');
  await e.orch.tick();
  const rt = await e.store.getRuntime();
  assert.ok(rt.nextClaimAt - 1_000_000 <= 60000);
});

test('a job that previously FAILED (never sent) is retried when it is claimed again', async () => {
  const e = setup({ sendResult: { outcome: 'sent', sendClicked: true, evidence: 'compose window closed' },
    stored: { completed: { job_1: { status: 'failed', at: 1 } } } });
  const r = await e.orch.tick();
  assert.deepEqual(r, { state: 'processed', status: 'sent' });
  assert.equal(e.calls.send.filter((m) => m.type === 'OEA_SEND_JOB').length, 1);
  assert.equal((await e.store.getCompletedMap()).job_1.status, 'sent', 'the record is updated to the new outcome');
});
test('previously SENT or NEEDS_REVIEW jobs are still blocked from being sent again', async () => {
  for (const status of ['sent', 'needs_review']) {
    const e = setup({ sendResult: { outcome: 'sent', evidence: 'x' }, stored: { completed: { job_1: { status, at: 1 } } } });
    assert.equal((await e.orch.tick()).state, 'duplicate', status);
    assert.equal(e.calls.send.filter((m) => m.type === 'OEA_SEND_JOB').length, 0, status);
    assert.equal(e.calls.reports.at(-1).status, status);
  }
});
test('retry after failure that fails again keeps the record as failed', async () => {
  const e = setup({ sendResult: { outcome: 'failed', sendClicked: false, message: 'still no recipient' },
    stored: { completed: { job_1: { status: 'failed', at: 1 } } } });
  await e.orch.tick();
  assert.equal((await e.store.getCompletedMap()).job_1.status, 'failed');
  assert.equal(e.calls.reports.at(-1).status, 'failed');
});
