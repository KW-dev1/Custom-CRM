import test from 'node:test';
import assert from 'node:assert/strict';
import { createJobService, JobError } from '../src/jobs.js';
import { FakeSheet, testConfig, lead } from './_fakes.js';
import { makeJobId, rowHash } from '../src/rows.js';

const setup = (rows, cfgExtra, clock = { t: 1_000_000 }) => {
  const sheet = new FakeSheet(rows);
  const svc = createJobService({ sheet, cfg: testConfig(cfgExtra), now: () => clock.t });
  return { sheet, svc, clock };
};

test('claims the first row with an empty status: marks pending, first email TO, others CC', async () => {
  const { sheet, svc } = setup([lead(1), lead(2)]);
  const job = await svc.claimNext();
  assert.equal(job.to, 'lead1@x.com');
  assert.deepEqual(job.cc, ['cc1@y.org']);
  assert.equal(job.subject, 'Subject 1');
  assert.equal(job.body, 'Body 1');
  assert.equal(sheet.rows[0].status, 'pending');
  assert.equal(sheet.rows[1].status, '');
  assert.match(sheet.rows[0].note, /claimed/);
});
test('rows that already have any status, and blank rows, are never claimed', async () => {
  const { sheet, svc } = setup([lead(1, { status: 'sent' }), lead(2, { status: 'pending' }), lead(3, { status: 'failed' }), lead(4, { status: 'whatever' }), {}, lead(6)]);
  const job = await svc.claimNext();
  assert.equal(job.metadata.row, 7);
  assert.equal(sheet.writes.length, 1);
});
test('nothing to do -> null, and nothing is written', async () => {
  const { sheet, svc } = setup([lead(1, { status: 'sent' }), {}]);
  assert.equal(await svc.claimNext(), null);
  assert.equal(sheet.writes.length, 0);
});
test('a row is never handed out twice', async () => {
  const { svc } = setup([lead(1), lead(2)]);
  const a = await svc.claimNext(); const b = await svc.claimNext(); const c = await svc.claimNext();
  assert.notEqual(a.id, b.id);
  assert.equal(c, null);
});
test('concurrent claims are serialised and get different rows', async () => {
  const { sheet, svc } = setup([lead(1), lead(2), lead(3)]);
  sheet.delay = 5;
  const jobs = await Promise.all([svc.claimNext(), svc.claimNext(), svc.claimNext()]);
  assert.equal(new Set(jobs.map((j) => j.id)).size, 3);
});
test('invalid rows are marked failed (with reason) and the next valid row is claimed', async () => {
  const { sheet, svc } = setup([{ emails: 'bad@', subject: 'S', body: 'B' }, { emails: 'a@x.com', subject: '', body: 'B' }, lead(3)]);
  const job = await svc.claimNext();
  assert.equal(job.metadata.row, 4);
  assert.equal(sheet.rows[0].status, 'failed');
  assert.match(sheet.rows[0].note, /Invalid email/);
  assert.equal(sheet.rows[1].status, 'failed');
});
test('minimum claim interval is enforced with Retry-After', async () => {
  const { svc, clock } = setup([lead(1), lead(2)], { MIN_CLAIM_INTERVAL_SEC: '300' });
  await svc.claimNext();
  clock.t += 100_000;
  await assert.rejects(svc.claimNext(), (e) => e instanceof JobError && e.status === 429 && e.retryAfterSec === 200);
  clock.t += 200_000;
  assert.ok(await svc.claimNext());
});
test('if the pending marker cannot be written, no job is returned', async () => {
  const { sheet, svc } = setup([lead(1)]);
  sheet.failWrites = true;
  await assert.rejects(svc.claimNext(), /sheet write failed/);
  assert.equal(sheet.rows[0].status, '');
});
test('custom status labels', async () => {
  const { sheet, svc } = setup([lead(1)], { STATUS_PENDING: 'in-progress', STATUS_SENT: 'done' });
  const job = await svc.claimNext();
  assert.equal(sheet.rows[0].status, 'in-progress');
  await svc.report(job.id, 'sent');
  assert.equal(sheet.rows[0].status, 'done');
});

test('report: pending -> sent writes the status and a note', async () => {
  const { sheet, svc } = setup([lead(1)]);
  const job = await svc.claimNext();
  const r = await svc.report(job.id, 'sent', 'Verified: compose window closed');
  assert.deepEqual(r, { status: 'sent', applied: true, row: 2 });
  assert.equal(sheet.rows[0].status, 'sent');
  assert.match(sheet.rows[0].note, /Verified/);
});
test('report is idempotent for the same status (no second write)', async () => {
  const { sheet, svc } = setup([lead(1)]);
  const job = await svc.claimNext();
  await svc.report(job.id, 'sent');
  const n = sheet.writes.length;
  assert.equal((await svc.report(job.id, 'sent')).applied, false);
  assert.equal(sheet.writes.length, n);
});
test('a final status can never be overwritten by a different one', async () => {
  const { sheet, svc } = setup([lead(1)]);
  const job = await svc.claimNext();
  await svc.report(job.id, 'sent');
  for (const s of ['failed', 'needs_review']) {
    await assert.rejects(svc.report(job.id, s), (e) => e.status === 409 && e.code === 'CONFLICT');
  }
  assert.equal(sheet.rows[0].status, 'sent');
});
test('"processing" is a no-op on a pending row, never downgrades a final row, and requires a claim', async () => {
  const { sheet, svc } = setup([lead(1), lead(2)]);
  const job = await svc.claimNext();
  assert.equal((await svc.report(job.id, 'processing')).applied, false);
  await svc.report(job.id, 'failed');
  assert.equal((await svc.report(job.id, 'processing')).status, 'failed');
  assert.equal(sheet.rows[0].status, 'failed');
  const unclaimed = makeJobId(3, rowHash(sheet.rows[1]));
  await assert.rejects(svc.report(unclaimed, 'processing'), (e) => e.code === 'NOT_CLAIMED');
});
test('a result for an unclaimed (or manually cleared) row is still recorded', async () => {
  const { sheet, svc } = setup([lead(1)]);
  const id = makeJobId(2, rowHash(sheet.rows[0]));
  assert.equal((await svc.report(id, 'sent')).applied, true);
});
test('a manually typed status is not overwritten', async () => {
  const { sheet, svc } = setup([lead(1)]);
  const job = await svc.claimNext();
  sheet.rows[0].status = 'on hold';
  await assert.rejects(svc.report(job.id, 'sent'), (e) => e.code === 'CONFLICT');
});
test('malformed ids / unknown statuses are rejected', async () => {
  const { svc } = setup([lead(1)]);
  await assert.rejects(svc.report('nope', 'sent'), (e) => e.code === 'BAD_JOB_ID');
  const job = await svc.claimNext();
  await assert.rejects(svc.report(job.id, 'delivered'), (e) => e.code === 'BAD_STATUS');
});
test('edited row (content hash differs) -> 404, nothing written', async () => {
  const { sheet, svc } = setup([lead(1)]);
  const job = await svc.claimNext();
  sheet.rows[0].subject = 'Edited after claim';
  const n = sheet.writes.length;
  await assert.rejects(svc.report(job.id, 'sent'), (e) => e.status === 404);
  assert.equal(sheet.writes.length, n);
});
test('row moved (sheet sorted): found by content and the right row is updated', async () => {
  const { sheet, svc } = setup([lead(1), lead(2), lead(3)]);
  const job = await svc.claimNext();
  const moved = sheet.rows.shift();
  sheet.rows.push(moved);
  sheet.rows.forEach((r, i) => { r.rowNumber = i + 2; });
  const res = await svc.report(job.id, 'sent');
  assert.equal(res.row, 4);
  assert.equal(sheet.rows[2].status, 'sent');
  assert.equal(sheet.rows[0].status, '');
});
test('getState reflects the lifecycle', async () => {
  const { svc } = setup([lead(1)]);
  const job = await svc.claimNext();
  assert.equal((await svc.getState(job.id)).status, 'pending');
  await svc.report(job.id, 'needs_review');
  assert.equal((await svc.getState(job.id)).status, 'needs_review');
});
