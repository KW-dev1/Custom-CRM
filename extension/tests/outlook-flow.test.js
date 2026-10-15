import test from 'node:test';
import assert from 'node:assert/strict';
import { createOutlook, newGuard } from './fixtures/outlook.js';

const job = (o = {}) => ({ id: 'job_1', to: 'a@example.com', subject: 'Hello there', body: 'Line 1\n\nLine 3', ...o });
const FAST = { ready: 300, compose: 500, recipient: 400, verify: 500 };

function run(o, j = job(), extra = {}) {
  const stages = [];
  const guard = extra.guard || newGuard(o.window);
  const p = o.OEA.processJob(j, {
    doc: o.doc, loc: o.loc, guard, timeouts: FAST,
    markStage: async (id, s) => { stages.push(s); if (extra.failStage === s) throw new Error('storage down'); },
  });
  return p.then((r) => ({ r, stages, guard }));
}

test('happy path: composes, fills correctly, clicks Send once, verifies close', async () => {
  const o = await createOutlook();
  const { r, stages } = await run(o);
  assert.equal(r.outcome, 'sent');
  assert.equal(r.sendClicked, true);
  assert.match(r.evidence, /closed/);
  assert.equal(o.controls.sendClicks, 1);
  assert.equal(o.controls.newMailClicks, 1);
  assert.equal(o.controls.ccTouched, false, 'Cc must never be touched');
  assert.deepEqual(stages, ['started', 'checking_ready', 'opening_compose', 'filling', 'validating', 'send_clicking', 'verifying']);
});
test('"message sent" notification also counts as evidence', async () => {
  const { r } = await run(await createOutlook({ sendMode: 'toast' }));
  assert.equal(r.outcome, 'sent');
  assert.match(r.evidence, /notification/);
});
test('progress is persisted BEFORE the Send click (send_clicking precedes click)', async () => {
  const o = await createOutlook();
  let clicksAtMark = null;
  await o.OEA.processJob(job(), {
    doc: o.doc, loc: o.loc, guard: newGuard(o.window), timeouts: FAST,
    markStage: async (id, s) => { if (s === 'send_clicking') clicksAtMark = o.controls.sendClicks; },
  });
  assert.equal(clicksAtMark, 0);
});
test('if progress cannot be persisted before Send, nothing is sent', async () => {
  const o = await createOutlook();
  const { r } = await run(o, job(), { failStage: 'send_clicking' });
  assert.equal(r.outcome, 'failed');
  assert.equal(r.sendClicked, false);
  assert.equal(o.controls.sendClicks, 0);
});
test('no confirmation within the timeout => uncertain, and Send is NOT clicked again', async () => {
  const o = await createOutlook({ sendMode: 'never' });
  const { r } = await run(o);
  assert.equal(r.outcome, 'uncertain');
  assert.equal(r.sendClicked, true);
  assert.equal(r.code, 'NO_CONFIRMATION');
  assert.equal(o.controls.sendClicks, 1);
});
test('slow Outlook (closes after the timeout) is still reported uncertain, not sent, not resent', async () => {
  const o = await createOutlook({ sendMode: 'late-close' });
  const { r } = await run(o, job(), { guard: undefined });
  assert.equal(r.outcome, 'uncertain');
  assert.equal(o.controls.sendClicks, 1);
});
test('send error banner => uncertain with the banner text', async () => {
  const o = await createOutlook({ sendMode: 'error' });
  const { r } = await run(o);
  assert.equal(r.outcome, 'uncertain');
  assert.equal(r.code, 'SEND_ERROR_SHOWN');
  assert.equal(o.controls.sendClicks, 1);
});
test('same job id again in the same tab is refused (never double-sent)', async () => {
  const o = await createOutlook();
  const guard = newGuard(o.window);
  const first = await run(o, job(), { guard });
  assert.equal(first.r.outcome, 'sent');
  const second = await run(o, job(), { guard });
  assert.equal(second.r.outcome, 'uncertain');
  assert.equal(second.r.code, 'DUPLICATE_JOB');
  assert.equal(o.controls.sendClicks, 1);
});
test('concurrent processing of the same job id sends once', async () => {
  const o = await createOutlook();
  const guard = newGuard(o.window);
  const [a, b] = await Promise.all([run(o, job(), { guard }), run(o, job(), { guard })]);
  assert.equal(o.controls.sendClicks, 1);
  assert.deepEqual([a.r.outcome, b.r.outcome].sort(), ['sent', 'uncertain']);
});
test('a different job while one is running is rejected as BUSY (not sent)', async () => {
  const o = await createOutlook();
  const guard = newGuard(o.window);
  const [a, b] = await Promise.all([run(o, job(), { guard }), run(o, job({ id: 'job_2' }), { guard })]);
  assert.equal(b.r.code, 'BUSY');
  assert.equal(b.r.outcome, 'failed');
  assert.equal(a.r.outcome, 'sent');
});
test('recipient never resolves => failed before Send', async () => {
  const o = await createOutlook({ resolve: false });
  const { r } = await run(o);
  assert.equal(r.outcome, 'failed');
  assert.equal(r.sendClicked, false);
  assert.equal(r.code, 'RECIPIENT_FAILED');
  assert.equal(o.controls.sendClicks, 0);
});
test('Outlook rejects the recipient => failed before Send', async () => {
  const o = await createOutlook({ resolve: 'invalid' });
  const { r } = await run(o);
  assert.equal(r.code, 'RECIPIENT_REJECTED');
  assert.equal(o.controls.sendClicks, 0);
});
test('invalid / missing job fields never reach Outlook', async () => {
  for (const bad of [{ to: '' }, { to: 'x' }, { to: 'a@b.com,c@d.com' }, { subject: '' }, { subject: 'a\nb' }, { body: '  ' }, { id: '' }]) {
    const o = await createOutlook();
    const { r } = await run(o, job(bad));
    assert.equal(r.code, 'INVALID_JOB', JSON.stringify(bad));
    assert.equal(o.controls.newMailClicks, 0);
  }
});
test('an already-open draft aborts without touching it', async () => {
  const o = await createOutlook({ pre: 'compose' });
  const { r } = await run(o);
  assert.equal(r.code, 'COMPOSE_ALREADY_OPEN');
  assert.equal(o.controls.sendClicks, 0);
});
test('signed-out page fails fast with a clear code', async () => {
  const o = await createOutlook({ pre: 'signedOut' });
  const t0 = Date.now();
  const { r } = await run(o);
  assert.equal(r.code, 'SIGNED_OUT');
  assert.ok(Date.now() - t0 < 250);
});
test('Outlook never finishing loading => bounded timeout failure', async () => {
  const o = await createOutlook({ pre: 'noShell' });
  const { r } = await run(o);
  assert.equal(r.outcome, 'failed');
  assert.equal(r.code, 'TIMEOUT');
});
test('wrong page => failed with unexpected_page', async () => {
  const o = await createOutlook({ url: 'https://outlook.live.com/calendar/' });
  const { r } = await run(o);
  assert.equal(r.code, 'UNEXPECTED_PAGE');
});
test('Send button missing => compose never recognised => failed, nothing sent', async () => {
  const o = await createOutlook({ noSend: true });
  const { r } = await run(o);
  assert.equal(r.outcome, 'failed');
  assert.equal(o.controls.sendClicks, 0);
});
test('multiline body with HTML-looking text is inserted as text only', async () => {
  const o = await createOutlook({ sendMode: 'never' });
  await run(o, job({ body: 'a\n<script>alert(1)</script>\nb' }));
  assert.equal(o.doc.querySelector('#compose script'), null);
});
