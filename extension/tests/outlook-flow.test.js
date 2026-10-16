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

// ---- CC support: first address is To, the rest are Cc ----
import { wellOf } from './fixtures/outlook.js';

test('CC: addresses go to the Cc well, To address stays in To; Send clicked once', async () => {
  const o = await createOutlook();
  const { r } = await run(o, job({ cc: ['b@example.com', 'c@example.com'] }));
  assert.equal(r.outcome, 'sent');
  assert.equal(o.controls.ccToggleClicks, 1, 'Cc field revealed via its toggle');
  assert.equal(o.controls.sendClicks, 1);
});
test('CC: pills land in the right wells (compose kept open)', async () => {
  const o = await createOutlook({ sendMode: 'never' });
  await run(o, job({ cc: ['b@example.com', 'c@example.com'] }));
  assert.equal(wellOf(o.doc, 'a@example.com'), 'to');
  assert.equal(wellOf(o.doc, 'b@example.com'), 'cc');
  assert.equal(wellOf(o.doc, 'c@example.com'), 'cc');
});
test('no CC in the job: Cc control is never touched', async () => {
  const o = await createOutlook();
  const { r } = await run(o, job({ cc: [] }));
  assert.equal(r.outcome, 'sent');
  assert.equal(o.controls.ccToggleClicks, 0);
  assert.equal(o.controls.ccTouched, false);
});
test('CC unresolved => failed before Send (never sends with a partial recipient list)', async () => {
  const o = await createOutlook({ ccResolve: false });
  const { r } = await run(o, job({ cc: ['b@example.com'] }));
  assert.equal(r.outcome, 'failed');
  assert.equal(r.sendClicked, false);
  assert.equal(o.controls.sendClicks, 0);
});
test('missing Cc control => failed before Send', async () => {
  const o = await createOutlook({ noCcToggle: true });
  const { r } = await run(o, job({ cc: ['b@example.com'] }));
  assert.equal(r.code, 'CC_FIELD_MISSING');
  assert.equal(o.controls.sendClicks, 0);
});
test('invalid or overlapping CC lists never reach Outlook', async () => {
  for (const cc of [['nope'], ['a@example.com'], ['b@example.com', 'B@example.com'], 'b@example.com']) {
    const o = await createOutlook();
    const { r } = await run(o, job({ cc }));
    assert.equal(r.code, 'INVALID_JOB', JSON.stringify(cc));
    assert.equal(o.controls.newMailClicks, 0);
  }
});

test('several CC addresses: earlier chips are NOT wiped when the next one is typed', async () => {
  const o = await createOutlook({ sendMode: 'never' });
  const cc = ['b@example.com', 'c@example.com', 'd@example.com'];
  const { r } = await run(o, job({ cc }));
  assert.equal(r.code, 'NO_CONFIRMATION', 'got as far as clicking Send');
  for (const a of cc) assert.equal(wellOf(o.doc, a), 'cc', a);
  assert.equal(wellOf(o.doc, 'a@example.com'), 'to');
});
test('address that only resolves on a typed ";" is handled (second attempt)', async () => {
  const o = await createOutlook({ resolveOn: 'semicolon', sendMode: 'never' });
  const { r } = await run(o, job({ cc: ['b@example.com'] }));
  assert.equal(r.code, 'NO_CONFIRMATION');
  assert.equal(wellOf(o.doc, 'a@example.com'), 'to');
  assert.equal(wellOf(o.doc, 'b@example.com'), 'cc');
});
test('the To field holds exactly one chip for the To address and no leftover raw text', async () => {
  const o = await createOutlook({ sendMode: 'never' });
  await run(o, job());
  const to = o.doc.querySelector('[aria-label="To"][contenteditable]');
  assert.equal(to.querySelectorAll('span[contenteditable="false"]').length, 1);
  assert.equal(Array.from(to.childNodes).filter((n) => n.nodeType === 3).length, 0);
});

// ---- Sent Items confirmation ----
function runVerified(o, j = job(), extra = {}) {
  const guard = newGuard(o.window);
  const stages = [];
  return o.OEA.processJob(j, {
    doc: o.doc, loc: o.loc, guard, timeouts: { ...FAST, sentItems: 800 }, options: { verifySentItems: true },
    markStage: async (id, s) => { stages.push(s); if (extra.failStage === s) throw new Error('storage down'); },
  }).then((r) => ({ r, stages }));
}

test('Sent Items confirmation: message found => sent, with stronger evidence; view restored to the Inbox', async () => {
  const o = await createOutlook();
  const { r, stages } = await runVerified(o);
  assert.equal(r.outcome, 'sent');
  assert.match(r.evidence, /listed in Sent Items/);
  assert.equal(o.controls.sendClicks, 1);
  assert.equal(o.controls.folderClicks.sent, 1);
  assert.equal(o.controls.folderClicks.inbox, 1, 'went back to where the user was');
  assert.equal(stages.at(-1), 'verifying_sent_items');
});
test('the Sent Items check is OFF unless enabled (no folder navigation)', async () => {
  const o = await createOutlook();
  const { r } = await run(o);
  assert.equal(r.outcome, 'sent');
  assert.equal(o.controls.folderClicks.sent, 0);
});
test('Sent Items: message listed after a delay is still found', async () => {
  const o = await createOutlook({ sentItems: 'delayed' });
  const { r } = await runVerified(o);
  assert.equal(r.outcome, 'sent');
});
test('Sent Items: message never listed => uncertain (needs review), never "sent", not resent', async () => {
  const o = await createOutlook({ sentItems: 'missing' });
  const { r } = await runVerified(o);
  assert.equal(r.outcome, 'uncertain');
  assert.equal(r.code, 'SENT_ITEMS_NOT_FOUND');
  assert.equal(r.sendClicked, true);
  assert.equal(o.controls.sendClicks, 1);
});
test('Sent Items: an OLD message with the same subject does not count', async () => {
  const o = await createOutlook({ sentItems: 'old' });
  const { r } = await runVerified(o);
  assert.equal(r.code, 'SENT_ITEMS_NOT_FOUND');
});
test('Sent Items: same subject to a DIFFERENT recipient does not count', async () => {
  const o = await createOutlook({ sentItems: 'wrong-recipient' });
  const { r } = await runVerified(o);
  assert.equal(r.code, 'SENT_ITEMS_NOT_FOUND');
});
test('Sent Items: a contact display name instead of the address is accepted', async () => {
  const o = await createOutlook({ sentItems: 'name' });
  const { r } = await runVerified(o);
  assert.equal(r.outcome, 'sent');
});
test('Sent Items folder control missing => uncertain with its own code', async () => {
  const o = await createOutlook({ sentItems: 'no-folder' });
  const { r } = await runVerified(o);
  assert.equal(r.outcome, 'uncertain');
  assert.equal(r.code, 'SENT_ITEMS_UNAVAILABLE');
  assert.equal(o.controls.sendClicks, 1);
});
test('Sent Items stage is recorded as post-send (crash recovery => needs review)', async () => {
  const { classifyInterrupted } = await import('../state.js');
  assert.equal(classifyInterrupted({ jobId: 'j', stage: 'verifying_sent_items', at: 1 }, 'j').status, 'needs_review');
});

test('Sent Items miss reports counts of what it saw (no personal data) so a mismatch can be diagnosed', async () => {
  const o = await createOutlook({ sentItems: 'old' });
  const { r } = await runVerified(o);
  assert.equal(r.code, 'SENT_ITEMS_NOT_FOUND');
  assert.match(r.message, /folder opened=yes; rows listed=1; rows with the subject=1; of those with the recipient=1; time readable=1; within the time window=0/);
  assert.ok(!/@/.test(r.message), 'no addresses in the message');
});
test('Sent Items miss with nothing listed says so', async () => {
  const o = await createOutlook({ sentItems: 'missing' });
  const { r } = await runVerified(o);
  assert.match(r.message, /rows listed=0; rows with the subject=0/);
});

test('multi-line body is entered with its line breaks intact and the email is sent', async () => {
  const o = await createOutlook({ sendMode: 'never' });
  const body = 'Hi Ann,\n\nQuick note.\nSecond line.\n\nThanks,\nKeith';
  const { r } = await run(o, job({ body }));
  assert.equal(r.code, 'NO_CONFIRMATION', 'got as far as clicking Send');
  const editor = o.doc.querySelector('[role="textbox"][aria-label^="Message body"]');
  assert.equal(o.OEA.bodyMatchesLines(editor, body), true);
});
test('an editor that merges the lines into one => BODY_FORMAT_FAILED and NOTHING is sent', async () => {
  const o = await createOutlook({ bodyEditor: 'flatten' });
  const { r } = await run(o, job({ body: 'Line one\nLine two\nLine three' }));
  assert.equal(r.outcome, 'failed');
  assert.equal(r.sendClicked, false);
  assert.equal(r.code, 'BODY_FORMAT_FAILED');
  assert.match(r.message, /line breaks/);
  assert.equal(o.controls.sendClicks, 0);
});
test('a single-line body is fine even with such an editor', async () => {
  const o = await createOutlook({ bodyEditor: 'flatten' });
  const { r } = await run(o, job({ body: 'One single line' }));
  assert.equal(r.outcome, 'sent');
});

test('Sent Items miss also reports the page state (list found, options, tab visibility, focus)', async () => {
  const o = await createOutlook({ sentItems: 'missing' });
  const { r } = await runVerified(o);
  assert.match(r.message, /Page: message lists found=1; options in page=\d+; visible options=\d+; tab visibility=\w+; window focused=\w+\./);
});
