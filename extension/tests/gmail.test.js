import test from 'node:test';
import assert from 'node:assert/strict';
import { createGmail } from './fixtures/gmail.js';

const guard = () => ({ processed: new Set(), current: null });
const ctxOf = (g, extra = {}) => ({ doc: g.doc, loc: g.loc, guard: guard(), markStage: async () => {}, timeouts: { ready: 500, compose: 800, recipient: 600, verify: 800 }, ...extra });
const job = (over = {}) => ({ id: 'r2-abc', to: 'a@example.com', subject: 'Hello there', body: 'Line one\nLine two', ...over });

test('Gmail registers as the active provider and is ready when the mailbox is shown', async () => {
  const g = await createGmail();
  assert.equal(g.OEA.activeProvider.name, 'gmail');
  assert.equal(g.OEA.activeProvider.detectReadiness(g.doc, g.loc).state, 'ready');
});
test('readiness: other hosts and lookalikes are unexpected_page; no interface yet is loading', async () => {
  const a = await createGmail({ url: 'https://mail.google.com.evil.example/mail/' });
  assert.equal(a.OEA.activeProvider.detectReadiness(a.doc, a.loc).state, 'unexpected_page');
  const b = await createGmail({ url: 'https://accounts.google.com/signin' });
  assert.equal(b.OEA.activeProvider.detectReadiness(b.doc, b.loc).state, 'unexpected_page');
  const c = await createGmail();
  c.doc.querySelector('[gh=cm]').remove();
  assert.equal(c.OEA.activeProvider.detectReadiness(c.doc, c.loc).state, 'loading');
});
test('sends a single-recipient message once, using real mouse events, and reports sent', async () => {
  const g = await createGmail();
  const r = await g.OEA.activeProvider.processJob(job(), ctxOf(g));
  assert.equal(r.outcome, 'sent', r.message);
  assert.equal(g.c.sendClicks, 1);
  assert.deepEqual(g.chips.To, ['a@example.com']);
  assert.match(g.st.bodyText.replace(/ /g, ' '), /Line one[\s\S]*Line two/);
});
test('first address is To, the others Cc (opened with the Cc link)', async () => {
  const g = await createGmail();
  const r = await g.OEA.activeProvider.processJob(job({ cc: ['b@example.com', 'c@example.com'] }), ctxOf(g));
  assert.equal(r.outcome, 'sent', r.message);
  assert.deepEqual(g.chips.To, ['a@example.com']);
  assert.deepEqual(g.chips.Cc, ['b@example.com', 'c@example.com']);
  assert.equal(g.c.ccLinkClicks, 1);
});
test('a trailing "," is used when Enter does not resolve', async () => {
  const g = await createGmail({ enterResolves: false });
  const r = await g.OEA.activeProvider.processJob(job(), ctxOf(g));
  assert.equal(r.outcome, 'sent', r.message);
});
test('a recipient that never resolves fails before Send', async () => {
  const g = await createGmail({ resolve: false });
  const r = await g.OEA.activeProvider.processJob(job(), ctxOf(g));
  assert.equal(r.outcome, 'failed');
  assert.equal(r.code, 'RECIPIENT_FAILED');
  assert.equal(g.c.sendClicks, 0);
});
test('an extra unrelated chip is caught by validation and nothing is sent', async () => {
  const g = await createGmail({ suggest: true });
  const r = await g.OEA.activeProvider.processJob(job(), ctxOf(g));
  assert.equal(r.outcome, 'failed');
  assert.equal(r.code, 'VALIDATION_FAILED');
  assert.equal(g.c.sendClicks, 0);
});
test('missing Cc link fails before Send', async () => {
  const g = await createGmail({ noCcLink: true });
  const r = await g.OEA.activeProvider.processJob(job({ cc: ['b@example.com'] }), ctxOf(g));
  assert.equal(r.code, 'CC_FIELD_MISSING');
  assert.equal(g.c.sendClicks, 0);
});
test('a compose window already open is not touched', async () => {
  const g = await createGmail({ pre: 'compose' });
  const r = await g.OEA.activeProvider.processJob(job(), ctxOf(g));
  assert.equal(r.code, 'COMPOSE_ALREADY_OPEN');
  assert.equal(g.c.sendClicks, 0);
});
test('"Message sent" notice with the compose still open counts as sent', async () => {
  const g = await createGmail({ sendMode: 'toast' });
  const r = await g.OEA.activeProvider.processJob(job(), ctxOf(g));
  assert.equal(r.outcome, 'sent', r.message);
  assert.match(r.evidence, /message sent/i);
});
test('no confirmation after Send is uncertain and never resent', async () => {
  const g = await createGmail({ sendMode: 'never' });
  const c = ctxOf(g, { timeouts: { ready: 500, compose: 800, recipient: 600, verify: 300 } });
  const r = await g.OEA.activeProvider.processJob(job(), c);
  assert.equal(r.outcome, 'uncertain');
  assert.equal(r.code, 'NO_CONFIRMATION');
  const again = await g.OEA.activeProvider.processJob(job(), c);
  assert.equal(again.code, 'DUPLICATE_JOB');
  assert.equal(g.c.sendClicks, 1);
});
test('an error shown after Send is uncertain', async () => {
  const g = await createGmail({ sendMode: 'error' });
  const r = await g.OEA.activeProvider.processJob(job(), ctxOf(g));
  assert.equal(r.outcome, 'uncertain');
  assert.equal(r.code, 'SEND_ERROR_SHOWN');
});
test('an invalid job is rejected before touching the page', async () => {
  const g = await createGmail();
  const r = await g.OEA.activeProvider.processJob(job({ to: 'nope' }), ctxOf(g));
  assert.equal(r.code, 'INVALID_JOB');
  assert.equal(g.c.composeClicks, 0);
});
test('Gmail\'s own copy of the subject never updating blocks the send (no "(No subject)" email)', async () => {
  const g = await createGmail({ subjectMirror: 'never' });
  const r = await g.OEA.activeProvider.processJob(job(), ctxOf(g, { timeouts: { ready: 500, compose: 800, recipient: 600, verify: 800 } }));
  assert.equal(r.outcome, 'failed');
  assert.equal(r.code, 'SUBJECT_FAILED');
  assert.equal(g.c.sendClicks, 0);
});
test('a subject that only registers from key events is entered that way', async () => {
  const g = await createGmail({ subjectMirror: 'keys' });
  const r = await g.OEA.activeProvider.processJob(job(), ctxOf(g, { timeouts: { ready: 500, compose: 800, recipient: 600, verify: 800 } }));
  assert.equal(r.outcome, 'sent', r.message);
  assert.equal(g.st.mirror, 'Hello there');
});
