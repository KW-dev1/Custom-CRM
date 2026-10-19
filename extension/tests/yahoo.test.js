import test from 'node:test';
import assert from 'node:assert/strict';
import { createYahoo } from './fixtures/yahoo.js';

const guard = () => ({ processed: new Set(), current: null });
const stages = [];
const ctxOf = (y, extra = {}) => ({ doc: y.doc, loc: y.loc, guard: guard(), markStage: async (id, s) => { stages.push(s); }, timeouts: { ready: 500, compose: 800, recipient: 600, verify: 800 }, ...extra });
const job = (over = {}) => ({ id: 'r2-abc', to: 'a@example.com', subject: 'Hello there', body: 'Line one\nLine two', ...over });

test('Yahoo registers as the active provider and is ready when the mailbox is shown', async () => {
  const y = await createYahoo();
  assert.equal(y.OEA.activeProvider.name, 'yahoo');
  assert.equal(y.OEA.activeProvider.detectReadiness(y.doc, y.loc).state, 'ready');
});
test('readiness: wrong host is unexpected_page; a lookalike host too', async () => {
  const y = await createYahoo({ url: 'https://mail.yahoo.com.evil.example/' });
  assert.equal(y.OEA.activeProvider.detectReadiness(y.doc, y.loc).state, 'unexpected_page');
});
test('readiness: no mail interface yet means loading', async () => {
  const y = await createYahoo();
  y.doc.querySelector('nav').remove();
  assert.equal(y.OEA.activeProvider.detectReadiness(y.doc, y.loc).state, 'loading');
});
test('sends a single-recipient message once and reports sent', async () => {
  const y = await createYahoo();
  const r = await y.OEA.activeProvider.processJob(job(), ctxOf(y));
  assert.equal(r.outcome, 'sent', r.message);
  assert.equal(r.sendClicked, true);
  assert.equal(y.c.sendClicks, 1);
  assert.ok(stages.indexOf('send_clicking') < stages.indexOf('verifying'));
  assert.match(y.composeState.bodyText.replace(/ /g, ' '), /Line one[\s\S]*Line two/);
});
test('first address is To and the others are Cc; works when the rows collapse into "+N more"', async () => {
  const y = await createYahoo();
  const r = await y.OEA.activeProvider.processJob(job({ cc: ['b@example.com', 'c@example.com'] }), ctxOf(y));
  assert.equal(r.outcome, 'sent', r.message);
  assert.deepEqual(y.chips.To, ['a@example.com']);
  assert.deepEqual(y.chips.Cc, ['b@example.com', 'c@example.com']);
  assert.equal(y.c.ccToggleClicks, 1);
});
test('keeps working when the rows never collapse', async () => {
  const y = await createYahoo({ collapse: false });
  const r = await y.OEA.activeProvider.processJob(job({ cc: ['b@example.com'] }), ctxOf(y));
  assert.equal(r.outcome, 'sent', r.message);
});
test('a ";" terminator is used when Enter does not resolve', async () => {
  const y = await createYahoo({ enterResolves: false });
  const r = await y.OEA.activeProvider.processJob(job(), ctxOf(y));
  assert.equal(r.outcome, 'sent', r.message);
  assert.deepEqual(y.chips.To, ['a@example.com']);
});
test('a recipient that never resolves fails before Send', async () => {
  const y = await createYahoo({ resolve: false });
  const r = await y.OEA.activeProvider.processJob(job(), ctxOf(y));
  assert.equal(r.outcome, 'failed');
  assert.equal(r.sendClicked, false);
  assert.equal(r.code, 'RECIPIENT_FAILED');
  assert.equal(y.c.sendClicks, 0);
});
test('an unrelated suggestion chip is caught by validation and nothing is sent', async () => {
  const y = await createYahoo({ suggest: true, collapse: false });
  const r = await y.OEA.activeProvider.processJob(job(), ctxOf(y));
  assert.equal(r.outcome, 'failed');
  assert.equal(r.code, 'VALIDATION_FAILED');
  assert.equal(y.c.sendClicks, 0);
});
test('when the editor refuses paragraph commands the lines are still built and verified line by line', async () => {
  const y = await createYahoo({ bodyEditor: 'flatten' });
  const r = await y.OEA.activeProvider.processJob(job(), ctxOf(y));
  assert.equal(r.outcome, 'sent', r.message);
  assert.match(y.composeState.bodyText.replace(/\u00a0/g, ' '), /Line one[\s\S]*Line two/);
});
test('a compose window already open is not touched', async () => {
  const y = await createYahoo({ pre: 'compose' });
  const r = await y.OEA.activeProvider.processJob(job(), ctxOf(y));
  assert.equal(r.code, 'COMPOSE_ALREADY_OPEN');
  assert.equal(y.c.sendClicks, 0);
});
test('sent toast with the compose still open counts as sent', async () => {
  const y = await createYahoo({ sendMode: 'toast' });
  const r = await y.OEA.activeProvider.processJob(job(), ctxOf(y));
  assert.equal(r.outcome, 'sent', r.message);
  assert.match(r.evidence, /message has been sent/i);
});
test('no confirmation after Send is uncertain and never resent', async () => {
  const y = await createYahoo({ sendMode: 'never' });
  const c = ctxOf(y, { timeouts: { ready: 500, compose: 800, recipient: 600, verify: 300 } });
  const r = await y.OEA.activeProvider.processJob(job(), c);
  assert.equal(r.outcome, 'uncertain');
  assert.equal(r.code, 'NO_CONFIRMATION');
  assert.equal(y.c.sendClicks, 1);
  const again = await y.OEA.activeProvider.processJob(job(), c);
  assert.equal(again.code, 'DUPLICATE_JOB');
  assert.equal(y.c.sendClicks, 1);
});
test('an error shown after Send is uncertain', async () => {
  const y = await createYahoo({ sendMode: 'error' });
  const r = await y.OEA.activeProvider.processJob(job(), ctxOf(y));
  assert.equal(r.outcome, 'uncertain');
  assert.equal(r.code, 'SEND_ERROR_SHOWN');
});
test('an invalid job is rejected before touching the page', async () => {
  const y = await createYahoo();
  const r = await y.OEA.activeProvider.processJob(job({ to: 'not-an-email' }), ctxOf(y));
  assert.equal(r.code, 'INVALID_JOB');
  assert.equal(y.c.newMessageClicks, 0);
});
test('the Outlook flow still registers itself as the active provider', async () => {
  const { createOutlook } = await import('./fixtures/outlook.js');
  const o = await createOutlook();
  assert.equal(o.window.OEA.activeProvider.name, 'outlook');
  assert.equal(o.window.OEA.activeProvider.detectReadiness(o.doc, o.loc).state, 'ready');
});
test('a subject Yahoo\'s own state never registered is caught before Send (no "(No subject)" email)', async () => {
  const y = await createYahoo({ ignoreSubject: true });
  const r = await y.OEA.activeProvider.processJob(job(), ctxOf(y));
  assert.equal(r.outcome, 'failed');
  assert.equal(r.code, 'VALIDATION_FAILED');
  assert.match(r.message, /subject/i);
  assert.equal(y.c.sendClicks, 0);
});
