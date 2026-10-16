import test from 'node:test';
import assert from 'node:assert/strict';
import { createOutlook } from './fixtures/outlook.js';

test('readiness: ready on the mail page with the New mail control', async () => {
  const o = await createOutlook();
  assert.equal(o.OEA.detectReadiness(o.doc, o.loc).state, 'ready');
});
test('readiness: signed out page is detected', async () => {
  const o = await createOutlook({ pre: 'signedOut' });
  assert.equal(o.OEA.detectReadiness(o.doc, o.loc).state, 'signed_out');
});
test('readiness: unknown shell is "loading", wrong URL is "unexpected_page"', async () => {
  const o = await createOutlook({ pre: 'noShell' });
  assert.equal(o.OEA.detectReadiness(o.doc, o.loc).state, 'loading');
  const w = await createOutlook({ url: 'https://outlook.live.com/calendar/' });
  assert.equal(w.OEA.detectReadiness(w.doc, w.loc).state, 'unexpected_page');
  const h = await createOutlook({ url: 'https://outlook.live.com.evil.example/mail/' });
  assert.equal(h.OEA.detectReadiness(h.doc, h.loc).state, 'unexpected_page');
});
test('element detection: New mail button; no compose before click', async () => {
  const o = await createOutlook();
  assert.ok(o.OEA.findNewMail(o.doc));
  assert.equal(o.OEA.findCompose(o.doc), null);
});
test('element detection: compose parts found by accessible names, To is not Cc', async () => {
  const o = await createOutlook({ pre: 'compose' });
  const c = o.OEA.findCompose(o.doc);
  assert.ok(c && c.to && c.subject && c.body && c.send);
  assert.equal(c.to.getAttribute('aria-label'), 'To');
});
test('element detection: missing Send button => no compose match', async () => {
  const o = await createOutlook({ pre: 'compose', noSend: true });
  assert.equal(o.OEA.findCompose(o.doc), null);
});
test('hidden elements are ignored', async () => {
  const o = await createOutlook();
  o.doc.querySelector('[aria-label="New mail"]').setAttribute('aria-hidden', 'true');
  assert.equal(o.OEA.findNewMail(o.doc), null);
});
test('waitFor resolves when the DOM changes', async () => {
  const o = await createOutlook();
  setTimeout(() => { const d = o.doc.createElement('i'); d.id = 'x'; o.doc.body.appendChild(d); }, 30);
  const el = await o.OEA.waitFor(() => o.doc.getElementById('x'), { timeoutMs: 1000, doc: o.doc });
  assert.equal(el.id, 'x');
});
test('waitFor times out with a bounded, descriptive error', async () => {
  const o = await createOutlook();
  const t0 = Date.now();
  await assert.rejects(o.OEA.waitFor(() => null, { timeoutMs: 100, description: 'a thing', doc: o.doc }), (e) => e.code === 'TIMEOUT' && /a thing/.test(e.message));
  assert.ok(Date.now() - t0 < 1000);
});
test('body text: multiline round-trips and HTML is not interpreted', async () => {
  const o = await createOutlook({ pre: 'compose' });
  const { body } = o.OEA.findCompose(o.doc);
  const text = 'Hi,\n\nSecond <b>line</b> & <img src=x onerror=alert(1)>\nBye';
  o.OEA.setEditableText(body, text);
  assert.equal(body.querySelector('b, img'), null);
  assert.equal(body.children.length, 4);
  assert.equal(o.OEA.squash(o.OEA.extractText(body)), o.OEA.squash(text));
});
test('extractText treats <br> and block elements as line breaks', async () => {
  const o = await createOutlook();
  const d = o.doc.createElement('div');
  d.innerHTML = '<div>a</div><div><br></div><div>b</div>c<br>d';
  assert.equal(o.OEA.extractText(d).replace(/\n+/g, '|'), 'a|b|c|d|');
});
