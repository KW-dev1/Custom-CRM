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

// ---- recipient chips: regression tests for the real Outlook structure ----
const bareCompose = async () => {
  const o = await createOutlook({ pre: 'compose' });
  return { o, c: o.OEA.findCompose(o.doc) };
};
const addChip = (o, field, email) => {
  const chip = o.doc.createElement('span');
  chip.setAttribute('contenteditable', 'false'); chip.setAttribute('aria-label', email); chip.textContent = email;
  field.appendChild(chip);
  return chip;
};
test('a resolved chip INSIDE the To field is recognised (real Outlook layout)', async () => {
  const { o, c } = await bareCompose();
  addChip(o, c.to, 'a@example.com');
  assert.ok(o.OEA.findRecipientPill(c, 'a@example.com'));
  assert.ok(o.OEA.findRecipientPill(c, 'A@Example.com'), 'case-insensitive');
});
test('raw typed text in the field is NOT a resolved recipient', async () => {
  const { o, c } = await bareCompose();
  c.to.appendChild(o.doc.createTextNode('a@example.com'));
  assert.equal(o.OEA.findRecipientPill(c, 'a@example.com'), null);
});
test('a chip for a different address does not match', async () => {
  const { o, c } = await bareCompose();
  addChip(o, c.to, 'other@example.com');
  assert.equal(o.OEA.findRecipientPill(c, 'a@example.com'), null);
});
test('fallback: a chip outside the field (other layouts) is still recognised', async () => {
  const { o, c } = await bareCompose();
  const chip = o.doc.createElement('span');
  chip.setAttribute('role', 'button'); chip.setAttribute('aria-label', 'a@example.com'); chip.textContent = 'a@example.com';
  c.root.appendChild(chip);
  assert.ok(o.OEA.findRecipientPill(c, 'a@example.com'));
});
test('appendEditableText keeps existing chips and adds text at the end', async () => {
  const { o, c } = await bareCompose();
  addChip(o, c.to, 'first@example.com');
  o.OEA.appendEditableText(c.to, 'second');
  assert.equal(c.to.firstChild.getAttribute('aria-label'), 'first@example.com');
  assert.equal(c.to.lastChild.data, 'second');
});
test('real Outlook structure: To has no role=textbox, Cc/Bcc are plain buttons, both are found', async () => {
  const { o, c } = await bareCompose();
  assert.equal(c.to.getAttribute('role'), null);
  assert.ok(o.OEA.findCcToggle(c.root));
  assert.equal(o.OEA.findCc(c.root), null, 'Cc field is hidden until its button is clicked');
});

// ---- Sent Items helpers (structure captured from real Outlook Web) ----
const NOW = new Date('2026-10-15T22:53:10').getTime(); // local time
const SENT_AT = new Date('2026-10-15T22:52:41').getTime();
const row = (label, { title = 'Thu 10/15/2026 10:52 PM', text = '10:52 PM' } = {}) => ({ label, timeTitle: title, timeText: text });
const q = { subject: 'potential collaboration', to: 'austinames410@gmail.com', sentAtMs: SENT_AT, nowMs: NOW };

test('Sent Items folder control is found by data-folder-name (both tree entries)', async () => {
  const o = await createOutlook();
  o.doc.body.insertAdjacentHTML('beforeend', `<div role="tree">
    <div role="treeitem" data-folder-name="sent items" aria-selected="false" title="Sent Items - 1 item (0 unread)"><span>Sent Items</span></div>
    <div role="treeitem" data-folder-name="inbox" aria-selected="true" title="Inbox"><span>Inbox</span></div></div>`);
  assert.ok(o.OEA.findFolder(o.doc, 'sent items'));
  assert.equal(o.OEA.findFolder(o.doc, 'archive'), null);
  assert.equal(o.OEA.findSelectedFolder(o.doc).getAttribute('data-folder-name'), 'inbox');
  assert.equal(o.OEA.isFolderSelected(o.doc, 'sent items'), false);
  o.OEA.findFolder(o.doc, 'sent items').setAttribute('aria-selected', 'true');
  assert.equal(o.OEA.isFolderSelected(o.doc, 'sent items'), true);
});
test('readMessageRows reads option rows with label and time span', async () => {
  const o = await createOutlook();
  o.doc.getElementById('msglist').remove();
  o.doc.body.insertAdjacentHTML('beforeend', `<div role="region"><div role="listbox" aria-label="Message list">
    <div role="option" aria-label="austinames410@gmail.com potential collaboration 10:52 PM hi there">
      <span>austinames410@gmail.com</span><span>potential collaboration</span><span title="Thu 10/15/2026 10:52 PM">10:52 PM</span><span>hi there</span></div></div></div>`);
  const rows = o.OEA.readMessageRows(o.doc);
  assert.equal(rows.length, 1);
  assert.match(rows[0].label, /potential collaboration/);
  assert.equal(rows[0].timeTitle, 'Thu 10/15/2026 10:52 PM');
  assert.equal(rows[0].timeText, '10:52 PM');
});
test('readMessageRows: no list => no rows', async () => {
  const o = await createOutlook();
  o.doc.getElementById('msglist').remove();
  assert.equal(o.OEA.readMessageRows(o.doc).length, 0);
});
test('parseRowTime: full title, 12h/24h clock fallbacks, day-month locale mix-ups, junk', async () => {
  const { OEA } = await createOutlook();
  const t = (info) => { const v = OEA.parseRowTime(info, NOW); return v === null ? null : new Date(v).toTimeString().slice(0, 5); };
  assert.equal(t(row('x')), '22:52');
  assert.equal(t(row('x', { title: '', text: '10:52 PM' })), '22:52');
  assert.equal(t(row('x', { title: '15/10/2026 22:52', text: '22:52' })), '22:52');
  assert.equal(t(row('x', { title: '', text: '12:05 AM' })), '00:05');
  assert.equal(t(row('x', { title: '', text: '12:05 PM' })), '12:05');
  assert.equal(t(row('x', { title: 'Mon 11/09/2026 10:52 PM', text: '10:52 PM' })), '22:52', 'a far-off date is ignored in favour of today\'s clock');
  assert.equal(t(row('x', { title: '', text: 'Thu 10/15' })), null);
});
test('matchSentRow: accepts the fresh message (email shown)', async () => {
  const { OEA } = await createOutlook();
  assert.ok(OEA.matchSentRow([row('austinames410@gmail.com potential collaboration 10:52 PM hi')], q));
});
test('matchSentRow: contact display name instead of an email is accepted (recipient not checkable)', async () => {
  const { OEA } = await createOutlook();
  assert.ok(OEA.matchSentRow([row('Austin Ames potential collaboration 10:52 PM hi')], q));
});
test('matchSentRow: rejects other recipient, other subject, old message, unreadable time', async () => {
  const { OEA } = await createOutlook();
  const no = (r) => assert.equal(OEA.matchSentRow([r], q), null);
  no(row('someone.else@gmail.com potential collaboration 10:52 PM'));
  no(row('austinames410@gmail.com different subject 10:52 PM'));
  no(row('austinames410@gmail.com potential collaboration 8:30 PM', { title: 'Thu 10/15/2026 8:30 PM', text: '8:30 PM' }));
  no(row('austinames410@gmail.com potential collaboration Tue', { title: 'Tue 10/13/2026 10:52 PM', text: 'Tue' }));
  no(row('austinames410@gmail.com potential collaboration', { title: '', text: '' }));
});
test('matchSentRow: minute-granularity margin, case-insensitive subject, picks the right row among several', async () => {
  const { OEA } = await createOutlook();
  const rows = [
    row('austinames410@gmail.com POTENTIAL   Collaboration 10:51 PM', { title: 'Thu 10/15/2026 10:51 PM', text: '10:51 PM' }),
    row('other@x.com potential collaboration 10:52 PM'),
  ];
  assert.equal(OEA.matchSentRow(rows, q), rows[0]);
  const early = row('austinames410@gmail.com potential collaboration 10:50 PM', { title: 'Thu 10/15/2026 10:50 PM', text: '10:50 PM' });
  assert.equal(OEA.matchSentRow([early], q), null, 'two minutes before the click is too old');
});

test('diagnoseSentRows counts rows, subject, recipient, readable time and window separately', async () => {
  const { OEA } = await createOutlook();
  const rows = [
    row('austinames410@gmail.com potential collaboration 10:52 PM'),
    row('other@x.com potential collaboration 10:52 PM'),
    row('austinames410@gmail.com potential collaboration 8:30 PM', { title: 'Thu 10/15/2026 8:30 PM', text: '8:30 PM' }),
    row('austinames410@gmail.com unrelated subject 10:52 PM'),
    row('austinames410@gmail.com potential collaboration', { title: '', text: '' }),
  ];
  assert.deepEqual({ ...OEA.diagnoseSentRows(rows, q) }, { rows: 5, subject: 4, recipient: 3, timeReadable: 2, inWindow: 1 });
});
