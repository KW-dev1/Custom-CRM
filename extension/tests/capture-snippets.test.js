import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createOutlook } from './fixtures/outlook.js';

// The console snippets in docs/CAPTURE_DOM.md are what a user runs on a real mail page. Run them here, as written, on the
// simulated pages: they must work, find the right elements and never leak an email address.
const doc = readFileSync(new URL('../../docs/CAPTURE_DOM.md', import.meta.url), 'utf8');
const blocks = [...doc.matchAll(/```js\n([\s\S]*?)```/g)].map((m) => m[1]);
const [compose, sent] = blocks;

function run(o, code) {
  let copied = null;
  const logs = [];
  // jsdom has no layout: give every element a size so the "visible" filters pass
  o.window.Element.prototype.getBoundingClientRect = () => ({ width: 10, height: 10, top: 0, left: 0, right: 10, bottom: 10 });
  o.window.copy = (t) => { copied = t; };
  o.window.console = { log: (m) => logs.push(m) };
  o.window.eval(code);
  return { copied, logs };
}

test('the document contains both snippets', () => assert.equal(blocks.length, 2));

test('compose snippet: runs, finds subject/To/body/Send, redacts addresses', async () => {
  const o = await createOutlook({ pre: 'compose' });
  const to = o.doc.querySelector('[aria-label="To"][contenteditable]');
  to.insertAdjacentHTML('beforeend', '<span contenteditable="false" aria-label="someone.private@example.com">someone.private@example.com</span>');
  const { copied, logs } = run(o, compose);
  assert.ok(copied, `expected output, got: ${logs.join(' | ')}`);
  const rows = JSON.parse(copied);
  assert.ok(rows.length > 4);
  assert.ok(rows.some((r) => r['aria-label'] === 'Subject'));
  assert.ok(rows.some((r) => r['aria-label'] === 'To' && r.contenteditable === 'true'));
  assert.ok(rows.some((r) => r.tag === 'button' && r['aria-label'] === 'Send'));
  assert.ok(rows.some((r) => r['aria-label'] === '<email>' && r.contenteditable === 'false'), 'the recipient chip is shown with its address redacted');
  assert.ok(!copied.includes('@'), 'no email address may appear in the output');
});
test('compose snippet: with no compose open it says so instead of failing', async () => {
  const o = await createOutlook();
  const { copied, logs } = run(o, compose);
  assert.equal(copied, null);
  assert.match(logs.join(' '), /NO SUBJECT FIELD FOUND/);
});
test('sent-folder snippet: finds the folder control and the message row, redacts addresses', async () => {
  const o = await createOutlook();
  const { copied, logs } = run(o, sent.replace('PUT YOUR TEST EMAIL SUBJECT HERE', 'Weekly digest'));
  assert.ok(copied, logs.join(' | '));
  const out = JSON.parse(copied);
  assert.ok(out.folder.length >= 1);
  assert.ok(out.folder.some((f) => f.parents.some((p) => p['data-folder-name'] === 'sent items')));
  assert.equal(out.subjectFound, true);
  assert.equal(out.rowGuess.role, 'option');
  assert.match(out.rowText, /<email> \| Weekly digest \| /, 'row pieces are separated, so the subject survives address redaction');
  assert.ok(out.rowDescendants.length >= 3);
  assert.ok(!copied.includes('@'), 'no email address may appear in the output');
});
test('sent-folder snippet: unknown subject is reported, not a crash', async () => {
  const o = await createOutlook();
  const { copied } = run(o, sent.replace('PUT YOUR TEST EMAIL SUBJECT HERE', 'a subject that is not on the page'));
  assert.equal(JSON.parse(copied).subjectFound, false);
});
test('the placeholder text is still in the document (users must replace it)', () => {
  assert.match(sent, /PUT YOUR TEST EMAIL SUBJECT HERE/);
});
