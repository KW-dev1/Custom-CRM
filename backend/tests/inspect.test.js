import test from 'node:test';
import assert from 'node:assert/strict';
import { summarizeRows } from '../src/inspect.js';
import { testConfig, lead } from './_fakes.js';

const rows = (list) => list.map((r, i) => ({ rowNumber: i + 2, emails: '', subject: '', body: '', status: '', ...r }));

test('summarizeRows counts every kind of row and never exposes content', () => {
  const s = summarizeRows(rows([
    lead(1), lead(2), { emails: 'a@x.com', subject: '', body: '' }, { emails: 'bad@', subject: 'S', body: 'B' }, {},
    lead(6, { status: 'pending' }), lead(7, { status: 'sent' }), lead(8, { status: 'FAILED' }), lead(9, { status: 'needs_review' }), lead(10, { status: 'on hold' }),
  ]), testConfig());
  assert.equal(s.total, 10);
  assert.equal(s.ready, 2);
  assert.equal(s.incomplete, 1);
  assert.deepEqual(s.invalid, [{ row: 5, reason: 'Invalid email address in column' }]);
  assert.equal(s.blank, 1);
  assert.deepEqual([s.pending, s.sent, s.failed, s.needs_review, s.other], [1, 1, 1, 1, 1]);
  assert.ok(!JSON.stringify(s).includes('@'), 'no email addresses in the summary');
});
test('custom status labels are recognised', () => {
  const s = summarizeRows(rows([lead(1, { status: 'done' })]), testConfig({ STATUS_SENT: 'done' }));
  assert.equal(s.sent, 1);
});
