import test from 'node:test';
import assert from 'node:assert/strict';
import { parseEmailList, rowToJob, rowHash, parseJobId, makeJobId, hasContent } from '../src/rows.js';

test('single address', () => assert.deepEqual(parseEmailList('a@example.com'), { emails: ['a@example.com'], invalid: [] }));
test('JSON array: first is TO, others follow in order', () => {
  assert.deepEqual(parseEmailList('["a@x.com","b@y.org","c@z.io"]').emails, ['a@x.com', 'b@y.org', 'c@z.io']);
});
test('comma / semicolon / newline / space separated and bracketed lists', () => {
  assert.deepEqual(parseEmailList('a@x.com, b@y.org;c@z.io\nd@w.net e@v.com').emails.length, 5);
  assert.deepEqual(parseEmailList("['a@x.com', 'b@y.org']").emails, ['a@x.com', 'b@y.org']);
});
test('display names and angle brackets', () => {
  assert.deepEqual(parseEmailList('Bob Smith <bob@x.com>, "Al, Jr" <al@y.org>').emails, ['bob@x.com', 'al@y.org']);
});
test('duplicates removed case-insensitively, first spelling kept', () => {
  assert.deepEqual(parseEmailList('A@X.com a@x.com b@y.org').emails, ['A@X.com', 'b@y.org']);
});
test('empty / blank / non-string cells', () => {
  for (const v of ['', '   ', null, undefined]) assert.deepEqual(parseEmailList(v), { emails: [], invalid: [] });
});
test('malformed addresses are reported, not silently dropped', () => {
  const r = parseEmailList('ok@x.com, bad@, @nope.com, two@@x.com, nodot@x');
  assert.deepEqual(r.emails, ['ok@x.com']);
  assert.equal(r.invalid.length, 4);
});
test('broken JSON falls back to loose parsing', () => {
  assert.deepEqual(parseEmailList('["a@x.com", "b@y.org"').emails, ['a@x.com', 'b@y.org']);
});

const row = (o = {}) => ({ rowNumber: 5, emails: '["a@x.com","b@y.org","c@z.io"]', subject: 'Hi', body: 'Line1\r\nLine2\r\n', ...o });

test('rowToJob: first email TO, rest CC, body normalised', () => {
  const r = rowToJob(row());
  assert.equal(r.ok, true);
  assert.equal(r.job.to, 'a@x.com');
  assert.deepEqual(r.job.cc, ['b@y.org', 'c@z.io']);
  assert.equal(r.job.body, 'Line1\nLine2');
  assert.deepEqual(r.job.metadata, { row: 5 });
});
test('rowToJob: one address means no CC', () => assert.deepEqual(rowToJob(row({ emails: 'a@x.com' })).job.cc, []));
test('rowToJob: subject newlines collapse to one line', () => assert.equal(rowToJob(row({ subject: 'A\nB  ' })).job.subject, 'A B'));
test('rowToJob: invalid rows give a useful error', () => {
  for (const [o, re] of [
    [{ emails: '' }, /No email/], [{ emails: 'a@x.com, oops@' }, /Invalid email/],
    [{ subject: '  ' }, /Subject/], [{ body: '\n ' }, /content/],
    [{ emails: 'a@x.com,b@x.com,c@x.com' }, /CC/],
  ]) {
    const r = rowToJob(row({ emails: 'a@x.com', ...o }), { maxCc: 1 });
    assert.equal(r.ok, false, JSON.stringify(o));
    assert.match(r.error, re);
  }
});
test('job id encodes row + content hash; editing the row changes the id', () => {
  const a = rowToJob(row()).job.id;
  const b = rowToJob(row({ subject: 'Changed' })).job.id;
  assert.notEqual(a, b);
  const p = parseJobId(a);
  assert.equal(p.rowNumber, 5);
  assert.equal(p.hash, rowHash(row()));
  assert.equal(a, makeJobId(5, p.hash));
});
test('parseJobId rejects junk', () => {
  for (const x of ['', 'r5', 'r5-XYZ', 'x5-aaaaaaaaaaaa', '../etc', null, 'r5-aaaaaaaaaaaaa']) assert.equal(parseJobId(x), null);
});
test('hasContent', () => {
  assert.equal(hasContent({}), false);
  assert.equal(hasContent({ emails: ' ', subject: '', body: '' }), false);
  assert.equal(hasContent({ body: 'x' }), true);
});
