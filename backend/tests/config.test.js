import test from 'node:test';
import assert from 'node:assert/strict';
import { loadConfig, ConfigError, parseSheetUrl, colIndex } from '../src/config.js';

const base = () => ({
  API_KEYS: 'abcdefghijklmnop1234',
  SHEET_URL: 'https://docs.google.com/spreadsheets/d/10sk2-aVC824trbNjPmfaARiTwSh18WGhqD5ryanEhbI/edit?gid=1793568655#gid=1793568655',
  GOOGLE_CLIENT_ID: 'id', GOOGLE_CLIENT_SECRET: 'secret', GOOGLE_REFRESH_TOKEN: 'refresh',
});

test('parses sheet id and gid from the browser URL', () => {
  assert.deepEqual(parseSheetUrl(base().SHEET_URL), { id: '10sk2-aVC824trbNjPmfaARiTwSh18WGhqD5ryanEhbI', gid: '1793568655' });
  assert.deepEqual(parseSheetUrl('nope'), { id: null, gid: null });
});
test('valid env loads with defaults G/H/I/J', () => {
  const c = loadConfig(base());
  assert.deepEqual(c.columns, { emails: 'G', subject: 'H', body: 'I', status: 'J', note: null });
  assert.equal(c.sheet.gid, '1793568655');
  assert.equal(c.minClaimIntervalSec, 300);
  assert.equal(c.statusLabels.pending, 'pending');
});
test('missing and weak settings are all reported together', () => {
  try { loadConfig({ API_KEYS: 'short' }); assert.fail('should throw'); } catch (e) {
    assert.ok(e instanceof ConfigError);
    for (const s of ['16 characters', 'SHEET_URL', 'GOOGLE_CLIENT_ID', 'GOOGLE_REFRESH_TOKEN']) assert.ok(e.message.includes(s), s);
  }
});
test('column letters validated and must be distinct', () => {
  assert.throws(() => loadConfig({ ...base(), COL_EMAILS: '7' }), /COL_EMAILS/);
  assert.throws(() => loadConfig({ ...base(), COL_BODY: 'G' }), /different/);
  assert.equal(loadConfig({ ...base(), COL_NOTE: 'k' }).columns.note, 'K');
});
test('numeric bounds and status label rules', () => {
  assert.throws(() => loadConfig({ ...base(), PORT: '99999' }), /PORT/);
  assert.throws(() => loadConfig({ ...base(), STATUS_SENT: 'pending' }), /distinct/);
});
test('multiple API keys for rotation', () => {
  assert.equal(loadConfig({ ...base(), API_KEYS: 'abcdefghijklmnop1234, zyxwvutsrqponmlk9876' }).apiKeys.length, 2);
});
test('column index helper', () => {
  assert.deepEqual(['A', 'G', 'J', 'Z', 'AA'].map(colIndex), [0, 6, 9, 25, 26]);
});
