import test from 'node:test';
import assert from 'node:assert/strict';
import { validateSettings, normalizeSettings, backendOriginPattern } from '../config.js';
import { goodSettings } from './_helpers.js';

test('valid settings pass and base URL trailing slash is stripped', () => {
  assert.equal(validateSettings(goodSettings()).ok, true);
  assert.equal(normalizeSettings(goodSettings()).apiBaseUrl, 'https://api.example.com');
});
test('missing URL and key are reported', () => {
  const r = validateSettings({});
  assert.equal(r.ok, false);
  assert.ok(r.errors.apiBaseUrl && r.errors.apiKey);
});
test('http is rejected except for localhost', () => {
  assert.ok(validateSettings({ ...goodSettings(), apiBaseUrl: 'http://api.example.com' }).errors.apiBaseUrl);
  assert.equal(validateSettings({ ...goodSettings(), apiBaseUrl: 'http://localhost:3000' }).ok, true);
});
test('embedded credentials and garbage URLs are rejected', () => {
  assert.ok(validateSettings({ ...goodSettings(), apiBaseUrl: 'https://u:p@api.example.com' }).errors.apiBaseUrl);
  assert.ok(validateSettings({ ...goodSettings(), apiBaseUrl: 'not a url' }).errors.apiBaseUrl);
});
test('status path needs {id}; delay bounds enforced', () => {
  assert.ok(validateSettings({ ...goodSettings(), statusPath: '/status' }).errors.statusPath);
  assert.ok(validateSettings({ ...goodSettings(), fetchMinMinutes: 0 }).errors.fetchMinMinutes);
  assert.ok(validateSettings({ ...goodSettings(), fetchMinMinutes: 20, fetchMaxMinutes: 10 }).errors.fetchMaxMinutes);
  assert.equal(validateSettings({ ...goodSettings(), fetchMinMinutes: 10, fetchMaxMinutes: 20 }).ok, true);
});
test('empty status vocabulary value is rejected', () => {
  assert.ok(validateSettings({ ...goodSettings(), statusValues: { sent: '' } }).errors.statusValues);
});
test('origin pattern', () => {
  assert.equal(backendOriginPattern('https://api.example.com/v1'), 'https://api.example.com/*');
  assert.equal(backendOriginPattern('nope'), null);
});

test('verifySentItems defaults to on and can be turned off', () => {
  assert.equal(normalizeSettings(goodSettings()).verifySentItems, true);
  assert.equal(normalizeSettings({ ...goodSettings(), verifySentItems: false }).verifySentItems, false);
  assert.equal(normalizeSettings({ ...goodSettings(), verifySentItems: undefined }).verifySentItems, true);
});

test('profile name: optional, trimmed, and restricted (no email addresses)', () => {
  assert.equal(validateSettings(goodSettings()).ok, true);
  assert.equal(validateSettings({ ...goodSettings(), profileName: 'outlook-A' }).ok, true);
  assert.equal(normalizeSettings({ ...goodSettings(), profileName: '  outlook-A  ' }).profileName, 'outlook-A');
  for (const bad of ['me@outlook.com', 'a'.repeat(41), 'bad/name', 'x;y']) {
    assert.ok(validateSettings({ ...goodSettings(), profileName: bad }).errors.profileName, bad);
  }
});

test('pauseAfterProblems: default 3, whole number 1-20', () => {
  assert.equal(normalizeSettings(goodSettings()).pauseAfterProblems, 3);
  assert.equal(validateSettings({ ...goodSettings(), pauseAfterProblems: 5 }).ok, true);
  for (const bad of [0, 21, 2.5, 'abc']) assert.ok(validateSettings({ ...goodSettings(), pauseAfterProblems: bad }).errors.pauseAfterProblems, String(bad));
});

test('provider: defaults to outlook, accepts outlook/yahoo (any case), rejects others', async () => {
  const { PROVIDERS } = await import('../config.js');
  assert.equal(normalizeSettings(goodSettings()).provider, 'outlook');
  assert.equal(normalizeSettings({ ...goodSettings(), provider: ' Yahoo ' }).provider, 'yahoo');
  assert.equal(validateSettings({ ...goodSettings(), provider: 'yahoo' }).ok, true);
  assert.ok(validateSettings({ ...goodSettings(), provider: 'gmail' }).errors.provider);
  assert.deepEqual(Object.keys(PROVIDERS), ['outlook', 'yahoo']);
  assert.equal(PROVIDERS.yahoo.tabPattern, 'https://mail.yahoo.com/*');
  assert.equal(PROVIDERS.outlook.tabPattern, 'https://outlook.live.com/mail/*');
});
