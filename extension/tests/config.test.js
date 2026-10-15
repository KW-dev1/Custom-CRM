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
