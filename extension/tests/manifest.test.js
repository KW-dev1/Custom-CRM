import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';

const root = new URL('../', import.meta.url);
const m = JSON.parse(readFileSync(new URL('manifest.json', root), 'utf8'));

test('Manifest V3 with a popup and module service worker', () => {
  assert.equal(m.manifest_version, 3);
  assert.equal(m.action.default_popup, 'popup.html');
  assert.equal(m.background.service_worker, 'background.js');
});
test('permissions are minimal', () => {
  assert.deepEqual([...m.permissions].sort(), ['alarms', 'storage']);
  assert.deepEqual(m.host_permissions, ['https://outlook.live.com/*', 'https://mail.yahoo.com/*'], 'only the two supported webmail sites');
});
test('no broad required host access; optional hosts are https or localhost only', () => {
  for (const h of m.optional_host_permissions) assert.match(h, /^(https:\/\/\*\/\*|http:\/\/(localhost|127\.0\.0\.1)\/\*)$/);
  assert.ok(!JSON.stringify(m.host_permissions).includes('<all_urls>'));
});
test('content scripts are restricted to the two supported webmail sites, each with its own provider files', () => {
  assert.equal(m.content_scripts.length, 2);
  const by = Object.fromEntries(m.content_scripts.map((c) => [c.matches[0], c]));
  assert.deepEqual(m.content_scripts.map((c) => c.matches), [['https://outlook.live.com/mail/*'], ['https://mail.yahoo.com/*']]);
  assert.deepEqual(by['https://outlook.live.com/mail/*'].js, ['outlook-dom.js', 'outlook-flow.js', 'content.js']);
  assert.deepEqual(by['https://mail.yahoo.com/*'].js, ['outlook-dom.js', 'outlook-flow.js', 'yahoo-dom.js', 'yahoo-flow.js', 'content.js']);
});
test('every file referenced by the manifest exists', () => {
  const files = [m.action.default_popup, m.background.service_worker, ...m.content_scripts.flatMap((c) => c.js), ...Object.values(m.icons)];
  for (const f of files) assert.ok(existsSync(new URL(f, root)), f);
});
test('CSP forbids remote and inline script', () => {
  assert.ok(!/unsafe-eval|unsafe-inline|https?:/.test(m.content_security_policy.extension_pages));
});
test('source contains no eval / Function constructor / remote script loading / innerHTML', () => {
  for (const f of ['api.js', 'background.js', 'config.js', 'content.js', 'logger.js', 'orchestrator.js', 'outlook-dom.js', 'outlook-flow.js', 'yahoo-dom.js', 'yahoo-flow.js', 'popup.js', 'state.js', 'store.js']) {
    const src = readFileSync(new URL(f, root), 'utf8');
    assert.ok(!/\beval\s*\(|new Function\s*\(|importScripts|document\.write|\.innerHTML\s*=/.test(src), f);
  }
  assert.ok(!/<script[^>]+src=["']https?:/.test(readFileSync(new URL('popup.html', root), 'utf8')));
});
