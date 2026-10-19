import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { JSDOM } from 'jsdom';

const src = (f) => readFileSync(new URL(`../${f}`, import.meta.url), 'utf8');
function load(url) {
  const dom = new JSDOM('<!doctype html><body><button aria-label="Send">Send</button></body>', { url, runScripts: 'outside-only' });
  for (const f of ['yahoo-dom.js', 'yahoo-flow.js']) dom.window.eval(src(f));
  return dom;
}

test('Yahoo placeholder registers as the active provider', () => {
  const { window } = load('https://mail.yahoo.com/d/folders/1');
  assert.equal(window.OEA.activeProvider.name, 'yahoo');
});
test('on mail.yahoo.com readiness is "unsupported" (so no job is ever claimed); elsewhere it is unexpected_page', () => {
  const a = load('https://mail.yahoo.com/n/folders/1');
  assert.equal(a.window.OEA.activeProvider.detectReadiness(a.window.document, a.window.location).state, 'unsupported');
  const b = load('https://mail.yahoo.com.evil.example/');
  assert.equal(b.window.OEA.activeProvider.detectReadiness(b.window.document, b.window.location).state, 'unexpected_page');
});
test('the placeholder processJob fails safely before sending and touches nothing', async () => {
  const { window } = load('https://mail.yahoo.com/d/folders/1');
  let clicks = 0;
  window.document.querySelector('button').addEventListener('click', () => { clicks++; });
  const r = await window.OEA.activeProvider.processJob({ id: 'j', to: 'a@b.co', subject: 's', body: 'b' }, {});
  assert.equal(r.outcome, 'failed');
  assert.equal(r.sendClicked, false);
  assert.equal(r.code, 'PROVIDER_UNSUPPORTED');
  assert.equal(clicks, 0);
});
test('the Outlook flow registers itself as the active provider too', async () => {
  const { createOutlook } = await import('./fixtures/outlook.js');
  const o = await createOutlook();
  assert.equal(o.window.OEA.activeProvider.name, 'outlook');
  assert.equal(typeof o.window.OEA.activeProvider.processJob, 'function');
  assert.equal(o.window.OEA.activeProvider.detectReadiness(o.doc, o.loc).state, 'ready');
});
