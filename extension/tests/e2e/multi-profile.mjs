// Manual real-Chromium test of SEVERAL Chrome profiles (one Outlook account each) sharing ONE backend.
//   node extension/tests/e2e/multi-profile.mjs [profiles=3]
// Real backend code (in-memory sheet) + N real browser profiles, each with the real extension and its own
// SIMULATED Outlook page. All profiles press Start at the same moment. Checks that no row is processed
// twice, each profile sends something, the sender column names the right profile, and every row ends `sent`.
// Env: PLAYWRIGHT_DIR (dir containing node_modules/playwright), CHROME_PATH (chromium binary).
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import express from 'express';
import { createApp } from '../../../backend/src/app.js';
import { createJobService } from '../../../backend/src/jobs.js';
import { FakeSheet, testConfig } from '../../../backend/tests/_fakes.js';
import { PAGE } from './outlook-page.mjs';

const require = createRequire((process.env.PLAYWRIGHT_DIR || fileURLToPath(new URL('../../../', import.meta.url))) + '/');
const { chromium } = require('playwright');
const ext = fileURLToPath(new URL('../../', import.meta.url));
const KEY = 'abcdefghijklmnop1234';
const N = Number(process.argv[2] || 3);
const ROWS = N * 2;

const sheet = new FakeSheet(Array.from({ length: ROWS }, (_, i) => ({
  emails: `lead${i + 1}@client.com`, subject: `Hello lead ${i + 1}`, body: `Body for lead ${i + 1}`,
})));
const cfg = testConfig(); // MIN_CLAIM_INTERVAL_SEC=0 so profiles can claim back-to-back; pacing is covered by unit tests
const app = createApp({ service: createJobService({ sheet, cfg }), cfg, log: () => {} });
const harness = express(); // test harness only: allow cross-origin calls (headless Chrome cannot grant the optional host permission)
harness.use((req, res, next) => {
  res.set({ 'access-control-allow-origin': '*', 'access-control-allow-headers': 'authorization,content-type,x-client-id', 'access-control-allow-methods': 'GET,POST,OPTIONS' });
  return req.method === 'OPTIONS' ? res.status(204).end() : next();
});
harness.use(app);
const server = await new Promise((r) => { const s = harness.listen(0, '127.0.0.1', () => r(s)); });
const port = server.address().port;

async function startProfile(name) {
  const ctx = await chromium.launchPersistentContext(mkdtempSync(join(tmpdir(), `oea-multi-${name}-`)), {
    headless: false, executablePath: process.env.CHROME_PATH || undefined,
    args: [`--disable-extensions-except=${ext}`, `--load-extension=${ext}`, '--no-sandbox', '--headless=new'],
  });
  let [sw] = ctx.serviceWorkers(); if (!sw) sw = await ctx.waitForEvent('serviceworker', { timeout: 20000 });
  await new Promise((r) => setTimeout(r, 1200));
  await ctx.route('https://outlook.live.com/mail/**', (r) => r.fulfill({ contentType: 'text/html', body: PAGE }));
  const outlook = await ctx.newPage();
  await outlook.goto('https://outlook.live.com/mail/0/');
  await sw.evaluate(({ port, KEY, name }) => chrome.storage.local.set({
    settings: { apiBaseUrl: `http://localhost:${port}`, apiKey: KEY, profileName: name, enabled: true, fetchMinMinutes: 10, fetchMaxMinutes: 20, verifySentItems: true },
    runtime: { running: false, consecutiveFailures: 0, lastFinishedAt: 0, nextClaimAt: 0 },
  }), { port, KEY, name });
  const popup = await ctx.newPage();
  await popup.goto(`chrome-extension://${new URL(sw.url()).host}/popup.html`);
  return { name, ctx, sw, outlook, popup };
}

const profiles = [];
for (let i = 0; i < N; i++) profiles.push(await startProfile(`outlook-${String.fromCharCode(65 + i)}`));
// Everyone presses Start at (almost) the same moment.
await Promise.all(profiles.map((p) => p.popup.evaluate(() => chrome.runtime.sendMessage({ type: 'OEA_START' }))));
for (let i = 0; i < 160 && sheet.rows.filter((r) => r.status === 'sent').length < N; i++) await new Promise((r) => setTimeout(r, 250));
await new Promise((r) => setTimeout(r, 800));

const results = [];
for (const p of profiles) results.push({ name: p.name, sent: await p.outlook.evaluate(() => window.__sent.map((m) => m.to[0])) });
console.log('per-profile emails composed:', JSON.stringify(results));
console.log('sheet status / sender      :', JSON.stringify(sheet.rows.map((r) => `${r.status || '-'}/${r.sender || '-'}`)));
const all = results.flatMap((r) => r.sent);
const ok = all.length === N && new Set(all).size === N                                   // every claimed row sent exactly once, no overlap
  && results.every((r) => r.sent.length === 1)                                            // each profile did exactly one
  && sheet.rows.filter((r) => r.status === 'sent').length === N
  && sheet.rows.filter((r) => r.status === 'sent').every((r) => /^outlook-[A-Z]$/.test(r.sender))
  && new Set(sheet.rows.filter((r) => r.sender).map((r) => r.sender)).size === N         // each profile named exactly once
  && sheet.rows.filter((r) => !r.status).length === ROWS - N;                             // the rest untouched (next 10-20 min window)
console.log(ok ? 'RESULT: PASS' : 'RESULT: FAIL');
for (const p of profiles) await p.ctx.close();
server.close();
process.exit(ok ? 0 : 1);
