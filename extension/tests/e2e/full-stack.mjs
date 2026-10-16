// Manual full-stack smoke test: REAL backend code (in-memory sheet instead of Google) +
// REAL extension in Chromium + SIMULATED Outlook page served at the real URL.
//   node extension/tests/e2e/full-stack.mjs
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

const require = createRequire((process.env.PLAYWRIGHT_DIR || fileURLToPath(new URL('../../../', import.meta.url))) + '/');
const { chromium } = require('playwright');
const ext = fileURLToPath(new URL('../../', import.meta.url));
const KEY = 'abcdefghijklmnop1234';

const sheet = new FakeSheet([
  { emails: '["first@client.com","second@client.com","third@client.com","fourth@client.com"]', subject: 'Quarterly follow-up', body: 'Hi team,\n\nFollowing up on our call.\n\nThanks' },
  { emails: 'solo@client.com', subject: 'Second row', body: 'Should wait for the next 10-20 min window.' },
]);
const cfg = testConfig();
const app = createApp({ service: createJobService({ sheet, cfg }), cfg, log: () => {} });
// Test harness only: headless Chrome cannot grant the optional host permission the popup requests
// on Save (which is what lets the real extension bypass CORS), so allow cross-origin calls here.
const harness = express();
harness.use((req, res, next) => {
  res.set({ 'access-control-allow-origin': '*', 'access-control-allow-headers': 'authorization,content-type', 'access-control-allow-methods': 'GET,POST,OPTIONS' });
  return req.method === 'OPTIONS' ? res.status(204).end() : next();
});
harness.use(app);
const server = await new Promise((r) => { const s = harness.listen(0, '127.0.0.1', () => r(s)); });
const port = server.address().port;

// Mirrors the structure captured from a real Outlook Web compose window: To/Cc are contenteditable divs
// (no role=textbox), resolved recipients are span[contenteditable=false][aria-label=address] INSIDE the field,
// Cc/Bcc are plain buttons, and Enter or ";" turns the typed text into a chip.
const PAGE = `<!doctype html><html><body><div role="main"><button aria-label="New mail">New mail</button></div>
<div role="tree"><div role="treeitem" data-folder-name="inbox" aria-selected="true"><span>Inbox</span></div><div role="treeitem" data-folder-name="sent items" aria-selected="false"><span>Sent Items</span></div></div>
<div role="region"><div role="listbox" aria-label="Message list" id="msglist"></div></div>
<script>
window.__sent = [];
window.__folderClicks = [];
const mb = { cur: 'inbox', sent: [] };
const pad = (n) => String(n).padStart(2, '0');
const clock = (d) => ((d.getHours() + 11) % 12 + 1) + ':' + pad(d.getMinutes()) + ' ' + (d.getHours() < 12 ? 'AM' : 'PM');
const full = (d) => ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][d.getDay()] + ' ' + (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear() + ' ' + clock(d);
function renderList() {
  const list = document.getElementById('msglist'); list.innerHTML = '';
  const rows = mb.cur === 'inbox' ? [{ who: 'news@example.com', subject: 'Weekly digest', at: new Date(Date.now() - 3600e3) }] : mb.sent;
  rows.forEach((m) => { const r = document.createElement('div'); r.setAttribute('role', 'option');
    r.setAttribute('aria-label', m.who + ' ' + m.subject + ' ' + clock(m.at) + ' preview');
    r.innerHTML = '<div><span></span><span></span><span class="t"></span><span>preview</span></div>';
    const sp = r.querySelectorAll('span'); sp[0].textContent = m.who; sp[1].textContent = m.subject; sp[2].textContent = clock(m.at); sp[2].setAttribute('title', full(m.at));
    list.appendChild(r); });
}
document.querySelectorAll('[role=treeitem]').forEach((it) => it.addEventListener('click', () => {
  mb.cur = it.getAttribute('data-folder-name'); window.__folderClicks.push(mb.cur);
  document.querySelectorAll('[role=treeitem]').forEach((x) => x.setAttribute('aria-selected', String(x === it))); renderList(); }));
renderList();
document.querySelector('[aria-label="New mail"]').addEventListener('click', () => setTimeout(() => {
  const w = document.createElement('div'); w.id = 'compose';
  w.innerHTML = '<div><button aria-label="Send" type="button">Send</button><button aria-label="Discard" type="button">Discard</button></div>'
    + '<div><button type="button" role="button">To</button><div role="group"><div aria-label="To" contenteditable="true" tabindex="0"></div></div><button type="button">Cc</button><button type="button">Bcc</button></div>'
    + '<div role="group"><div aria-label="Cc" contenteditable="true" tabindex="0" hidden></div></div>'
    + '<input aria-label="Subject" placeholder="Add a subject" type="text"><div role="textbox" aria-label="Message body" contenteditable="true" tabindex="0"></div><div role="alert"></div>';
  document.body.appendChild(w);
  const ccToggle = [...w.querySelectorAll('button')].find((b) => b.textContent === 'Cc');
  ccToggle.addEventListener('click', () => { w.querySelector('[aria-label="Cc"][contenteditable]').hidden = false; });
  const typed = (f) => [...f.childNodes].filter((n) => n.nodeType === 3).map((n) => n.data).join('');
  const resolve = (f) => { const t = typed(f).replace(/;$/, '').trim(); if (!t) return;
    [...f.childNodes].filter((n) => n.nodeType === 3).forEach((n) => n.remove());
    const chip = document.createElement('span'); chip.setAttribute('contenteditable', 'false'); chip.setAttribute('aria-label', t); chip.textContent = t; f.appendChild(chip); };
  for (const label of ['To', 'Cc']) {
    const f = w.querySelector('[aria-label="' + label + '"][contenteditable]');
    f.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); resolve(f); } });
    f.addEventListener('input', () => { if (typed(f).endsWith(';')) resolve(f); });
  }
  w.querySelector('button[aria-label="Send"]').addEventListener('click', () => {
    const chips = (label) => [...w.querySelectorAll('[aria-label="' + label + '"][contenteditable] span[contenteditable="false"]')].map((x) => x.getAttribute('aria-label'));
    setTimeout(() => { mb.sent.unshift({ who: chips('To')[0], subject: w.querySelector('input').value, at: new Date() }); if (mb.cur === 'sent items') renderList(); }, 300);
    window.__sent.push({ to: chips('To'), cc: chips('Cc'), subject: w.querySelector('input').value, body: w.querySelector('[aria-label^="Message body"]').innerText });
    setTimeout(() => w.remove(), 50);
  });
}, 20));
</script></body></html>`;

const ctx = await chromium.launchPersistentContext(mkdtempSync(join(tmpdir(), 'oea-fullstack-')), { // fresh profile each run (local duplicate memory would block a repeat)
  headless: false, executablePath: process.env.CHROME_PATH || undefined,
  args: [`--disable-extensions-except=${ext}`, `--load-extension=${ext}`, '--no-sandbox', '--headless=new'],
});
let [sw] = ctx.serviceWorkers(); if (!sw) sw = await ctx.waitForEvent('serviceworker', { timeout: 15000 });
const id = new URL(sw.url()).host;
await ctx.route('https://outlook.live.com/mail/**', (r) => r.fulfill({ contentType: 'text/html', body: PAGE }));
const outlook = await ctx.newPage();
await outlook.goto('https://outlook.live.com/mail/0/');
await outlook.waitForTimeout(500);
await sw.evaluate(({ port, KEY }) => chrome.storage.local.set({
  settings: { apiBaseUrl: `http://localhost:${port}`, apiKey: KEY, enabled: true, fetchMinMinutes: 10, fetchMaxMinutes: 20 },
  runtime: { running: false, consecutiveFailures: 0, lastFinishedAt: 0, nextClaimAt: 0 },
}), { port, KEY });
const popup = await ctx.newPage();
await popup.goto(`chrome-extension://${id}/popup.html`);
await popup.evaluate(() => chrome.runtime.sendMessage({ type: 'OEA_START' }));
for (let i = 0; i < 80 && sheet.rows[0].status !== 'sent'; i++) await popup.waitForTimeout(250);
await popup.waitForTimeout(500);

const sent = await outlook.evaluate(() => window.__sent);
const rt = await sw.evaluate(() => chrome.storage.local.get(['runtime', 'status']));
console.log('Outlook received:', JSON.stringify(sent));
console.log('folder clicks   :', JSON.stringify(await outlook.evaluate(() => window.__folderClicks)), '(expected sent items, then back to inbox)');
console.log('sheet statuses  :', JSON.stringify(sheet.rows.map((r) => r.status)));
console.log('sheet note[0]   :', sheet.rows[0].note);
console.log('extension status:', rt.status.code, '|', rt.status.message);
const gapMin = (rt.runtime.nextClaimAt - Date.now()) / 60000;
console.log('next fetch in   :', gapMin.toFixed(1), 'min (expected 10-20)');
const ok = sent.length === 1 && sent[0].to.join() === 'first@client.com' && sent[0].cc.join() === 'second@client.com,third@client.com,fourth@client.com'
  && sheet.rows[0].status === 'sent' && /Sent Items/.test(sheet.rows[0].note) && sheet.rows[1].status === '' && gapMin > 9.5 && gapMin <= 20.1;
console.log(ok ? 'RESULT: PASS' : 'RESULT: FAIL');
await ctx.close(); server.close();
process.exit(ok ? 0 : 1);
