// Manual end-to-end smoke test in real Chrome (NOT part of `npm test`).
// Loads the unpacked extension, serves a SIMULATED Outlook page at the real URL via request
// interception, runs a mock backend on localhost, and prints what happened.
//   usage: node extension/tests/e2e/chrome-smoke.mjs [close|never]
// Requires playwright (global or local) and a Chromium binary: set CHROME_PATH and, if
// playwright is not resolvable from this repo, PLAYWRIGHT_DIR (e.g. /usr/lib/node_modules).
import { createRequire } from 'node:module';
import http from 'node:http';
import { fileURLToPath } from 'node:url';
const require = createRequire((process.env.PLAYWRIGHT_DIR || fileURLToPath(new URL('../../../', import.meta.url))) + '/');
const { chromium } = require('playwright');
const ext = fileURLToPath(new URL('../../', import.meta.url));
const CHROME = process.env.CHROME_PATH || undefined;
const MODE = process.argv[2] || 'close';

// ---- mock backend (CORS enabled so no host permission prompt is needed) ----
const reports = []; const claims = [];
const jobs = [{ id: 'job_e2e_1', to: 'client@example.com', subject: 'Follow-up', body: 'Hi,\n\nJust following up.\n\nThanks' }];
const server = http.createServer((req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'authorization,content-type');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  if (req.method === 'OPTIONS') { res.statusCode = 204; return res.end(); }
  let body = ''; req.on('data', (d) => (body += d)); req.on('end', () => {
    if (req.url === '/api/email-jobs/claim') {
      claims.push(req.headers.authorization);
      const j = jobs.shift();
      if (!j) { res.statusCode = 204; return res.end(); }
      res.setHeader('content-type', 'application/json'); return res.end(JSON.stringify({ job: j }));
    }
    if (/\/status$/.test(req.url)) { reports.push(JSON.parse(body)); res.setHeader('content-type', 'application/json'); return res.end('{}'); }
    res.statusCode = 404; res.end();
  });
}).listen(0);
const port = server.address().port;

// ---- simulated Outlook page served at the real URL ----
const PAGE = `<!doctype html><html><body><div role="main"><button aria-label="New mail">New mail</button></div><div id="t" role="status"></div>
<script>
const MODE = ${JSON.stringify(MODE)};
window.__sendClicks = 0;
document.querySelector('[aria-label="New mail"]').addEventListener('click', () => setTimeout(() => {
  const w = document.createElement('div'); w.id = 'compose';
  w.innerHTML = '<div><button aria-label="Send">Send</button></div><div class="r"></div><div role="textbox" aria-label="To" contenteditable="true"></div><input aria-label="Subject"><div role="textbox" aria-label="Message body, press Alt+F10 to exit" contenteditable="true"></div>';
  document.body.appendChild(w);
  const to = w.querySelector('[aria-label="To"]');
  to.addEventListener('keydown', (e) => { if (e.key !== 'Enter') return; const t = to.textContent.replace(/;$/,'').trim(); if (!t) return;
    const p = document.createElement('span'); p.setAttribute('role','button'); p.setAttribute('aria-label', t); p.textContent = t; w.querySelector('.r').appendChild(p); to.textContent=''; });
  w.querySelector('[aria-label="Send"]').addEventListener('click', () => { window.__sendClicks++; if (MODE==='close') setTimeout(() => w.remove(), 50); });
}, 20));
</script></body></html>`;

const ctx = await chromium.launchPersistentContext((process.env.TMPDIR || '/tmp') + '/oea-e2e-profile-' + MODE, {
  headless: false, executablePath: CHROME,
  args: [`--disable-extensions-except=${ext}`, `--load-extension=${ext}`, '--no-sandbox', '--headless=new'],
});
let [sw] = ctx.serviceWorkers(); if (!sw) sw = await ctx.waitForEvent('serviceworker', { timeout: 15000 });
const id = new URL(sw.url()).host;
await ctx.route('https://outlook.live.com/mail/**', (r) => r.fulfill({ contentType: 'text/html', body: PAGE }));
const outlook = await ctx.newPage();
await outlook.goto('https://outlook.live.com/mail/0/');
await outlook.waitForTimeout(500);

await sw.evaluate((port) => chrome.storage.local.set({
  settings: { apiBaseUrl: `http://localhost:${port}`, apiKey: 'test-key-123', enabled: true, fetchMinMinutes: 10, fetchMaxMinutes: 20 },
  runtime: { running: false, consecutiveFailures: 0, lastFinishedAt: 0, nextClaimAt: 0 },
}), port);
const popup = await ctx.newPage();
await popup.goto(`chrome-extension://${id}/popup.html`);
await popup.evaluate(() => chrome.runtime.sendMessage({ type: 'OEA_START' }));

for (let i = 0; i < 80 && reports.length < 2; i++) await popup.waitForTimeout(250);
await popup.waitForTimeout(500);
const st = await sw.evaluate(() => chrome.storage.local.get(['status', 'lastResult', 'completed', 'runtime', 'run', 'pendingReport']));
console.log('MODE', MODE);
console.log('reports:', JSON.stringify(reports.map((r) => ({ status: r.status, id: r.jobId, msg: r.message }))));
console.log('auth header:', claims[0]);
console.log('send clicks:', await outlook.evaluate(() => window.__sendClicks));
console.log('compose body:', JSON.stringify(await outlook.evaluate(() => document.querySelector('#compose [aria-label^="Message body"]')?.innerText ?? null)));
console.log('subject:', await outlook.evaluate(() => document.querySelector('#compose input')?.value ?? null));
console.log('status:', st.status.code, '|', st.status.message);
console.log('runtime:', JSON.stringify(st.runtime), 'run:', st.run, 'pending:', st.pendingReport);
await popup.reload(); await popup.waitForTimeout(400);
console.log('popup badge:', await popup.textContent('#badge'), '| last:', await popup.textContent('#lastResult'));
await ctx.close(); server.close();
