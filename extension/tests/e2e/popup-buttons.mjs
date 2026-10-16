// Manual real-Chromium check of the popup's Start/Stop buttons (colours, disabled states, instant toggle).
//   node extension/tests/e2e/popup-buttons.mjs        Env: PLAYWRIGHT_DIR, CHROME_PATH
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
const require = createRequire((process.env.PLAYWRIGHT_DIR || fileURLToPath(new URL('../../../', import.meta.url))) + '/');
const { chromium } = require('playwright');
const ext = fileURLToPath(new URL('../../', import.meta.url));
const ctx = await chromium.launchPersistentContext(mkdtempSync(join(tmpdir(), 'oea-btn-')), {
  headless: false, executablePath: process.env.CHROME_PATH || undefined,
  args: [`--disable-extensions-except=${ext}`, `--load-extension=${ext}`, '--no-sandbox', '--headless=new'],
});
let [sw] = ctx.serviceWorkers(); if (!sw) sw = await ctx.waitForEvent('serviceworker', { timeout: 15000 });
await new Promise((r) => setTimeout(r, 1500));
const base = { apiBaseUrl: 'http://localhost:3000', apiKey: 'k-12345', enabled: true };
const page = await ctx.newPage();
const url = `chrome-extension://${new URL(sw.url()).host}/popup.html`;
const look = () => page.evaluate(() => {
  const g = (id) => { const b = document.getElementById(id); const c = getComputedStyle(b); return { disabled: b.disabled, bg: c.backgroundColor, opacity: c.opacity, cursor: c.cursor }; };
  return { start: g('start'), stop: g('stop') };
});
let failed = false;
const check = (name, ok, detail) => { console.log(`${ok ? 'ok  ' : 'FAIL'} ${name}${ok ? '' : ' ' + JSON.stringify(detail)}`); if (!ok) failed = true; };
const state = async (settings, running, shot) => {
  await sw.evaluate(({ settings, running }) => chrome.storage.local.set({ settings, runtime: { running, consecutiveFailures: 0, lastFinishedAt: 0, nextClaimAt: 9e15 } }), { settings, running });
  await page.goto(url); await page.waitForTimeout(500);
  if (shot) await page.screenshot({ path: shot });
  return look();
};
let s = await state(base, false, process.env.SHOT_STOPPED);
check('stopped: Start enabled + green, Stop disabled + grey', !s.start.disabled && s.stop.disabled && s.start.bg !== s.stop.bg && /26, 127, 55|46, 160, 67/.test(s.start.bg) && s.stop.opacity < 1, s);
s = await state(base, true, process.env.SHOT_RUNNING);
check('running: Stop enabled + red, Start disabled + grey', s.start.disabled && !s.stop.disabled && /207, 34, 46|248, 81, 73/.test(s.stop.bg) && s.start.opacity < 1 && s.start.cursor === 'not-allowed', s);
s = await state({ ...base, enabled: false }, false);
check('automation disabled: both greyed out', s.start.disabled && s.stop.disabled, s);
// instant toggle on click (no waiting for the background worker)
await state(base, false);
await page.click('#start');
const afterStart = await look();
check('click Start: buttons flip immediately', afterStart.start.disabled && !afterStart.stop.disabled, afterStart);
await page.click('#stop');
const afterStop = await look();
check('click Stop: buttons flip back immediately', !afterStop.start.disabled && afterStop.stop.disabled, afterStop);
await ctx.close();
process.exit(failed ? 1 : 0);
