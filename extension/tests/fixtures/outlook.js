// A small simulated Outlook Web page for DOM-automation tests (jsdom). It mimics the structure
// the extension relies on; it is NOT the real Outlook, so passing these tests is not proof that
// automation works against the live site.
import { readFileSync } from 'node:fs';
import { JSDOM } from 'jsdom';

const src = (f) => readFileSync(new URL(`../../${f}`, import.meta.url), 'utf8');
export const SCRIPTS = ['outlook-dom.js', 'outlook-flow.js'].map(src);

const SHELL = `<!doctype html><html><body>
<div role="main"><button aria-label="New mail">New mail</button></div>
<div id="toasts" role="status"></div></body></html>`;

/**
 * opts.sendMode: 'close' | 'never' | 'error' | 'toast' | 'late-close'
 * opts.resolve: false => recipient never resolves; 'invalid' => Outlook shows an error
 * opts.pre: 'compose' => a draft is already open; 'signedOut' | 'noShell'
 * opts.noSend: omit the Send button
 */
export async function createOutlook(opts = {}) {
  const { sendMode = 'close', resolve = true, pre = null, noSend = false, url = 'https://outlook.live.com/mail/0/' } = opts;
  const dom = new JSDOM(pre === 'signedOut' ? '<body><a data-task="signin" href="#">Sign in</a></body>' : pre === 'noShell' ? '<body><p>Loading…</p></body>' : SHELL,
    { url, runScripts: 'outside-only', pretendToBeVisual: true });
  const { window } = dom;
  const doc = window.document;
  const c = { sendClicks: 0, newMailClicks: 0, ccTouched: false };

  const openCompose = () => {
    const wrap = doc.createElement('div');
    wrap.id = 'compose';
    wrap.innerHTML = `
      <div class="toolbar">${noSend ? '' : '<button aria-label="Send">Send</button>'}</div>
      <div class="recipients"></div>
      <div role="textbox" aria-label="To" contenteditable="true"></div>
      <div role="textbox" aria-label="Cc" contenteditable="true"></div>
      <input aria-label="Subject" placeholder="Add a subject">
      <div role="textbox" aria-label="Message body, press Alt+F10 to exit" contenteditable="true"></div>`;
    doc.body.appendChild(wrap);
    const to = wrap.querySelector('[aria-label="To"]');
    wrap.querySelector('[aria-label="Cc"]').addEventListener('input', () => { c.ccTouched = true; });
    to.addEventListener('keydown', (ev) => {
      if (ev.key !== 'Enter') return;
      const text = to.textContent.replace(/;$/, '').trim();
      if (!text) return;
      if (resolve === 'invalid') {
        const a = doc.createElement('div'); a.setAttribute('role', 'alert'); a.textContent = `We couldn't find a match for ${text}.`;
        wrap.appendChild(a); return;
      }
      if (resolve === false) return;
      const pill = doc.createElement('span');
      pill.setAttribute('role', 'button'); pill.setAttribute('aria-label', text); pill.textContent = text;
      wrap.querySelector('.recipients').appendChild(pill);
      to.textContent = '';
    });
    const send = wrap.querySelector('[aria-label="Send"]');
    if (send) send.addEventListener('click', () => {
      c.sendClicks += 1;
      if (sendMode === 'close') setTimeout(() => wrap.remove(), 20);
      if (sendMode === 'late-close') setTimeout(() => wrap.remove(), 900);
      if (sendMode === 'error') { const a = doc.createElement('div'); a.setAttribute('role', 'alert'); a.textContent = "We couldn't send your message. Try again."; doc.body.appendChild(a); }
      if (sendMode === 'toast') doc.getElementById('toasts').textContent = 'Your message has been sent.';
    });
  };
  doc.querySelectorAll('[aria-label="New mail"]').forEach((b) => b.addEventListener('click', () => { c.newMailClicks += 1; setTimeout(openCompose, 10); }));
  if (pre === 'compose') openCompose();

  if (doc.readyState !== 'complete') await new Promise((r) => window.addEventListener('load', r));
  for (const s of SCRIPTS) window.eval(s);
  return { dom, window, doc, OEA: window.OEA, controls: c, loc: window.location };
}

export function newGuard(window) {
  return window.eval('({ processed: new Set(), current: null })');
}
