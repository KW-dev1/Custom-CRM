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
 * The compose markup mirrors a REAL Outlook Web compose window (captured from the DOM):
 * To/Cc are contenteditable divs WITHOUT role=textbox, a resolved recipient is a
 * span[contenteditable=false][aria-label=<address>] INSIDE the field, Cc/Bcc are plain <button>s
 * with text only, and there is an (initially empty) role=alert element.
 *
 * opts.sendMode: 'close' | 'never' | 'error' | 'toast' | 'late-close'
 * opts.resolve: false => recipient never resolves; 'invalid' => Outlook shows an error
 * opts.resolveOn: 'enter' (default) | 'semicolon' (only a typed ";" resolves an address)
 * opts.pre: 'compose' => a draft is already open; 'signedOut' | 'noShell'
 * opts.noSend: omit the Send button
 * opts.noCcToggle: omit the "Cc" button; opts.ccResolve=false: Cc recipients never resolve
 */
export async function createOutlook(opts = {}) {
  const { sendMode = 'close', resolve = true, resolveOn = 'enter', pre = null, noSend = false, noCcToggle = false, ccResolve = true, url = 'https://outlook.live.com/mail/0/' } = opts;
  const dom = new JSDOM(pre === 'signedOut' ? '<body><a data-task="signin" href="#">Sign in</a></body>' : pre === 'noShell' ? '<body><p>Loading…</p></body>' : SHELL,
    { url, runScripts: 'outside-only', pretendToBeVisual: true });
  const { window } = dom;
  const doc = window.document;
  const c = { sendClicks: 0, newMailClicks: 0, ccTouched: false, ccToggleClicks: 0 };

  const openCompose = () => {
    const wrap = doc.createElement('div');
    wrap.id = 'compose';
    wrap.innerHTML = `
      <div class="toolbar">${noSend ? '' : '<button aria-label="Send" type="button">Send</button>'}<button aria-label="Discard" type="button">Discard</button></div>
      <div class="to-row"><button type="button" role="button">To</button>
        <div role="group"><div aria-label="To" contenteditable="true" tabindex="0"></div></div>
        ${noCcToggle ? '' : '<button type="button">Cc</button>'}<button type="button">Bcc</button></div>
      <div role="group" class="cc-row"><div aria-label="Cc" contenteditable="true" tabindex="0" hidden></div></div>
      <input aria-label="Subject" placeholder="Add a subject" type="text">
      <div role="textbox" aria-label="Message body" contenteditable="true" tabindex="0"></div>
      <div role="alert"></div>`;
    doc.body.appendChild(wrap);
    const alertBox = wrap.querySelector('[role="alert"]');
    const ccField = wrap.querySelector('[aria-label="Cc"][contenteditable]');
    const toggle = Array.from(wrap.querySelectorAll('button')).find((b) => b.textContent === 'Cc');
    if (toggle) toggle.addEventListener('click', () => { c.ccToggleClicks += 1; ccField.hidden = false; });
    ccField.addEventListener('input', () => { c.ccTouched = true; });

    const typed = (f) => Array.from(f.childNodes).filter((n) => n.nodeType === 3).map((n) => n.data).join('');
    const resolveField = (field, well, ok) => {
      const text = typed(field).replace(/;$/, '').trim();
      if (!text) return;
      if (ok === 'invalid') { alertBox.textContent = `We couldn't find a match for ${text}.`; return; }
      if (ok === false) return;
      Array.from(field.childNodes).filter((n) => n.nodeType === 3).forEach((n) => n.remove());
      const chip = doc.createElement('span');
      chip.setAttribute('contenteditable', 'false');
      chip.setAttribute('aria-label', text);
      chip.dataset.well = well;
      chip.textContent = text;
      field.appendChild(chip);
    };
    const wire = (field, well, ok) => {
      field.addEventListener('keydown', (ev) => { if (ev.key === 'Enter' && resolveOn === 'enter') resolveField(field, well, ok); });
      field.addEventListener('input', () => { if (typed(field).endsWith(';')) resolveField(field, well, ok); });
    };
    wire(wrap.querySelector('[aria-label="To"]'), 'to', resolve);
    wire(ccField, 'cc', ccResolve);

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

export const wellOf = (doc, email) => {
  const chip = Array.from(doc.querySelectorAll('span[contenteditable="false"][aria-label]')).find((x) => x.getAttribute('aria-label') === email);
  if (!chip) return null;
  return chip.closest('[aria-label="Cc"]') ? 'cc' : chip.closest('[aria-label="To"]') ? 'to' : null;
};

export function newGuard(window) {
  return window.eval('({ processed: new Set(), current: null })');
}
