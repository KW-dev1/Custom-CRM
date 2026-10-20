// A small simulated Gmail page (jsdom) built from DOM captured on a real Gmail compose window. It mimics the
// structure the extension relies on; it is NOT the real Gmail, so passing these tests is not proof it works live.
import { readFileSync } from 'node:fs';
import { JSDOM } from 'jsdom';

const src = (f) => readFileSync(new URL(`../../${f}`, import.meta.url), 'utf8');
export const SCRIPTS = ['outlook-dom.js', 'outlook-flow.js', 'gmail-dom.js', 'gmail-flow.js'].map(src);

/**
 * opts.sendMode: 'close' | 'toast' (stays open, notice) | 'never' | 'error'
 * opts.enterResolves: false => Enter does nothing, only a trailing "," resolves
 * opts.resolve: false => nothing ever resolves; opts.suggest: true => Enter also adds an unrelated chip
 * opts.noCcLink: omit the Cc link; opts.url; opts.pre: 'compose'
 */
export async function createGmail(opts = {}) {
  const { sendMode = 'close', enterResolves = true, resolve = true, suggest = false, noCcLink = false, pre = null, url = 'https://mail.google.com/mail/u/0/#inbox' } = opts;
  const dom = new JSDOM('<!doctype html><body><div role="button" gh="cm">Compose</div><div id="app"></div><div id="toasts"></div></body>', { url, runScripts: 'outside-only', pretendToBeVisual: true });
  const { window } = dom;
  const doc = window.document;
  const c = { sendClicks: 0, composeClicks: 0, ccLinkClicks: 0 };
  const chips = { To: [], Cc: [] };
  const st = { open: false, ccOpen: false, typed: { To: '', Cc: '' }, subject: '', bodyHtml: '' };
  const el = (tag, attrs = {}, text) => {
    const e = doc.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
    if (text !== undefined) e.textContent = text;
    return e;
  };
  const rootEl = el('div', { role: 'dialog' });

  function row(name) {
    const wrap = el('div');
    const box = el('div', { 'aria-label': name, name: name.toLowerCase() });
    const lb = el('div', { role: 'listbox', 'aria-label': 'Search Field' });
    for (const a of chips[name]) {
      const chip = el('div', { role: 'option', 'data-hovercard-id': a });
      chip.append(el('span', {}, a), el('div', { 'aria-hidden': 'true', 'aria-label': 'Press delete to remove this chip' }));
      lb.append(chip);
    }
    const input = el('input', { role: 'combobox', 'aria-label': `${name} recipients`, type: 'text' });
    input.value = st.typed[name];
    const add = (a) => { if (suggest) chips[name].push('suggested@other.com'); chips[name].push(a); st.typed[name] = ''; render(); };
    input.addEventListener('keydown', (ev) => { if (ev.key === 'Enter' && enterResolves && resolve && input.value.trim()) add(input.value.trim()); });
    input.addEventListener('input', () => { st.typed[name] = input.value; if (resolve && /,$/.test(input.value)) add(input.value.replace(/,$/, '')); });
    lb.append(el('div', { role: 'presentation' }));
    lb.lastChild.append(input);
    box.append(lb);
    wrap.append(box);
    return wrap;
  }
  function render() {
    const app = doc.getElementById('app');
    if (!st.open) { app.innerHTML = ''; return; }
    rootEl.textContent = '';
    rootEl.append(row('To'));
    if (!noCcLink && !st.ccOpen) {
      const link = el('span', { role: 'link', 'aria-label': 'Add Cc recipients (Ctrl-Shift-C)', tabindex: '1' }, 'Cc');
      link.addEventListener('click', () => { c.ccLinkClicks++; st.ccOpen = true; render(); });
      rootEl.append(link);
    }
    if (st.ccOpen) rootEl.append(row('Cc'));
    const subject = el('input', { name: 'subjectbox', 'aria-label': 'Subject', placeholder: 'Subject' });
    subject.value = st.subject;
    subject.addEventListener('input', () => { st.subject = subject.value; });
    const body = el('div', { role: 'textbox', 'aria-label': 'Message Body', contenteditable: 'true' });
    body.innerHTML = st.bodyHtml || '<div><br></div>';
    body.addEventListener('input', () => { st.bodyHtml = body.innerHTML; });
    const send = el('div', { role: 'button', 'aria-label': 'Send (Ctrl-Enter)' }, 'Send');
    const more = el('div', { role: 'button', 'aria-label': 'More send options' });
    let sent = false;
    send.addEventListener('mousedown', () => { sent = true; });
    send.addEventListener('click', () => { // a real Gmail button needs the mouse sequence, not only click
      if (!sent) return;
      c.sendClicks++;
      st.bodyText = body.innerText || body.textContent;
      if (sendMode === 'never') return;
      if (sendMode === 'error') { doc.getElementById('toasts').append(el('div', { role: 'alert' }, 'Your message could not be sent.')); return; }
      if (sendMode === 'toast') { doc.getElementById('toasts').append(el('span', {}, 'Message sent.')); return; }
      st.open = false; render();
      doc.getElementById('toasts').append(el('span', {}, 'Message sent.'));
    });
    rootEl.append(subject, body, el('table', { role: 'group' }), send, more, el('div', { role: 'button', 'aria-label': 'Discard draft (Ctrl-Shift-D)' }));
    if (!app.contains(rootEl)) app.append(rootEl);
  }
  doc.querySelector('[gh=cm]').addEventListener('mousedown', () => { st.pressed = true; });
  doc.querySelector('[gh=cm]').addEventListener('click', () => { c.composeClicks++; st.open = true; render(); });
  if (pre === 'compose') { st.open = true; render(); }
  for (const code of SCRIPTS) window.eval(code);
  return { dom, window, doc, loc: window.location, c, chips, st, OEA: window.OEA };
}
