// A small simulated Yahoo Mail page (jsdom) built from DOM captured on a real Yahoo compose window. It mimics the
// structure the extension relies on; it is NOT the real Yahoo, so passing these tests is not proof it works live.
import { readFileSync } from 'node:fs';
import { JSDOM } from 'jsdom';

const src = (f) => readFileSync(new URL(`../../${f}`, import.meta.url), 'utf8');
export const SCRIPTS = ['outlook-dom.js', 'outlook-flow.js', 'yahoo-dom.js', 'yahoo-flow.js'].map(src);

const SHELL = '<!doctype html><body><nav><button type="button" aria-label="New message">New message</button></nav><div id="app"></div><div id="toasts"></div></body>';

/**
 * opts.sendMode: 'close' | 'toast' (stays open, toast) | 'never' | 'error'
 * opts.enterResolves: false => Enter does nothing, only a ';' resolves (default true)
 * opts.resolve: false => nothing ever resolves; opts.suggest: true => Enter also adds an unrelated suggestion chip
 * opts.collapse: true => rows collapse into "<first> +N more" when the subject changes (as the real page does)
 * opts.bodyEditor: 'normal' | 'flatten'; opts.pre: 'compose'; opts.url
 */
export async function createYahoo(opts = {}) {
  const { sendMode = 'close', enterResolves = true, resolve = true, suggest = false, collapse = true, ignoreSubject = false, bodyEditor = 'normal', pre = null, url = 'https://mail.yahoo.com/d/folders/1' } = opts;
  const dom = new JSDOM(SHELL, { url, runScripts: 'outside-only', pretendToBeVisual: true });
  const { window } = dom;
  const doc = window.document;
  const c = { sendClicks: 0, newMessageClicks: 0, ccToggleClicks: 0 };
  const chips = { To: [], Cc: [] };
  let open = false;
  const composeRoot = doc.createElement('div');
  composeRoot.setAttribute('role', 'main');
  composeRoot.setAttribute('aria-label', 'New message view');
  composeRoot.setAttribute('data-test-id', 'compose');

  const el = (tag, attrs = {}, text) => {
    const e = doc.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
    if (text !== undefined) e.textContent = text;
    return e;
  };
  const makePill = (addr) => {
    const wrap = el('div', { 'data-test-id': 'y-pill' });
    const cont = el('div', { 'data-test-id': 'pill-container' });
    cont.append(el('div', { role: 'button', 'aria-label': addr, tabindex: '0' }, addr), el('button', { type: 'button', 'aria-label': `Clear contact, ${addr}` }));
    wrap.append(cont);
    return wrap;
  };

  function render() {
    const app = doc.getElementById('app');
    if (!open) { app.innerHTML = ''; return; }
    const root = composeRoot; // the compose root persists across re-renders, like a framework-managed element
    root.textContent = '';
    const header = el('div');
    header.append(el('span', { role: 'heading' }, composeState.subject || 'New message'), el('button', { 'aria-label': 'Minimize', type: 'button' }));
    root.append(header);
    const addresses = [...chips.To, ...chips.Cc];
    if (collapse && composeState.collapsed && addresses.length) {
      const sum = el('div', { tabindex: '0' });
      sum.append(el('span', { 'aria-label': `<${addresses[0]}>`, title: `<${addresses[0]}>` }, addresses[0]));
      if (addresses.length > 1) sum.append(el('div', {}, `+${addresses.length - 1} more`));
      root.append(sum);
    } else {
      const rows = ['To', ...(composeState.ccOpen ? ['Cc', 'Bcc'] : [])];
      for (const name of rows) {
        const row = el('div');
        const pillBox = el('div');
        for (const a of chips[name] || []) pillBox.append(makePill(a));
        const input = el('input', { role: 'combobox', 'aria-label': name, type: 'text', id: `message-${name.toLowerCase()}-field` });
        input.value = composeState.typed[name] || '';
        input.addEventListener('keydown', (ev) => {
          if (ev.key !== 'Enter' || !enterResolves || !resolve || !input.value.trim()) return;
          if (suggest) chips[name].push('support@upwork.com');
          chips[name].push(input.value.trim());
          composeState.typed[name] = '';
          render();
          focusInput(name);
        });
        input.addEventListener('input', () => {
          composeState.typed[name] = input.value;
          if (resolve && /;$/.test(input.value)) { chips[name].push(input.value.replace(/;$/, '')); composeState.typed[name] = ''; render(); }
        });
        row.append(pillBox, input);
        if (name === 'To') {
          const toggle = el('button', { type: 'button' }, composeState.ccOpen ? 'Hide Cc Bcc' : 'Cc Bcc');
          toggle.addEventListener('click', () => { c.ccToggleClicks++; composeState.ccOpen = !composeState.ccOpen; render(); });
          row.append(toggle);
        }
        root.append(row);
      }
    }
    const subject = el('input', { 'aria-label': 'Subject', placeholder: 'Subject' });
    subject.value = composeState.subject || '';
    subject.addEventListener('input', () => { if (ignoreSubject) return; composeState.subject = subject.value; composeState.collapsed = true; const h = root.querySelector('[role=heading]'); if (h) h.textContent = subject.value || 'New message'; });
    const wrapper = el('div', { 'data-test-id': 'editorWrapper' });
    const body = el('div', { role: 'textbox', 'aria-label': 'Message body', contenteditable: 'true', 'data-test-id': 'rte', tabindex: '0' });
    body.innerHTML = composeState.bodyHtml || '<div><br></div>';
    body.addEventListener('input', () => { composeState.bodyHtml = body.innerHTML; });
    wrapper.append(body);
    const send = el('button', { 'aria-label': 'Send this email', 'data-test-id': 'compose-send-button' }, 'Send');
    send.addEventListener('click', () => {
      c.sendClicks++;
      composeState.bodyText = body.innerText || body.textContent;
      if (sendMode === 'never') return;
      if (sendMode === 'error') { doc.getElementById('toasts').append(el('div', { role: 'alert' }, 'Your message could not be sent.')); return; }
      if (sendMode === 'toast') { doc.getElementById('toasts').append(el('div', {}, 'Your message has been sent.')); return; }
      open = false;
      render();
      doc.getElementById('toasts').append(el('div', {}, 'Your message has been sent.'));
    });
    root.append(subject, wrapper, el('div', { 'data-test-id': 'compose-toolbar' }), send, el('span', { role: 'status' }, 'Saved'), el('button', { 'aria-label': 'Delete this draft', type: 'button' }));
    if (!app.contains(root)) app.append(root);
  }
  function focusInput(name) { const i = doc.querySelector(`input[aria-label="${name}"]`); if (i) i.focus(); }

  const composeState = { typed: {}, ccOpen: false, collapsed: false, subject: '', bodyHtml: '' };
  doc.querySelector('button[aria-label="New message"]').addEventListener('click', () => { c.newMessageClicks++; open = true; render(); });
  if (bodyEditor === 'flatten') {
    const proto = window.Document.prototype;
    const orig = proto.execCommand;
    proto.execCommand = function (cmd, ui, val) { return cmd === 'insertParagraph' || cmd === 'insertLineBreak' ? false : orig ? orig.call(this, cmd, ui, val) : false; };
  }
  if (pre === 'compose') { open = true; render(); }
  for (const code of SCRIPTS) window.eval(code);
  return { dom, window, doc, loc: window.location, c, chips, composeState, OEA: window.OEA };
}
