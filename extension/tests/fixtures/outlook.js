// A small simulated Outlook Web page for DOM-automation tests (jsdom). It mimics the structure
// the extension relies on; it is NOT the real Outlook, so passing these tests is not proof that
// automation works against the live site.
import { readFileSync } from 'node:fs';
import { JSDOM } from 'jsdom';

const src = (f) => readFileSync(new URL(`../../${f}`, import.meta.url), 'utf8');
export const SCRIPTS = ['outlook-dom.js', 'outlook-flow.js'].map(src);

const SHELL = `<!doctype html><html><body>
<div role="main"><button aria-label="New mail">New mail</button></div>
<div id="toasts" role="status"></div>
<div role="tree" id="tree">
  <div role="treeitem" data-folder-name="inbox" aria-selected="true" title="Inbox"><span>Inbox</span></div>
  <div role="treeitem" data-folder-name="sent items" aria-selected="false" title="Sent Items - 0 items"><span>Sent Items</span></div>
</div>
<div role="region"><div role="listbox" aria-label="Message list" id="msglist"></div></div></body></html>`;

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
 * opts.bodyEditor: 'normal' | 'flatten' (an editor that turns every inserted paragraph into plain text: line breaks are lost)
 * opts.sentItems: how the Sent Items folder behaves after Send:
 *   'normal' (row appears) | 'delayed' (row appears after 700ms) | 'missing' (never) | 'old' (only an old row with
 *   the same subject) | 'name' (shows a display name, not an address) | 'wrong-recipient' | 'no-folder'
 */
export async function createOutlook(opts = {}) {
  const { sendMode = 'close', resolve = true, resolveOn = 'enter', pre = null, noSend = false, noCcToggle = false, ccResolve = true, sentItems = 'normal', bodyEditor = 'normal', url = 'https://outlook.live.com/mail/0/' } = opts;
  const dom = new JSDOM(pre === 'signedOut' ? '<body><a data-task="signin" href="#">Sign in</a></body>' : pre === 'noShell' ? '<body><p>Loading…</p></body>' : SHELL,
    { url, runScripts: 'outside-only', pretendToBeVisual: true });
  const { window } = dom;
  const doc = window.document;
  const c = { sendClicks: 0, newMailClicks: 0, ccTouched: false, ccToggleClicks: 0, folderClicks: { sent: 0, inbox: 0 } };
  const mailbox = { current: 'inbox', sent: [] };

  // ---- folder tree + message list ----
  const pad = (n) => String(n).padStart(2, '0');
  const clock = (d) => `${((d.getHours() + 11) % 12) + 1}:${pad(d.getMinutes())} ${d.getHours() < 12 ? 'AM' : 'PM'}`;
  const fullDate = (d) => `${['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][d.getDay()]} ${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()} ${clock(d)}`;
  const renderList = () => {
    const list = doc.getElementById('msglist');
    if (!list) return;
    const rows = mailbox.current === 'inbox'
      ? [{ who: 'news@example.com', subject: 'Weekly digest', at: new Date(Date.now() - 3600e3) }]
      : mailbox.sent;
    list.innerHTML = '';
    rows.forEach((m, i) => {
      const row = doc.createElement('div');
      row.setAttribute('role', 'option');
      row.setAttribute('aria-label', `${m.who} ${m.subject} ${clock(m.at)} preview text`);
      row.setAttribute('aria-posinset', String(i + 1));
      row.innerHTML = '<div><span></span><span></span><span class="t"></span><span>preview text</span></div>';
      const [a, b, t] = row.querySelectorAll('span');
      a.textContent = m.who; b.textContent = m.subject; t.textContent = clock(m.at); t.setAttribute('title', fullDate(m.at));
      list.appendChild(row);
    });
  };
  doc.querySelectorAll('[role="treeitem"]').forEach((item) => {
    if (sentItems === 'no-folder' && item.getAttribute('data-folder-name') === 'sent items') item.remove();
    else item.addEventListener('click', () => {
      const name = item.getAttribute('data-folder-name') === 'sent items' ? 'sent' : 'inbox';
      c.folderClicks[name] += 1;
      mailbox.current = name === 'sent' ? 'sent items' : 'inbox';
      doc.querySelectorAll('[role="treeitem"]').forEach((x) => x.setAttribute('aria-selected', String(x === item)));
      renderList();
    });
  });
  renderList();
  const recordSent = (to, subject) => {
    const add = (m) => { mailbox.sent.unshift(m); if (mailbox.current === 'sent items') renderList(); };
    const now = new Date();
    if (sentItems === 'normal') add({ who: to, subject, at: now });
    if (sentItems === 'name') add({ who: 'Austin Ames', subject, at: now });
    if (sentItems === 'wrong-recipient') add({ who: 'someone.else@example.com', subject, at: now });
    if (sentItems === 'old') add({ who: to, subject, at: new Date(Date.now() - 2 * 3600e3) });
    if (sentItems === 'delayed') setTimeout(() => add({ who: to, subject, at: new Date() }), 700);
  };

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
    if (bodyEditor === 'flatten') {
      const body = wrap.querySelector('[role="textbox"][aria-label^="Message body"]');
      body.appendChild = (node) => window.Node.prototype.appendChild.call(body, doc.createTextNode(node.textContent)); // paragraphs/breaks are lost
    }
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
      const toChip = wrap.querySelector('[aria-label="To"] span[contenteditable="false"]');
      recordSent(toChip ? toChip.getAttribute('aria-label') : '?', wrap.querySelector('input[aria-label="Subject"]').value);
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
