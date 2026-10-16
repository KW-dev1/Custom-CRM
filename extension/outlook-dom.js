// Outlook Web DOM helpers (classic content script; attaches to globalThis.OEA).
// All selectors live in SELECTORS so they are easy to adjust when Outlook changes its markup.
// They prefer accessible names/labels over generated class names. English UI only.
(function (root) {
  'use strict';
  const OEA = (root.OEA = root.OEA || {});
  const MAIL_PREFIX = 'https://outlook.live.com/mail/';

  class OEAError extends Error {
    constructor(code, message) {
      super(message);
      this.name = 'OEAError';
      this.code = code;
    }
  }

  const SELECTORS = {
    newMail: ['button[aria-label="New mail"]', 'button[title="New mail"]', 'button[aria-label="New message"]'],
    to: ['[role="textbox"][aria-label="To"]', '[contenteditable="true"][aria-label="To"]', 'input[aria-label="To"]'],
    cc: ['[role="textbox"][aria-label="Cc"]', '[contenteditable="true"][aria-label="Cc"]', 'input[aria-label="Cc"]'],
    subject: ['input[aria-label="Subject"]', 'input[placeholder="Add a subject"]'],
    body: ['[role="textbox"][aria-label^="Message body"]', '[contenteditable="true"][aria-label^="Message body"]'],
    send: ['button[aria-label="Send"]', 'button[title="Send"]'],
    signedOut: ['input[name="loginfmt"]', 'a[data-task="signin"]', 'a[href*="login.live.com"]'],
    mailShell: ['[role="main"]', '[aria-label="Navigation pane"]', '[aria-label="Mail"]'],
    messageList: ['[role="listbox"][aria-label="Message list"]'],
    banners: ['[role="alert"]', '[role="alertdialog"]', '[role="dialog"]'],
    toasts: ['[role="status"]', '[role="alert"]', '[aria-live="polite"]', '[aria-live="assertive"]'],
  };
  const LABELS = { newMail: ['new mail', 'new message'], send: ['send'], ccToggle: ['cc'] };
  const SEND_ERROR_RE = /(couldn['’]t|can['’]t|could not|unable to|failed to)\s+(be\s+)?send|(wasn['’]t|was not|not)\s+sent|delivery (has )?failed/i;
  const SENT_TOAST_RE = /\b(message|email|mail)\s+(has been\s+|was\s+)?sent\b/i;

  function isVisible(el) {
    if (!el || !el.isConnected) return false;
    if (el.closest('[hidden],[aria-hidden="true"]')) return false;
    const view = el.ownerDocument.defaultView;
    const st = view && view.getComputedStyle ? view.getComputedStyle(el) : null;
    return !st || (st.display !== 'none' && st.visibility !== 'hidden');
  }
  function first(scope, selectors) {
    for (const sel of selectors) {
      const hit = Array.from(scope.querySelectorAll(sel)).find(isVisible);
      if (hit) return hit;
    }
    return null;
  }
  function all(scope, selectors) {
    const out = new Set();
    for (const sel of selectors) scope.querySelectorAll(sel).forEach((e) => isVisible(e) && out.add(e));
    return Array.from(out);
  }
  function byLabel(scope, labels) {
    const want = new Set(labels);
    return Array.from(scope.querySelectorAll('button,[role="button"]')).find((el) => {
      if (!isVisible(el)) return false;
      const names = [el.getAttribute('aria-label'), el.getAttribute('title'), el.textContent];
      return names.some((n) => n && want.has(n.trim().toLowerCase()));
    }) || null;
  }
  function lowestCommonAncestor(nodes) {
    let node = nodes[0];
    while (node && !nodes.every((n) => node.contains(n))) node = node.parentElement;
    return node;
  }

  // ---- finders ----
  const findNewMail = (doc) => first(doc, SELECTORS.newMail) || byLabel(doc, LABELS.newMail);
  const findSubject = (doc) => first(doc, SELECTORS.subject);
  const findCc = (scope) => first(scope, SELECTORS.cc);
  const findCcToggle = (scope) => byLabel(scope, LABELS.ccToggle);
  const countCompose = (doc) => all(doc, SELECTORS.subject).length;

  /** Locate the (single) open compose form: {root,to,subject,body,send} or null. */
  function findCompose(doc) {
    const subject = findSubject(doc);
    const body = first(doc, SELECTORS.body);
    if (!subject || !body) return null;
    let rootEl = lowestCommonAncestor([subject, body]);
    for (let i = 0; rootEl && i < 6; i++) {
      const to = first(rootEl, SELECTORS.to);
      const send = first(rootEl, SELECTORS.send) || byLabel(rootEl, LABELS.send);
      if (to && send) return { root: rootEl, to, subject, body, send };
      if (all(rootEl.parentElement || rootEl, SELECTORS.subject).length > 1) break; // would span two composes
      rootEl = rootEl.parentElement;
    }
    const to = first(doc, SELECTORS.to);
    const send = first(doc, SELECTORS.send) || byLabel(doc, LABELS.send);
    return to && send && countCompose(doc) === 1 ? { root: lowestCommonAncestor([to, subject, body, send]), to, subject, body, send } : null;
  }

  /** @returns {{state: 'ready'|'loading'|'signed_out'|'unexpected_page', message: string}} */
  function detectReadiness(doc, loc) {
    if (!String(loc.href).startsWith(MAIL_PREFIX)) {
      return { state: 'unexpected_page', message: 'The tab is not on https://outlook.live.com/mail/.' };
    }
    if (first(doc, SELECTORS.signedOut)) {
      return { state: 'signed_out', message: 'Outlook shows a sign-in page. Sign in manually, then retry.' };
    }
    if (doc.readyState === 'loading') return { state: 'loading', message: 'Outlook is still loading.' };
    if (findNewMail(doc) || findCompose(doc)) return { state: 'ready', message: 'Outlook is ready.' };
    return { state: 'loading', message: 'Outlook mail interface not detected yet.' };
  }

  // ---- waiting ----
  function waitFor(fn, { timeoutMs = 10000, description = 'condition', doc = root.document, intervalMs = 250 } = {}) {
    return new Promise((resolve, reject) => {
      let done = false;
      let obs = null;
      let poll = null;
      let timer = null;
      const finish = (err, val) => {
        if (done) return;
        done = true;
        if (obs) obs.disconnect();
        clearInterval(poll);
        clearTimeout(timer);
        if (err) reject(err); else resolve(val);
      };
      const check = () => {
        try { const v = fn(); if (v) finish(null, v); } catch (e) { finish(e); }
      };
      obs = new root.MutationObserver(check);
      obs.observe(doc.documentElement || doc, { subtree: true, childList: true, attributes: true, characterData: true });
      poll = setInterval(check, intervalMs);
      timer = setTimeout(() => finish(new OEAError('TIMEOUT', `Timed out after ${timeoutMs}ms waiting for ${description}.`)), timeoutMs);
      check();
    });
  }

  // ---- text ----
  const BLOCKS = new Set(['DIV', 'P', 'LI', 'BLOCKQUOTE', 'TR']);
  function extractText(el) {
    let out = '';
    (function walk(n) {
      if (n.nodeType === 3) { out += n.data; return; }
      if (n.nodeType !== 1) return;
      if (n.tagName === 'BR') { out += '\n'; return; }
      const block = BLOCKS.has(n.tagName);
      if (block && out && !out.endsWith('\n')) out += '\n';
      n.childNodes.forEach(walk);
      if (block && !out.endsWith('\n')) out += '\n';
    })(el);
    return out;
  }
  const squash = (s) => String(s).replace(/[\s ​]+/g, '').toLowerCase();
  const splitLines = (text) => String(text).split(/\r\n|\r|\n/);

  // ---- writing ----
  function fireInput(el, data, inputType) {
    const Ev = root.InputEvent || root.Event;
    el.dispatchEvent(new Ev('input', { bubbles: true, inputType: inputType || 'insertText', data }));
  }
  function setInputValue(input, value) {
    const proto = input.ownerDocument.defaultView.HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, 'value').set.call(input, value); // works with framework-controlled inputs
    input.dispatchEvent(new root.Event('input', { bubbles: true }));
    input.dispatchEvent(new root.Event('change', { bubbles: true }));
  }
  function selectContents(el) {
    const doc = el.ownerDocument;
    const sel = doc.defaultView.getSelection && doc.defaultView.getSelection();
    if (!sel) return;
    const range = doc.createRange();
    range.selectNodeContents(el);
    sel.removeAllRanges();
    sel.addRange(range);
  }
  function fillByDom(el, lines) {
    while (el.firstChild) el.removeChild(el.firstChild);
    const doc = el.ownerDocument;
    for (const line of lines) {
      const row = doc.createElement('div');
      if (line) row.appendChild(doc.createTextNode(line)); // text node only: never parsed as HTML
      else row.appendChild(doc.createElement('br'));
      el.appendChild(row);
    }
    fireInput(el, null, 'insertText');
  }
  /** Replace the content of a contenteditable with plain text, preserving line breaks. */
  function setEditableText(el, text) {
    const doc = el.ownerDocument;
    const lines = splitLines(text);
    el.focus();
    let method = 'dom';
    if (typeof doc.execCommand === 'function') {
      try {
        selectContents(el);
        let ok = doc.execCommand('delete');
        lines.forEach((line, i) => {
          if (i > 0) ok = doc.execCommand('insertParagraph') && ok;
          if (line) ok = doc.execCommand('insertText', false, line) && ok;
        });
        method = 'execCommand';
        if (!ok || squash(extractText(el)) !== squash(text)) method = 'dom';
      } catch { method = 'dom'; }
    }
    if (method === 'dom') fillByDom(el, lines);
    return method;
  }
  /**
   * Type text at the END of a contenteditable without touching its existing content. Recipient
   * fields hold already-resolved chips, so replacing the content would delete them.
   */
  function appendEditableText(el, text) {
    const doc = el.ownerDocument;
    el.focus();
    const sel = doc.defaultView.getSelection && doc.defaultView.getSelection();
    if (sel) {
      const range = doc.createRange();
      range.selectNodeContents(el);
      range.collapse(false);
      sel.removeAllRanges();
      sel.addRange(range);
    }
    if (typeof doc.execCommand === 'function') {
      const before = el.textContent;
      try {
        if (doc.execCommand('insertText', false, text) && el.textContent !== before) return 'execCommand';
      } catch { /* fall through to DOM insertion */ }
    }
    el.appendChild(doc.createTextNode(text));
    fireInput(el, text, 'insertText');
    return 'dom';
  }
  function pressKey(el, key, keyCode) {
    const init = { key, code: key, keyCode, which: keyCode, bubbles: true, cancelable: true };
    el.dispatchEvent(new root.KeyboardEvent('keydown', init));
    el.dispatchEvent(new root.KeyboardEvent('keyup', init));
  }

  // ---- recipient detection ----
  /**
   * A resolved recipient chip. In Outlook Web (verified against a real compose window) a resolved
   * address is a non-editable element INSIDE the To/Cc field: span[contenteditable=false]
   * [aria-label="<address>"]. Raw typed text is a plain text node and never matches. As a fallback
   * (other layouts) chips outside the field are also accepted.
   */
  function findRecipientPill(compose, email, field = compose.to) {
    const needle = email.toLowerCase();
    const hay = (el) => `${el.getAttribute('aria-label') || ''} ${el.getAttribute('title') || ''} ${el.textContent || ''}`.toLowerCase();
    for (const el of field.querySelectorAll('[contenteditable="false"]')) {
      if (hay(el).includes(needle) && (el.textContent || '').length < 200) return el;
    }
    const candidates = compose.root.querySelectorAll('[aria-label],[title],span,button,[role="button"],[role="option"]');
    for (const el of candidates) {
      if (el === field || field.contains(el) || el.contains(field)) continue;
      if (compose.body.contains(el) || el.contains(compose.body) || compose.subject === el) continue;
      if (hay(el).includes(needle) && (el.textContent || '').length < 200) return el;
    }
    return null;
  }
  const bannerText = (doc, re) => {
    for (const el of all(doc, SELECTORS.banners)) { const t = (el.textContent || '').trim(); if (re.test(t)) return t.slice(0, 200); }
    return null;
  };
  const findSendError = (doc) => bannerText(doc, SEND_ERROR_RE);
  const findRecipientError = (doc) => bannerText(doc, /couldn['’]t find|isn['’]t valid|invalid|can['’]t be resolved|not found/i);
  function findSentToast(doc) {
    for (const el of all(doc, SELECTORS.toasts)) if (SENT_TOAST_RE.test((el.textContent || '').trim())) return true;
    return false;
  }


  // ---- Sent Items confirmation (selectors verified against a real Outlook Web page) ----
  const folderName = (el) => String(el.getAttribute('data-folder-name') || '').toLowerCase();
  /** Folder tree entry, e.g. div[role=treeitem][data-folder-name="sent items"]. Appears in Favorites and under the account. */
  const findFolder = (doc, name) => Array.from(doc.querySelectorAll('[role="treeitem"][data-folder-name]')).find((el) => isVisible(el) && folderName(el) === name) || null;
  const findSelectedFolder = (doc) => Array.from(doc.querySelectorAll('[role="treeitem"][data-folder-name][aria-selected="true"]')).find(isVisible) || null;
  const isFolderSelected = (doc, name) => Array.from(doc.querySelectorAll('[role="treeitem"][data-folder-name][aria-selected="true"]')).some((el) => isVisible(el) && folderName(el) === name);

  /** Rows of the message list: label = "<recipient> <subject> <time> <preview>", plus the time span (title holds the full date). */
  function readMessageRows(doc) {
    const list = first(doc, SELECTORS.messageList);
    if (!list) return [];
    return Array.from(list.querySelectorAll('[role="option"]')).filter(isVisible).map((row) => {
      const timeSpan = Array.from(row.querySelectorAll('span')).find((sp) => /\d{1,2}:\d{2}/.test(sp.getAttribute('title') || '') || /^\d{1,2}:\d{2}(\s?[AP]M)?$/i.test((sp.textContent || '').trim()));
      return {
        label: (row.getAttribute('aria-label') || row.textContent || '').trim(),
        timeTitle: timeSpan ? timeSpan.getAttribute('title') || '' : '',
        timeText: timeSpan ? (timeSpan.textContent || '').trim() : '',
      };
    });
  }

  /** Row timestamp in ms, or null. Prefers the full date in the span title; falls back to today's clock text. */
  function parseRowTime(info, nowMs) {
    const full = Date.parse(info.timeTitle);
    // A parse that lands far from "now" is a day/month locale mix-up: ignore it and use the clock text.
    if (!Number.isNaN(full) && Math.abs(full - nowMs) < 36 * 3600 * 1000) return full;
    const m = /^(\d{1,2}):(\d{2})\s*([AP]M)?$/i.exec(info.timeText || '');
    if (!m) return null;
    let h = Number(m[1]);
    const ap = m[3] && m[3].toUpperCase();
    if (ap === 'PM' && h < 12) h += 12;
    if (ap === 'AM' && h === 12) h = 0;
    const d = new Date(nowMs);
    d.setHours(h, Number(m[2]), 0, 0);
    return d.getTime();
  }

  const normText = (t) => String(t).toLowerCase().replace(/\s+/g, ' ').trim();
  /**
   * Find the row for the message we just sent: same subject, sent within the last few minutes
   * (Outlook shows minutes only, so a one-minute margin either side), and, when the row shows an
   * email address rather than a contact's display name, the same recipient.
   */
  function matchSentRow(rows, { subject, to, sentAtMs, nowMs }) {
    const subj = normText(subject);
    const addr = String(to).toLowerCase();
    const from = sentAtMs - (sentAtMs % 60000) - 60000;
    for (const r of rows) {
      const label = normText(r.label);
      if (!label.includes(subj)) continue;
      if (label.includes('@') && !label.includes(addr)) continue;
      const t = parseRowTime(r, nowMs);
      if (t === null || t < from || t > nowMs + 60000) continue;
      return r;
    }
    return null;
  }

  Object.assign(OEA, {
    MAIL_PREFIX, OEAError, SELECTORS, isVisible, lowestCommonAncestor, findNewMail, findSubject, findCc, findCcToggle, countCompose, findCompose,
    detectReadiness, waitFor, extractText, squash, setInputValue, setEditableText, appendEditableText, pressKey, findRecipientPill,
    findSendError, findRecipientError, findSentToast,
    findFolder, findSelectedFolder, isFolderSelected, readMessageRows, parseRowTime, matchSentRow,
  });
})(typeof globalThis !== 'undefined' ? globalThis : window);
