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
  function pressKey(el, key, keyCode) {
    const init = { key, code: key, keyCode, which: keyCode, bubbles: true, cancelable: true };
    el.dispatchEvent(new root.KeyboardEvent('keydown', init));
    el.dispatchEvent(new root.KeyboardEvent('keyup', init));
  }

  // ---- recipient detection ----
  /** A resolved recipient chip inside the compose form (outside the To text box itself). */
  function findRecipientPill(compose, email, field = compose.to) {
    const needle = email.toLowerCase();
    const candidates = compose.root.querySelectorAll('[aria-label],[title],span,button,[role="button"],[role="option"]');
    for (const el of candidates) {
      if (el === field || field.contains(el) || el.contains(field)) continue;
      if (compose.body.contains(el) || el.contains(compose.body) || compose.subject === el) continue;
      const hay = `${el.getAttribute('aria-label') || ''} ${el.getAttribute('title') || ''} ${el.textContent || ''}`.toLowerCase();
      if (hay.includes(needle) && (el.textContent || '').length < 200) return el;
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

  Object.assign(OEA, {
    MAIL_PREFIX, OEAError, SELECTORS, isVisible, lowestCommonAncestor, findNewMail, findSubject, findCc, findCcToggle, countCompose, findCompose,
    detectReadiness, waitFor, extractText, squash, setInputValue, setEditableText, pressKey, findRecipientPill,
    findSendError, findRecipientError, findSentToast,
  });
})(typeof globalThis !== 'undefined' ? globalThis : window);
