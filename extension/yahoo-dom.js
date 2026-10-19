// Yahoo Mail page helpers (classic content script; attaches to globalThis.OEA).
// Selectors were captured from a real Yahoo Mail compose window (English UI, current layout). They prefer
// aria-labels and data-test-id values over generated class names. Generic helpers (waitFor, setInputValue,
// setEditableText, ...) come from outlook-dom.js, which is loaded before this file.
//
// Yahoo specifics (verified from captures):
//  * To/Cc/Bcc are <input role=combobox aria-label=To|Cc|Bcc>; Enter turns the typed address into a chip:
//    div[data-test-id=pill-container] > div[role=button][aria-label=<address>] (+ a "Clear contact" button).
//  * Once focus leaves the recipient rows, Yahoo collapses them into "<first address> +N more" and the inputs
//    disappear, so recipients are verified right after each Enter, then again from the summary.
//  * Cc/Bcc rows are revealed by a plain "Cc Bcc" button (label becomes "Hide Cc Bcc" once open).
//  * After Send the compose view closes and a toast "Your message has been sent." appears.
(function (root) {
  'use strict';
  const OEA = (root.OEA = root.OEA || {});
  const YAHOO_MAIL_HOST = 'mail.yahoo.com';

  const SELECTORS = {
    newMail: ['button[aria-label="New message"]'],
    compose: ['[data-test-id="compose"]'],
    subject: 'input[aria-label="Subject"]',
    body: '[data-test-id="rte"][contenteditable="true"], [role="textbox"][aria-label="Message body"][contenteditable="true"]',
    send: 'button[data-test-id="compose-send-button"], button[aria-label="Send this email"]',
    pills: '[data-test-id="pill-container"] [role="button"][aria-label]',
    alerts: '[role="alert"], [role="alertdialog"], [role="dialog"], [role="status"], [data-test-id*="toast"], [data-test-id*="notification"]',
  };
  const SEND_ERROR_RE = /(couldn['’]t|can['’]t|could not|unable to|failed to)\s+(be\s+)?send|(wasn['’]t|was not|not)\s+sent|delivery (has )?failed|message (could not|couldn['’]t) be sent/i;
  const SENT_TOAST_RE = /your\s+message\s+has\s+been\s+sent/i;
  const RECIPIENT_ERROR_RE = /isn['’]t a valid|not a valid|invalid (email|address)|can['’]t be found|couldn['’]t find/i;

  const visible = (el) => OEA.isVisible(el);
  const norm = (t) => String(t || '').replace(/\s+/g, ' ').trim().toLowerCase();

  function findNewMail(doc) {
    for (const sel of SELECTORS.newMail) {
      const hit = Array.from(doc.querySelectorAll(sel)).find(visible);
      if (hit) return hit;
    }
    return null;
  }
  const composeRoots = (doc) => Array.from(doc.querySelectorAll(SELECTORS.compose.join(','))).filter(visible);
  const countCompose = (doc) => composeRoots(doc).length;

  /** The open compose window: { root, subject, body, send }. Recipient inputs are looked up on demand (they collapse). */
  function findCompose(doc) {
    const roots = composeRoots(doc);
    if (roots.length !== 1) return null;
    const rootEl = roots[0];
    const subject = Array.from(rootEl.querySelectorAll(SELECTORS.subject)).find(visible);
    const body = Array.from(rootEl.querySelectorAll(SELECTORS.body)).find(visible);
    const send = Array.from(rootEl.querySelectorAll(SELECTORS.send)).find(visible);
    if (!subject || !body || !send) return null;
    return { root: rootEl, subject, body, send };
  }

  /** The recipient input for 'To' | 'Cc' | 'Bcc', or null when it is not shown (not opened yet, or collapsed). */
  function findField(compose, name) {
    return Array.from(compose.root.querySelectorAll('input[role="combobox"]')).find((i) => i.getAttribute('aria-label') === name && visible(i)) || null;
  }
  /** The row of a recipient input: the largest ancestor that holds that input and no other recipient input. */
  function recipientRow(compose, input) {
    let best = input.parentElement;
    for (let n = input.parentElement; n && n !== compose.root; n = n.parentElement) {
      const others = Array.from(n.querySelectorAll('input[role="combobox"]')).filter((i) => i !== input);
      if (others.length || n.contains(compose.subject) || n.contains(compose.body)) break;
      best = n;
    }
    return best;
  }
  /** Addresses shown as chips in a recipient row (lower case). */
  function pillAddresses(compose, input) {
    const row = recipientRow(compose, input);
    return Array.from(row.querySelectorAll(SELECTORS.pills)).map((p) => norm(p.getAttribute('aria-label')));
  }
  function findRecipientPill(compose, email, input) {
    const row = recipientRow(compose, input);
    const needle = norm(email);
    return Array.from(row.querySelectorAll(SELECTORS.pills)).find((p) => norm(p.getAttribute('aria-label')) === needle) || null;
  }
  /** The "Cc Bcc" toggle (not the "Hide Cc Bcc" state). */
  function findCcToggle(compose) {
    return Array.from(compose.root.querySelectorAll('button,[role="button"]')).find((el) => visible(el) && norm(el.textContent) === 'cc bcc') || null;
  }
  /** After the rows collapse: the first address and the "+N more" count, or null when the summary is not shown. */
  function readRecipientSummary(compose) {
    const span = Array.from(compose.root.querySelectorAll('span[aria-label][title]')).find((s) => visible(s) && /@/.test(s.getAttribute('aria-label') || '') && !compose.body.contains(s));
    if (!span) return null;
    const m = /\+\s*(\d+)\s*more/i.exec(span.parentElement ? span.parentElement.textContent : '');
    return { first: norm(span.getAttribute('aria-label').replace(/[<>]/g, '')), more: m ? Number(m[1]) : 0 };
  }

  /**
   * Type into a Yahoo input the way a user does (a real editing command, so the page's own state sees it), falling
   * back to the framework-safe value setter. Ends with change + blur so the value is committed.
   */
  function typeIntoInput(input, value) {
    const doc = input.ownerDocument;
    input.focus();
    try { input.select(); } catch { /* not selectable */ }
    let used = false;
    try { used = typeof doc.execCommand === 'function' && doc.execCommand('insertText', false, value) === true; } catch { used = false; }
    if (!used || input.value !== value) OEA.setInputValue(input, value);
    input.dispatchEvent(new root.Event('change', { bubbles: true }));
    input.blur();
  }
  /** Type character by character with real key events (keydown, input, keyup), as a last resort for pages that ignore bulk edits. */
  function typeIntoInputByKeys(input, value) {
    const win = input.ownerDocument.defaultView;
    input.focus();
    OEA.setInputValue(input, '');
    let acc = '';
    for (const ch of value) {
      const init = { key: ch, bubbles: true, cancelable: true };
      input.dispatchEvent(new win.KeyboardEvent('keydown', init));
      input.dispatchEvent(new win.KeyboardEvent('keypress', init));
      acc += ch;
      Object.getOwnPropertyDescriptor(win.HTMLInputElement.prototype, 'value').set.call(input, acc);
      input.dispatchEvent(new (win.InputEvent || win.Event)('input', { bubbles: true, inputType: 'insertText', data: ch }));
      input.dispatchEvent(new win.KeyboardEvent('keyup', init));
    }
    input.dispatchEvent(new win.Event('change', { bubbles: true }));
    input.blur();
  }
  /** True when the compose title shows the subject (Yahoo's own state has it). */
  function subjectRegistered(compose, subject) {
    const heading = readHeading(compose);
    if (heading === null) return true; // no title element to judge by
    const want = String(subject).trim().toLowerCase().replace(/\s+/g, ' ');
    return heading.includes(want.slice(0, 25)) || want.includes(heading);
  }
  /** The compose header title; Yahoo shows the subject there once ITS state has it (before that: "New message"). */
  function readHeading(compose) {
    const h = Array.from(compose.root.querySelectorAll('[role="heading"]')).find(visible);
    return h ? norm(h.textContent) : null;
  }

  function detectReadiness(doc, loc) {
    if (loc.hostname !== YAHOO_MAIL_HOST) return { state: 'unexpected_page', message: 'The tab is not on mail.yahoo.com.' };
    if (doc.readyState === 'loading') return { state: 'loading', message: 'Yahoo Mail is still loading.' };
    if (findNewMail(doc) || countCompose(doc) > 0) return { state: 'ready', message: 'Yahoo Mail is ready.' };
    return { state: 'loading', message: 'Yahoo Mail interface not detected yet (is the mailbox open and signed in?).' };
  }

  function bannerText(doc, re) {
    for (const el of doc.querySelectorAll(SELECTORS.alerts)) {
      if (!visible(el)) continue;
      const t = (el.textContent || '').trim();
      if (re.test(t)) return t.slice(0, 200);
    }
    return null;
  }
  const findSendError = (doc) => bannerText(doc, SEND_ERROR_RE);
  const findRecipientError = (doc) => bannerText(doc, RECIPIENT_ERROR_RE);
  /** The "Your message has been sent." toast (any visible element; its container is not stable). */
  function findSentToast(doc) {
    const body = doc.body;
    if (!body || !SENT_TOAST_RE.test(body.textContent || '')) return false;
    return Array.from(body.querySelectorAll('*')).some((el) => (el.textContent || '').length < 80 && SENT_TOAST_RE.test(el.textContent) && visible(el));
  }

  Object.assign(OEA, {
    YAHOO_MAIL_HOST,
    yahoo: {
      SELECTORS, findNewMail, countCompose, findCompose, findField, recipientRow, pillAddresses, findRecipientPill, findCcToggle,
      readRecipientSummary, typeIntoInput, typeIntoInputByKeys, subjectRegistered, readHeading, detectReadiness, findSendError, findRecipientError, findSentToast,
    },
  });
})(typeof globalThis !== 'undefined' ? globalThis : window);
