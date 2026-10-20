// Gmail page helpers (classic content script; attaches to globalThis.OEA). Generic helpers (waitFor, setInputValue,
// setEditableText, ...) come from outlook-dom.js, loaded before this file.
// Selectors come from DOM captures of a real Gmail compose window (English UI, current layout):
//  * Subject: input[name=subjectbox]; body: div[role=textbox][aria-label="Message Body"][contenteditable=true].
//  * Send: div[role=button][aria-label^="Send"] (the label has a "(Ctrl-Enter)" hint); "More send options" is separate.
//  * To/Cc: input[role=combobox][aria-label="To recipients" | "Cc recipients"]; a resolved recipient is a
//    div[role=option] chip in the same listbox as the input. Cc/Bcc rows are opened with span[role=link][aria-label^="Add Cc recipients"].
//  * After Send the compose window closes and a "Message sent." notice appears.
// Recipient chips are matched by their text/data-hovercard-id, so small markup changes do not break the match.
(function (root) {
  'use strict';
  const OEA = (root.OEA = root.OEA || {});
  const GMAIL_MAIL_HOST = 'mail.google.com';

  const SEND_ERROR_RE = /(couldn['’]t|can['’]t|could not|unable to|failed to)\s+(be\s+)?send|(wasn['’]t|was not|not)\s+sent|delivery (has )?failed|message (could not|couldn['’]t) be sent/i;
  const SENT_TOAST_RE = /message\s+sent\b/i;
  const RECIPIENT_ERROR_RE = /isn['’]t a valid|not a valid|invalid (email|address)|couldn['’]t find|no match/i;
  const ALERTS = '[role="alert"], [role="alertdialog"], [role="dialog"], [role="status"], [aria-live]';

  const visible = (el) => OEA.isVisible(el);
  const norm = (t) => String(t || '').replace(/[‎‏‪-‮]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();

  const subjectInputs = (doc) => Array.from(doc.querySelectorAll('input[name="subjectbox"]')).filter(visible);
  const countCompose = (doc) => subjectInputs(doc).length;

  function findNewMail(doc) {
    const byGh = Array.from(doc.querySelectorAll('[gh="cm"]')).find(visible);
    if (byGh) return byGh;
    return Array.from(doc.querySelectorAll('[role="button"],button')).find((el) => visible(el) && norm(el.getAttribute('aria-label') || el.textContent) === 'compose') || null;
  }

  /** The open compose window: { root, subject, body, send }. */
  function findCompose(doc) {
    const subs = subjectInputs(doc);
    if (subs.length !== 1) return null;
    const subject = subs[0];
    let rootEl = subject;
    let body = null;
    let send = null;
    for (let i = 0; i < 14 && rootEl.parentElement; i++) {
      rootEl = rootEl.parentElement;
      body = Array.from(rootEl.querySelectorAll('div[role="textbox"][contenteditable="true"]')).find((e) => visible(e) && /message body/i.test(e.getAttribute('aria-label') || ''));
      send = Array.from(rootEl.querySelectorAll('[role="button"]')).find((e) => visible(e) && /^send\b/i.test(norm(e.getAttribute('aria-label') || '')) && !/options/i.test(e.getAttribute('aria-label') || ''));
      if (body && send) break;
    }
    if (!body || !send) return null;
    return { root: rootEl, subject, body, send };
  }

  const FIELD_LABELS = { To: 'to recipients', Cc: 'cc recipients', Bcc: 'bcc recipients' };
  /** The recipient input for 'To' | 'Cc' | 'Bcc', or null when it is not shown. */
  function findField(compose, name) {
    return Array.from(compose.root.querySelectorAll('input[role="combobox"],input[type="text"]')).find((i) => norm(i.getAttribute('aria-label')) === FIELD_LABELS[name] && visible(i)) || null;
  }
  /** The element holding a field's chips: div[name=to|cc|bcc] (verified), else the input's listbox, else its nearest ancestor. */
  function chipScope(compose, input) {
    const key = Object.keys(FIELD_LABELS).find((k) => FIELD_LABELS[k] === norm(input.getAttribute('aria-label')));
    const named = key && Array.from(compose.root.querySelectorAll(`div[name="${key.toLowerCase()}"]`)).find((d) => d.querySelector('input') === input || d.contains(input) || d.parentElement.contains(input));
    if (named && named.querySelector('[role="option"]') !== null) return named;
    const lb = input.closest('[role="listbox"]');
    if (lb && compose.root.contains(lb)) return lb;
    let best = input.parentElement;
    for (let n = input.parentElement; n && n !== compose.root; n = n.parentElement) {
      const others = Array.from(n.querySelectorAll('input[role="combobox"]')).filter((i) => i !== input);
      if (others.length || n.contains(compose.subject) || n.contains(compose.body)) break;
      best = n;
    }
    return best;
  }
  const chipText = (chip) => norm(`${chip.getAttribute('data-hovercard-id') || ''} ${chip.getAttribute('email') || ''} ${chip.textContent || ''}`);
  const chips = (compose, input) => Array.from(chipScope(compose, input).querySelectorAll('[role="option"]')).filter((c) => !c.querySelector('input'));
  /** Number of chips in a field, and whether each wanted address is shown in one. */
  function chipState(compose, input, wanted) {
    const list = chips(compose, input);
    return { count: list.length, found: wanted.map((a) => list.some((c) => chipText(c).includes(norm(a)))) };
  }
  function findRecipientPill(compose, email, input) {
    const needle = norm(email);
    return chips(compose, input).find((c) => chipText(c).includes(needle)) || null;
  }
  /** Gmail keeps its own copy of the subject in a hidden input[name=subject]; null when that input is not there. */
  function readSubjectMirror(compose) {
    const h = compose.root.querySelector('input[type="hidden"][name="subject"]');
    return h ? h.value : null;
  }
  const subjectRegistered = (compose, subject) => { const m = readSubjectMirror(compose); return m === null || m === subject; };
  /** The "Cc" link that opens the Cc row. */
  function findCcToggle(compose) {
    return Array.from(compose.root.querySelectorAll('[role="link"],[role="button"],span')).find((el) => visible(el) && /^add cc recipients/.test(norm(el.getAttribute('aria-label') || ''))) || null;
  }

  /** A full mouse sequence: Gmail's div[role=button] controls react to mousedown/mouseup, not only click(). */
  function pressButton(el) {
    const win = el.ownerDocument.defaultView;
    for (const type of ['mousedown', 'mouseup', 'click']) el.dispatchEvent(new win.MouseEvent(type, { bubbles: true, cancelable: true, view: win }));
  }

  function detectReadiness(doc, loc) {
    if (loc.hostname !== GMAIL_MAIL_HOST || !/^\/mail\//.test(loc.pathname)) return { state: 'unexpected_page', message: 'The tab is not on mail.google.com/mail.' };
    if (doc.readyState === 'loading') return { state: 'loading', message: 'Gmail is still loading.' };
    if (findNewMail(doc) || countCompose(doc) > 0) return { state: 'ready', message: 'Gmail is ready.' };
    return { state: 'loading', message: 'Gmail interface not detected yet (is the mailbox open and signed in?).' };
  }

  function bannerText(doc, re) {
    for (const el of doc.querySelectorAll(ALERTS)) {
      if (!visible(el)) continue;
      const t = (el.textContent || '').trim();
      if (re.test(t)) return t.slice(0, 200);
    }
    return null;
  }
  const findSendError = (doc) => bannerText(doc, SEND_ERROR_RE);
  const findRecipientError = (doc) => bannerText(doc, RECIPIENT_ERROR_RE);
  function findSentToast(doc) {
    const body = doc.body;
    if (!body) return false;
    // Short elements only: a whole-page text match could be glued across elements ("Send" + "Message sent").
    return Array.from(body.querySelectorAll('*')).some((el) => (el.textContent || '').length < 60 && SENT_TOAST_RE.test(el.textContent) && visible(el));
  }

  Object.assign(OEA, {
    GMAIL_MAIL_HOST,
    gmail: { countCompose, findNewMail, findCompose, findField, chipState, findRecipientPill, findCcToggle, readSubjectMirror, subjectRegistered, pressButton, detectReadiness, findSendError, findRecipientError, findSentToast },
  });
})(typeof globalThis !== 'undefined' ? globalThis : window);
