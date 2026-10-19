// Yahoo Mail page helpers (classic content script; attaches to globalThis.OEA).
// PLACEHOLDER: Yahoo Mail's real page structure has not been captured yet, so nothing here touches the page.
// Selectors and finders will be added from DOM samples taken on a real Yahoo Mail compose window.
(function (root) {
  'use strict';
  const OEA = (root.OEA = root.OEA || {});
  const YAHOO_MAIL_HOST = 'mail.yahoo.com';
  Object.assign(OEA, { YAHOO_MAIL_HOST });
})(typeof globalThis !== 'undefined' ? globalThis : window);
