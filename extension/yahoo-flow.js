// Yahoo Mail provider for the content script. PLACEHOLDER until the page structure is captured:
// it reports "unsupported", so the extension never claims a job for it and nothing is sent.
(function (root) {
  'use strict';
  const OEA = (root.OEA = root.OEA || {});

  const MESSAGE = 'Yahoo Mail automation is not implemented yet in this version.';

  OEA.activeProvider = {
    name: 'yahoo',
    detectReadiness(doc, loc) {
      if (loc.hostname !== OEA.YAHOO_MAIL_HOST) return { state: 'unexpected_page', message: 'The tab is not on mail.yahoo.com.' };
      return { state: 'unsupported', message: MESSAGE };
    },
    // Never reached while readiness says "unsupported", but fails safely (before any click) if it ever is.
    async processJob() {
      return { outcome: 'failed', sendClicked: false, code: 'PROVIDER_UNSUPPORTED', message: MESSAGE };
    },
  };
})(typeof globalThis !== 'undefined' ? globalThis : window);
