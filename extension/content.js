// Content script entry (runs only on https://outlook.live.com/mail/*). Bridges chrome messages to
// OEA.processJob and persists progress so the service worker can recover safely after a restart.
(function () {
  'use strict';
  const OEA = globalThis.OEA;
  const guard = { processed: new Set(), current: null };

  async function markStage(jobId, stage, extra) {
    // Progress records contain no message content. A failed write aborts the flow before Send.
    await chrome.storage.local.set({ inflight: { jobId, stage, at: Date.now(), ...(extra || {}) } });
    console.debug(`[OEA] job=${jobId} stage=${stage}`);
  }

  chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (sender.id !== chrome.runtime.id || !msg || typeof msg.type !== 'string') return false;

    if (msg.type === 'OEA_CHECK_READY') {
      sendResponse(OEA.activeProvider.detectReadiness(document, location));
      return false;
    }
    if (msg.type === 'OEA_SEND_JOB') {
      (async () => {
        const job = msg.job;
        const jobId = job && typeof job.id === 'string' ? job.id : 'unknown';
        let result;
        try {
          const options = { verifySentItems: Boolean(msg.options && msg.options.verifySentItems === true) };
          result = await OEA.activeProvider.processJob(job, { doc: document, loc: location, guard, markStage, options });
        } catch (e) {
          result = { outcome: guard.processed.has(jobId) ? 'uncertain' : 'failed', sendClicked: guard.processed.has(jobId), code: 'UNEXPECTED', message: `Unexpected error: ${e && e.message}` };
        }
        console.debug(`[OEA] job=${jobId} outcome=${result.outcome} code=${result.code || '-'}`);
        try { await markStage(jobId, 'done', { result }); } catch { /* the response below still reaches the worker */ }
        sendResponse(result);
      })();
      return true; // async response
    }
    return false;
  });
})();
