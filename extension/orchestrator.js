// Job lifecycle coordinator. Event-driven (one `tick` per alarm / user action) and stateless
// between ticks: everything important is persisted so a suspended service worker is harmless.
// All side effects are injected, so the whole flow is unit-testable.

import { validateSettings, STATUS, UI_STATUS } from './config.js';
import { validateJob, ApiError } from './api.js';
import {
  PHASE, startRun, advanceRun, getCompleted, recordCompleted, classifyInterrupted, resultToReport, STAGE,
} from './state.js';

const DEFAULTS = { sendTimeoutMs: 100000, staleMs: 120000, pollSec: 60, pauseAfterFailures: 3 };
const NO_RECEIVER = /receiving end does not exist|could not establish connection/i;

const withTimeout = (p, ms) => new Promise((resolve, reject) => {
  const t = setTimeout(() => reject(new Error('Timed out waiting for the Outlook page.')), ms);
  p.then((v) => { clearTimeout(t); resolve(v); }, (e) => { clearTimeout(t); reject(e); });
});

export function createOrchestrator({ store, api, tabs, logger, now = Date.now, scheduleNext = () => {}, keepAlive = () => () => {}, random = Math.random, options = {} }) {
  const cfg = { ...DEFAULTS, ...options };
  let ticking = false;

  const setStatus = (code, message) => store.setStatus({ code, message, at: now() });
  const log = (e) => logger.log(e);

  // Random gap between backend fetches (default 10-20 min), drawn fresh after every fetch.
  const nextFetchDelayMs = (settings) => {
    const min = settings.fetchMinMinutes * 60000;
    const max = settings.fetchMaxMinutes * 60000;
    return Math.round(min + random() * (max - min));
  };
  const setNextClaim = (ms) => store.patchRuntime({ nextClaimAt: now() + ms });
  const scheduleFor = (nextClaimAt) => scheduleNext(nextClaimAt > now() ? Math.ceil((nextClaimAt - now()) / 1000) : cfg.pollSec);
  const clock = (t) => new Date(t).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  async function pauseAutomation() {
    await store.patchRuntime({ running: false });
  }

  async function flushPending(settings) {
    const p = await store.getPendingReport();
    if (!p) return true;
    try {
      await api.reportStatus(settings, p);
    } catch (e) {
      await log({ jobId: p.jobId, stage: 'report', ok: false, category: e.kind || 'unknown', message: e.message });
      if (e.kind === 'auth') await pauseAutomation();
      await setStatus(UI_STATUS.FAILED, `Result for the last job is saved locally but could not be reported: ${e.message} Will retry before taking new jobs.`);
      return false;
    }
    await store.clearPendingReport();
    const run = await store.getRun();
    if (run && run.jobId === p.jobId) {
      await store.clearRun();
      await store.clearInflight();
    }
    await log({ jobId: p.jobId, stage: 'report', ok: true, message: p.status });
    return true;
  }

  async function finalize(settings, job, status, message) {
    const t = now();
    // 1. Local record FIRST so a crash after this point can never lead to a re-send.
    await store.setCompletedMap(recordCompleted(await store.getCompletedMap(), job.id, status, t));
    // 2. Durable pending report, then try to deliver it.
    await store.setPendingReport({ jobId: job.id, status, message, timestamp: new Date(t).toISOString() });
    const run = (await store.getRun()) || startRun(job.id, t);
    await store.setRun(run.phase === PHASE.REPORTING ? { ...run, updatedAt: t } : advanceRun(run, PHASE.REPORTING, t));
    await store.setLastResult({ jobId: job.id, status, message, at: t });

    const rt = await store.getRuntime();
    const patch = { lastFinishedAt: t };
    let paused = false;
    if (status === STATUS.SENT) patch.consecutiveFailures = 0;
    else if (status === STATUS.NEEDS_REVIEW) { patch.running = false; paused = true; }
    else {
      patch.consecutiveFailures = (rt.consecutiveFailures || 0) + 1;
      if (patch.consecutiveFailures >= cfg.pauseAfterFailures) { patch.running = false; paused = true; }
    }
    await store.patchRuntime(patch);
    await log({ jobId: job.id, stage: 'finalize', ok: status === STATUS.SENT, category: status === STATUS.SENT ? null : status, message });

    const reported = await flushPending(settings);
    const suffix = (paused ? ' Automation paused: review before restarting.' : '') + (reported ? '' : ' (not yet reported to backend)');
    if (status === STATUS.SENT) await setStatus(UI_STATUS.SENT, `Email sent (verified). ${message}${suffix}`);
    else if (status === STATUS.NEEDS_REVIEW) await setStatus(UI_STATUS.NEEDS_REVIEW, `${message}${suffix}`);
    else await setStatus(UI_STATUS.FAILED, `${message}${suffix}`);
    scheduleFor((await store.getRuntime()).nextClaimAt || 0);
  }

  async function handleApiError(e, stage) {
    await log({ stage, ok: false, category: e.kind || 'unknown', message: e.message });
    if (e.kind === 'auth') {
      await pauseAutomation();
      await setStatus(UI_STATUS.FAILED, `${e.message} Automation paused; update the API key and press Start.`);
      return;
    }
    const wait = e.kind === 'rate_limit' ? (e.retryAfterMs ?? 60000) : cfg.pollSec * 1000;
    await setNextClaim(wait);
    scheduleNext(Math.ceil(wait / 1000));
    await setStatus(UI_STATUS.FAILED, `${e.message || 'Backend error.'} Will retry.`);
  }

  async function recoverRun(settings, run, inflight) {
    const rep = classifyInterrupted(inflight, run.jobId);
    await log({ jobId: run.jobId, stage: 'recover', ok: false, category: 'interrupted', message: rep.status });
    await finalize(settings, { id: run.jobId }, rep.status, rep.message);
  }

  async function runOnce() {
    const settings = await store.getSettings();
    const v = validateSettings(settings);
    if (!v.ok) {
      await setStatus(UI_STATUS.NOT_CONFIGURED, `Not configured: ${Object.values(v.errors)[0]}`);
      return { state: 'not_configured' };
    }
    const rt = await store.getRuntime();
    if (!settings.enabled || !rt.running) {
      await setStatus(UI_STATUS.PAUSED, settings.enabled ? 'Paused. Press Start to process jobs.' : 'Automation is disabled.');
      return { state: 'paused' };
    }

    if (!(await flushPending(settings))) { scheduleNext(cfg.pollSec); return { state: 'report_pending' }; }

    const run = await store.getRun();
    if (run) {
      const inflight = await store.getInflight();
      const mine = inflight && inflight.jobId === run.jobId ? inflight : null;
      const finished = mine && mine.stage === STAGE.DONE;
      const last = Math.max(run.updatedAt || 0, mine ? mine.at || 0 : 0);
      if (!finished && now() - last < cfg.staleMs) {
        await setStatus(UI_STATUS.PROCESSING, 'A job is still being processed.');
        scheduleNext(15);
        return { state: 'in_progress' };
      }
      await recoverRun(settings, run, mine);
      return { state: 'recovered' };
    }

    if (now() < (rt.nextClaimAt || 0)) {
      await setStatus(UI_STATUS.READY, `Ready. Next check for jobs at ${clock(rt.nextClaimAt)}.`);
      scheduleFor(rt.nextClaimAt);
      return { state: 'waiting' };
    }

    const tab = await tabs.findOutlookTab();
    if (!tab) {
      await setStatus(UI_STATUS.WAITING_OUTLOOK, 'Open Outlook Web (outlook.live.com/mail) in a tab and sign in.');
      scheduleNext(cfg.pollSec);
      return { state: 'no_tab' };
    }
    let ready;
    try { ready = await tabs.send(tab.id, { type: 'OEA_CHECK_READY' }); }
    catch { ready = { state: 'no_content_script' }; }
    if (!ready || ready.state !== 'ready') {
      const why = {
        loading: 'Outlook is still loading.', signed_out: 'Outlook is signed out; please sign in.',
        unexpected_page: 'The Outlook tab is not on the mail page.',
        no_content_script: 'The Outlook tab is not responding; reload it.',
      }[ready && ready.state] || 'Outlook is not ready.';
      await setStatus(UI_STATUS.WAITING_OUTLOOK, why);
      scheduleNext(cfg.pollSec);
      return { state: 'outlook_not_ready' };
    }

    let job;
    try { job = await api.claimJob(settings); }
    catch (e) { if (e instanceof ApiError) { await handleApiError(e, 'claim'); return { state: 'claim_error' }; } throw e; }
    const gap = nextFetchDelayMs(settings);
    await setNextClaim(gap); // one fetch per random 10-20 min window, whether or not a job came back
    if (!job) {
      await setStatus(UI_STATUS.READY, `Ready. No jobs waiting; next check at ${clock(now() + gap)}.`);
      scheduleNext(Math.ceil(gap / 1000));
      return { state: 'idle' };
    }
    await log({ jobId: job.id, stage: 'claim', ok: true });

    const check = validateJob(job);
    if (!check.ok) {
      if (typeof job.id === 'string' && job.id) {
        await finalize(settings, job, STATUS.FAILED, `Invalid job data: ${check.errors.join(' ')}`);
      } else {
        await log({ stage: 'validate', ok: false, category: 'invalid_job', message: check.errors.join(' ') });
        await setStatus(UI_STATUS.FAILED, `Backend returned a job without an id: ${check.errors.join(' ')}`);
      }
      return { state: 'invalid_job' };
    }

    // Only outcomes that may have SENT the email block a re-claimed job. A 'failed' job provably never
    // reached Send, so once the row is requeued (status cleared) it is safe, and expected, to try again.
    const prev = getCompleted(await store.getCompletedMap(), job.id);
    if (prev && prev.status === STATUS.FAILED) {
      await log({ jobId: job.id, stage: 'retry', ok: true, message: 'Retrying a job that previously failed before sending.' });
    } else if (prev) {
      await finalize(settings, job, prev.status, `Duplicate delivery: this job was already processed locally (${prev.status}). Not sending again.`);
      return { state: 'duplicate' };
    }

    try {
      const st = await api.getJobState(settings, job.id);
      if (st && [settings.statusValues[STATUS.SENT], settings.statusValues[STATUS.NEEDS_REVIEW]].includes(st)) {
        await store.setCompletedMap(recordCompleted(await store.getCompletedMap(), job.id, STATUS.NEEDS_REVIEW, now()));
        await setStatus(UI_STATUS.READY, 'Backend already has a final state for the job; skipped.');
        await log({ jobId: job.id, stage: 'state_check', ok: false, category: 'backend_final', message: st });
        scheduleNext(cfg.pollSec);
        return { state: 'backend_final' };
      }
    } catch (e) {
      await log({ jobId: job.id, stage: 'state_check', ok: false, category: e.kind || 'unknown', message: e.message });
    }

    await store.setRun(startRun(job.id, now()));
    await setStatus(UI_STATUS.PROCESSING, 'Processing email…');
    try { await api.reportStatus(settings, { jobId: job.id, status: STATUS.PROCESSING }, { retries: 1 }); }
    catch (e) { await log({ jobId: job.id, stage: 'report_processing', ok: false, category: e.kind || 'unknown', message: e.message }); }

    await store.setRun(advanceRun(await store.getRun(), PHASE.SENDING, now()));
    await store.clearInflight();
    const stopKeepAlive = keepAlive();
    let rep;
    try {
      const result = await withTimeout(tabs.send(tab.id, { type: 'OEA_SEND_JOB', job }), cfg.sendTimeoutMs);
      rep = resultToReport(result);
    } catch (e) {
      const inflight = await store.getInflight();
      if (!inflight && NO_RECEIVER.test(String(e && e.message))) {
        rep = { status: STATUS.FAILED, message: 'The Outlook tab did not receive the job; the email was not sent. Reload the Outlook tab.' };
      } else {
        rep = classifyInterrupted(inflight, job.id);
      }
      await log({ jobId: job.id, stage: 'send', ok: false, category: 'interrupted', message: String(e && e.message) });
    } finally {
      stopKeepAlive();
    }
    await finalize(settings, job, rep.status, rep.message);
    return { state: 'processed', status: rep.status };
  }

  return {
    async tick() {
      if (ticking) return { state: 'busy' };
      ticking = true;
      try {
        return await runOnce();
      } catch (e) {
        await log({ stage: 'tick', ok: false, category: 'internal', message: e && e.message });
        await setStatus(UI_STATUS.FAILED, `Unexpected error: ${e && e.message}`);
        return { state: 'error' };
      } finally {
        ticking = false;
      }
    },
    async start() {
      await store.patchRuntime({ running: true, consecutiveFailures: 0, nextClaimAt: 0 });
      return this.tick();
    },
    /**
     * Forget which jobs were already processed, so rows that were sent / needed review can be sent
     * again on purpose (after the status cell is cleared). Refused while a job or report is in flight.
     */
    async clearHistory() {
      if (await store.getRun()) return { ok: false, reason: 'A job is being processed right now. Try again in a minute.' };
      if (await store.getPendingReport()) return { ok: false, reason: 'A result is still waiting to be reported to the backend.' };
      await store.setCompletedMap({});
      await log({ stage: 'clear_history', ok: true, message: 'Local job history cleared by the user.' });
      return { ok: true };
    },
    async stop() {
      await store.patchRuntime({ running: false });
      await setStatus(UI_STATUS.PAUSED, 'Paused. Press Start to process jobs.');
    },
  };
}
