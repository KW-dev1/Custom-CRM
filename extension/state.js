// Job run state machine, duplicate protection and interrupted-run classification. Pure module.

import { STATUS } from './config.js';

export const PHASE = Object.freeze({
  CLAIMED: 'claimed',
  SENDING: 'sending',
  REPORTING: 'reporting',
});

const TRANSITIONS = {
  [PHASE.CLAIMED]: [PHASE.SENDING, PHASE.REPORTING],
  [PHASE.SENDING]: [PHASE.REPORTING],
  [PHASE.REPORTING]: [],
};

export function canTransition(from, to) {
  return (TRANSITIONS[from] || []).includes(to);
}

export function startRun(jobId, now) {
  return { jobId, phase: PHASE.CLAIMED, startedAt: now, updatedAt: now };
}

export function advanceRun(run, to, now) {
  if (!run || !canTransition(run.phase, to)) {
    throw new Error(`Illegal run transition ${run ? run.phase : 'none'} -> ${to}`);
  }
  return { ...run, phase: to, updatedAt: now };
}

// ---- local duplicate protection (supplementary; backend idempotency is authoritative) ----

export const MAX_COMPLETED = 500;

export function getCompleted(map, id) {
  return map && Object.hasOwn(map, id) ? map[id] : null;
}

export function recordCompleted(map, id, status, now, max = MAX_COMPLETED) {
  const next = { ...(map || {}), [id]: { status, at: now } };
  const ids = Object.keys(next);
  if (ids.length > max) {
    ids.sort((a, b) => next[a].at - next[b].at).slice(0, ids.length - max).forEach((k) => delete next[k]);
  }
  return next;
}

// ---- crash / interruption recovery ----

// Content-script stages, in order. Once 'send_clicking' is recorded the message may have been sent.
export const STAGE = Object.freeze({
  STARTED: 'started',
  CHECKING_READY: 'checking_ready',
  OPENING_COMPOSE: 'opening_compose',
  FILLING: 'filling',
  VALIDATING: 'validating',
  SEND_CLICKING: 'send_clicking',
  VERIFYING: 'verifying',
  DONE: 'done',
});
const PRE_SEND = new Set([STAGE.STARTED, STAGE.CHECKING_READY, STAGE.OPENING_COMPOSE, STAGE.FILLING, STAGE.VALIDATING]);

/**
 * Decide the safe final status for a run whose result never reached the service worker.
 * Only a recorded pre-send stage proves nothing was sent; anything else needs human review.
 * @returns {{status: string, message: string}}
 */
export function classifyInterrupted(inflight, jobId) {
  if (!inflight || inflight.jobId !== jobId) {
    return { status: STATUS.NEEDS_REVIEW, message: 'Processing was interrupted and no progress record exists; the email may or may not have been sent. Check Sent Items before retrying.' };
  }
  if (inflight.stage === STAGE.DONE && inflight.result) {
    return resultToReport(inflight.result);
  }
  if (PRE_SEND.has(inflight.stage)) {
    return { status: STATUS.FAILED, message: `Interrupted before sending (stage: ${inflight.stage}); the email was not sent. A draft may be open in Outlook.` };
  }
  return { status: STATUS.NEEDS_REVIEW, message: `Interrupted at stage "${inflight.stage}" after Send was attempted; the email may have been sent. Check Sent Items before retrying.` };
}

/** Map a content-script result {outcome, message, evidence} to a report. Only verified sends are "sent". */
export function resultToReport(result) {
  if (!result || typeof result !== 'object') {
    return { status: STATUS.NEEDS_REVIEW, message: 'Unexpected result from Outlook page; the email may have been sent.' };
  }
  const msg = result.message || '';
  switch (result.outcome) {
    case 'sent':
      return { status: STATUS.SENT, message: `Verified: ${result.evidence || msg || 'Outlook accepted the message'}` };
    case 'failed':
      // 'failed' is only legitimate when Send was never clicked.
      if (result.sendClicked) return { status: STATUS.NEEDS_REVIEW, message: msg || 'Failure after Send was clicked; outcome unknown.' };
      return { status: STATUS.FAILED, message: msg || 'Failed before sending.' };
    default:
      return { status: STATUS.NEEDS_REVIEW, message: msg || 'Send outcome could not be verified.' };
  }
}
