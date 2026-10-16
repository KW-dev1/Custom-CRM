import test from 'node:test';
import assert from 'node:assert/strict';
import {
  PHASE, STAGE, canTransition, startRun, advanceRun, recordCompleted, getCompleted,
  classifyInterrupted, resultToReport, MAX_COMPLETED,
} from '../state.js';
import { STATUS } from '../config.js';

test('state transitions: legal path and illegal jumps', () => {
  let r = startRun('j', 1);
  r = advanceRun(r, PHASE.SENDING, 2);
  r = advanceRun(r, PHASE.REPORTING, 3);
  assert.equal(r.phase, PHASE.REPORTING);
  assert.equal(canTransition(PHASE.REPORTING, PHASE.SENDING), false);
  assert.throws(() => advanceRun(startRun('j', 1), 'bogus', 2), /Illegal/);
  assert.throws(() => advanceRun(null, PHASE.SENDING, 2), /Illegal/);
});
test('completed map: record, lookup, prune oldest', () => {
  let m = {};
  m = recordCompleted(m, 'a', STATUS.SENT, 1);
  assert.equal(getCompleted(m, 'a').status, 'sent');
  assert.equal(getCompleted(m, 'zzz'), null);
  assert.equal(getCompleted(m, 'constructor'), null);
  for (let i = 0; i < MAX_COMPLETED + 5; i++) m = recordCompleted(m, `k${i}`, STATUS.SENT, 10 + i);
  assert.equal(Object.keys(m).length, MAX_COMPLETED);
  assert.equal(getCompleted(m, 'a'), null);
});
test('interrupted: no record or other job => needs_review', () => {
  assert.equal(classifyInterrupted(null, 'j').status, STATUS.NEEDS_REVIEW);
  assert.equal(classifyInterrupted({ jobId: 'other', stage: STAGE.FILLING }, 'j').status, STATUS.NEEDS_REVIEW);
});
test('interrupted: before send click => failed (safe); at/after click => needs_review', () => {
  for (const s of [STAGE.STARTED, STAGE.FILLING, STAGE.VALIDATING]) {
    assert.equal(classifyInterrupted({ jobId: 'j', stage: s }, 'j').status, STATUS.FAILED);
  }
  for (const s of [STAGE.SEND_CLICKING, STAGE.VERIFYING]) {
    assert.equal(classifyInterrupted({ jobId: 'j', stage: s }, 'j').status, STATUS.NEEDS_REVIEW);
  }
});
test('interrupted: persisted final result is honoured', () => {
  const r = classifyInterrupted({ jobId: 'j', stage: STAGE.DONE, result: { outcome: 'sent', evidence: 'compose window closed' } }, 'j');
  assert.equal(r.status, STATUS.SENT);
});
test('resultToReport: only verified outcome=sent is sent', () => {
  assert.equal(resultToReport({ outcome: 'sent', evidence: 'x' }).status, STATUS.SENT);
  assert.equal(resultToReport({ outcome: 'uncertain' }).status, STATUS.NEEDS_REVIEW);
  assert.equal(resultToReport({ outcome: 'failed', sendClicked: false }).status, STATUS.FAILED);
  assert.equal(resultToReport({ outcome: 'failed', sendClicked: true }).status, STATUS.NEEDS_REVIEW);
  assert.equal(resultToReport(undefined).status, STATUS.NEEDS_REVIEW);
  assert.equal(resultToReport({ outcome: 'weird' }).status, STATUS.NEEDS_REVIEW);
});
