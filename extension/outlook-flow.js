// Compose -> validate -> send -> verify flow for one job. Classic content script (globalThis.OEA).
// Safety rules:
//  * Send is clicked at most once per job; progress is persisted BEFORE the click.
//  * Any failure after the click is reported as 'uncertain', never as a plain failure.
//  * A job is 'sent' only if Outlook gave positive evidence (compose closed / sent toast).
(function (root) {
  'use strict';
  const OEA = (root.OEA = root.OEA || {});
  const { OEAError } = OEA;

  const DEFAULT_TIMEOUTS = { ready: 15000, compose: 15000, recipient: 10000, verify: 20000 };
  const EMAIL_RE = /^[^\s@,;<>()"]+@[^\s@,;<>()"]+\.[^\s@,;<>()"]{2,}$/;

  function validateJobFields(job) {
    const errors = [];
    if (!job || typeof job !== 'object') return ['Job is not an object.'];
    if (typeof job.id !== 'string' || !job.id) errors.push('Missing job id.');
    if (typeof job.to !== 'string' || !EMAIL_RE.test(job.to.trim())) errors.push('Recipient must be a single valid email address.');
    const cc = job.cc === undefined ? [] : job.cc;
    if (!Array.isArray(cc) || cc.length > 20 || cc.some((a) => typeof a !== 'string' || !EMAIL_RE.test(a.trim()))) errors.push('CC must be a list of valid email addresses.');
    else {
      const all = [job.to, ...cc].filter((a) => typeof a === 'string').map((a) => a.trim().toLowerCase());
      if (new Set(all).size !== all.length) errors.push('Recipients must be unique.');
    }
    if (typeof job.subject !== 'string' || !job.subject.trim() || /[\r\n]/.test(job.subject)) errors.push('Subject must be a non-empty single line.');
    if (typeof job.body !== 'string' || !job.body.trim()) errors.push('Missing body.');
    return errors;
  }

  async function enterRecipient(doc, compose, field, email, timeoutMs) {
    field.focus();
    if (field.tagName === 'INPUT') OEA.setInputValue(field, email);
    else OEA.setEditableText(field, email);
    // Attempt 1: Enter. Attempt 2: ';' terminator, which Outlook also treats as "resolve".
    const attempts = [() => OEA.pressKey(field, 'Enter', 13), () => { OEA.setEditableText(field, `${email};`); OEA.pressKey(field, 'Enter', 13); }];
    for (const attempt of attempts) {
      attempt();
      try {
        await OEA.waitFor(
          () => {
            const err = OEA.findRecipientError(doc);
            if (err) throw new OEAError('RECIPIENT_REJECTED', `Outlook did not accept the recipient: ${err}`);
            return OEA.findRecipientPill(compose, email, field);
          },
          { timeoutMs: Math.ceil(timeoutMs / 2), description: 'the recipient to resolve', doc },
        );
        return;
      } catch (e) {
        if (e.code !== 'TIMEOUT') throw e;
      }
    }
    throw new OEAError('RECIPIENT_FAILED', 'Outlook did not resolve the recipient into the To field.');
  }

  function validateComposed(doc, job, compose) {
    const problems = [];
    const cc = job.cc || [];
    if (OEA.countCompose(doc) !== 1) problems.push('Expected exactly one open compose window.');
    const live = OEA.findCompose(doc);
    if (!live || live.root !== compose.root) problems.push('The compose window changed during filling.');
    if (!OEA.findRecipientPill(compose, job.to.trim())) problems.push('Recipient is not present in the To field.');
    if (cc.length) {
      const ccField = OEA.findCc(compose.root);
      if (!ccField) problems.push('The Cc field is not available.');
      else for (const a of cc) if (!OEA.findRecipientPill(compose, a.trim(), ccField)) problems.push('A CC recipient is not present in the Cc field.');
    }
    if ((compose.subject.value || '') !== job.subject) problems.push('Subject does not match.');
    if (!OEA.squash(OEA.extractText(compose.body)).includes(OEA.squash(job.body))) problems.push('Body does not match.');
    if (compose.send.disabled || compose.send.getAttribute('aria-disabled') === 'true') problems.push('The Send button is disabled.');
    return problems;
  }

  /**
   * @param job {id,to,subject,body}
   * @param ctx {doc, loc, guard:{processed:Set,current:string|null}, markStage(jobId,stage):Promise, timeouts?}
   * @returns {{outcome:'sent'|'failed'|'uncertain', sendClicked:boolean, code?:string, message:string, evidence?:string}}
   */
  async function processJob(job, ctx) {
    const doc = ctx.doc;
    const t = { ...DEFAULT_TIMEOUTS, ...(ctx.timeouts || {}) };
    const guard = ctx.guard;
    let sendClicked = false;

    const bad = validateJobFields(job);
    if (bad.length) return { outcome: 'failed', sendClicked: false, code: 'INVALID_JOB', message: `Invalid job: ${bad.join(' ')}` };
    if (guard.processed.has(job.id) || guard.current === job.id) {
      return { outcome: 'uncertain', sendClicked: false, code: 'DUPLICATE_JOB', message: 'This job was already handled by this Outlook tab; refusing to send again.' };
    }
    if (guard.current) return { outcome: 'failed', sendClicked: false, code: 'BUSY', message: 'Another job is being processed in this tab.' };

    guard.current = job.id;
    try {
      await ctx.markStage(job.id, 'started');

      await ctx.markStage(job.id, 'checking_ready');
      await OEA.waitFor(() => {
        const r = OEA.detectReadiness(doc, ctx.loc);
        if (r.state === 'signed_out' || r.state === 'unexpected_page') throw new OEAError(r.state.toUpperCase(), r.message);
        return r.state === 'ready';
      }, { timeoutMs: t.ready, description: 'Outlook to finish loading', doc });

      await ctx.markStage(job.id, 'opening_compose');
      if (OEA.countCompose(doc) > 0) {
        throw new OEAError('COMPOSE_ALREADY_OPEN', 'A compose window is already open. Send or discard that draft, then retry.');
      }
      const newMail = OEA.findNewMail(doc);
      if (!newMail) throw new OEAError('NEW_MAIL_MISSING', 'Could not find the "New mail" button.');
      newMail.click();
      const compose = await OEA.waitFor(() => OEA.findCompose(doc), { timeoutMs: t.compose, description: 'the compose window', doc });

      await ctx.markStage(job.id, 'filling');
      await enterRecipient(doc, compose, compose.to, job.to.trim(), t.recipient);
      const ccList = (job.cc || []).map((a) => a.trim());
      if (ccList.length) {
        let ccField = OEA.findCc(compose.root);
        if (!ccField) {
          const toggle = OEA.findCcToggle(compose.root);
          if (!toggle) throw new OEAError('CC_FIELD_MISSING', 'Could not find the Cc control in the compose window.');
          toggle.click();
          ccField = await OEA.waitFor(() => OEA.findCc(compose.root), { timeoutMs: t.compose, description: 'the Cc field', doc });
        }
        for (const addr of ccList) await enterRecipient(doc, compose, ccField, addr, t.recipient);
      }
      OEA.setInputValue(compose.subject, job.subject);
      if (compose.subject.value !== job.subject) throw new OEAError('SUBJECT_FAILED', 'Could not set the subject.');
      OEA.setEditableText(compose.body, job.body);
      if (!OEA.squash(OEA.extractText(compose.body)).includes(OEA.squash(job.body))) {
        throw new OEAError('BODY_FAILED', 'Could not populate the message body.');
      }

      await ctx.markStage(job.id, 'validating');
      const problems = validateComposed(doc, job, compose);
      if (problems.length) throw new OEAError('VALIDATION_FAILED', `Message validation failed: ${problems.join(' ')}`);

      // Point of no return. Persist first; if that fails we abort without sending.
      await ctx.markStage(job.id, 'send_clicking');
      guard.processed.add(job.id);
      sendClicked = true;
      compose.send.click(); // exactly once

      await ctx.markStage(job.id, 'verifying');
      const verdict = await OEA.waitFor(() => {
        if (OEA.countCompose(doc) === 0) return { kind: 'closed' };
        const err = OEA.findSendError(doc);
        if (err) return { kind: 'error', text: err };
        if (OEA.findSentToast(doc)) return { kind: 'toast' };
        return null;
      }, { timeoutMs: t.verify, description: 'Outlook to confirm the send', doc });

      if (verdict.kind === 'error') {
        return { outcome: 'uncertain', sendClicked, code: 'SEND_ERROR_SHOWN', message: `Outlook reported a problem after Send: "${verdict.text}". The email was not resent.` };
      }
      return {
        outcome: 'sent', sendClicked, code: 'OK',
        evidence: verdict.kind === 'closed' ? 'compose window closed after Send' : 'Outlook showed a "message sent" notification',
        message: 'Outlook accepted the message.',
      };
    } catch (e) {
      const code = e instanceof OEAError ? e.code : 'UNEXPECTED';
      const message = e instanceof OEAError ? e.message : `Unexpected error: ${e && e.message}`;
      if (sendClicked) {
        return { outcome: 'uncertain', sendClicked, code: code === 'TIMEOUT' ? 'NO_CONFIRMATION' : code,
          message: code === 'TIMEOUT' ? 'Send was clicked but Outlook gave no confirmation in time; the message may or may not have been sent. It was not resent.' : message };
      }
      return { outcome: 'failed', sendClicked, code, message };
    } finally {
      guard.current = null;
    }
  }

  Object.assign(OEA, { processJob, validateJobFields, DEFAULT_TIMEOUTS });
})(typeof globalThis !== 'undefined' ? globalThis : window);
