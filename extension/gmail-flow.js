// Gmail provider: compose -> validate -> send -> verify for one job. Classic content script (globalThis.OEA).
// Same safety rules as the other providers: Send is clicked at most once, progress is persisted before the click,
// anything uncertain after the click is reported 'uncertain' (never resent), and 'sent' needs positive evidence
// (compose closed / "Message sent" notice). Sent-folder confirmation is not implemented for Gmail.
(function (root) {
  'use strict';
  const OEA = (root.OEA = root.OEA || {});
  const { OEAError } = OEA;
  const G = OEA.gmail;

  async function enterRecipient(doc, compose, name, email, timeoutMs) {
    const attempts = [
      (f) => { OEA.setInputValue(f, email); OEA.pressKey(f, 'Enter', 13); },
      (f) => { OEA.setInputValue(f, `${email},`); },
    ];
    for (const attempt of attempts) {
      const f = G.findField(compose, name);
      if (!f) throw new OEAError('RECIPIENT_FIELD_MISSING', `Could not find the ${name} field.`);
      f.focus();
      attempt(f);
      try {
        await OEA.waitFor(() => {
          const err = G.findRecipientError(doc);
          if (err) throw new OEAError('RECIPIENT_REJECTED', `Gmail did not accept the recipient: ${err}`);
          const live = G.findField(compose, name); // re-resolved: the page may re-render the row
          return live ? G.findRecipientPill(compose, email, live) : null;
        }, { timeoutMs: Math.ceil(timeoutMs / 2), description: 'the recipient to resolve', doc });
        return;
      } catch (e) {
        if (e.code !== 'TIMEOUT') throw e;
      }
    }
    throw new OEAError('RECIPIENT_FAILED', `Gmail did not turn the address into a recipient in the ${name} field.`);
  }

  /** Gmail's own state must hold the subject (its hidden copy updates); try a plain set, an editing command, then key events. */
  async function enterSubject(doc, compose, subject, timeoutMs) {
    const methods = [(i) => { OEA.setInputValue(i, subject); }, (i) => OEA.typeIntoInput(i, subject), (i) => OEA.typeIntoInputByKeys(i, subject)];
    for (const method of methods) {
      const live = G.findCompose(doc);
      const input = live && live.subject;
      if (!input) throw new OEAError('SUBJECT_FAILED', 'The Subject field disappeared.');
      method(input);
      try {
        await OEA.waitFor(() => input.value === subject && G.subjectRegistered(compose, subject), { timeoutMs: Math.min(timeoutMs, 2500), description: 'Gmail to register the subject', doc });
        return;
      } catch (e) {
        if (e.code !== 'TIMEOUT') throw e;
      }
    }
    throw new OEAError('SUBJECT_FAILED', 'Gmail did not register the subject (its own copy of the subject stayed different). Nothing was sent.');
  }

  function validateComposed(doc, job, compose, confirmed) {
    const problems = [];
    const cc = (job.cc || []).map((a) => a.trim());
    const to = job.to.trim();
    if (G.countCompose(doc) !== 1) problems.push('Expected exactly one open compose window.');
    const live = G.findCompose(doc);
    if (!live || live.root !== compose.root) problems.push('The compose window changed during filling.');
    for (const a of [to, ...cc]) if (!confirmed.has(a.toLowerCase())) problems.push('A recipient was not confirmed as a chip.');
    const toField = G.findField(compose, 'To');
    if (!toField) problems.push('The To field is not available.');
    else {
      const s = G.chipState(compose, toField, [to]);
      if (s.count !== 1 || !s.found[0]) problems.push('The To field does not contain exactly the intended recipient.');
    }
    if (cc.length) {
      const ccField = G.findField(compose, 'Cc');
      if (!ccField) problems.push('The Cc field is not available.');
      else {
        const s = G.chipState(compose, ccField, cc);
        if (s.count !== cc.length || s.found.some((f) => !f)) problems.push('The Cc field does not contain exactly the intended recipients.');
      }
    } else {
      const ccField = G.findField(compose, 'Cc');
      if (ccField && G.chipState(compose, ccField, []).count) problems.push('The Cc field has recipients that should not be there.');
    }
    if ((compose.subject.value || '') !== job.subject) problems.push('Subject does not match.');
    if (!G.subjectRegistered(compose, job.subject)) problems.push('Gmail did not register the subject.');
    if (!OEA.squash(OEA.extractText(compose.body)).includes(OEA.squash(job.body))) problems.push('Body does not match.');
    else if (!OEA.bodyMatchesLines(compose.body, job.body)) problems.push('Body line breaks do not match.');
    if (compose.send.getAttribute('aria-disabled') === 'true') problems.push('The Send button is disabled.');
    return problems;
  }

  /**
   * @param job {id,to,cc?,subject,body}
   * @param ctx {doc, loc, guard, markStage(jobId,stage):Promise, timeouts?}
   */
  async function processJob(job, ctx) {
    const doc = ctx.doc;
    const t = { ...OEA.DEFAULT_TIMEOUTS, ...(ctx.timeouts || {}) };
    const guard = ctx.guard;
    let sendClicked = false;

    const bad = OEA.validateJobFields(job);
    if (bad.length) return { outcome: 'failed', sendClicked: false, code: 'INVALID_JOB', message: `Invalid job: ${bad.join(' ')}` };
    if (guard.processed.has(job.id) || guard.current === job.id) {
      return { outcome: 'uncertain', sendClicked: false, code: 'DUPLICATE_JOB', message: 'This job was already handled by this Gmail tab; refusing to send again.' };
    }
    if (guard.current) return { outcome: 'failed', sendClicked: false, code: 'BUSY', message: 'Another job is being processed in this tab.' };

    guard.current = job.id;
    try {
      await ctx.markStage(job.id, 'started');

      await ctx.markStage(job.id, 'checking_ready');
      await OEA.waitFor(() => {
        const r = G.detectReadiness(doc, ctx.loc);
        if (r.state === 'unexpected_page') throw new OEAError('UNEXPECTED_PAGE', r.message);
        return r.state === 'ready';
      }, { timeoutMs: t.ready, description: 'Gmail to finish loading', doc });

      await ctx.markStage(job.id, 'opening_compose');
      if (G.countCompose(doc) > 0) throw new OEAError('COMPOSE_ALREADY_OPEN', 'A compose window is already open. Send or discard that draft, then retry.');
      const newMail = G.findNewMail(doc);
      if (!newMail) throw new OEAError('NEW_MAIL_MISSING', 'Could not find the "Compose" button.');
      G.pressButton(newMail);
      const compose = await OEA.waitFor(() => {
        const c = G.findCompose(doc);
        return c && G.findField(c, 'To') ? c : null;
      }, { timeoutMs: t.compose, description: 'the compose window', doc });

      await ctx.markStage(job.id, 'filling');
      const confirmed = new Set();
      const to = job.to.trim();
      await enterRecipient(doc, compose, 'To', to, t.recipient);
      confirmed.add(to.toLowerCase());
      const ccList = (job.cc || []).map((a) => a.trim());
      if (ccList.length) {
        if (!G.findField(compose, 'Cc')) {
          const toggle = G.findCcToggle(compose);
          if (!toggle) throw new OEAError('CC_FIELD_MISSING', 'Could not find the "Cc" link in the compose window.');
          G.pressButton(toggle);
          await OEA.waitFor(() => G.findField(compose, 'Cc'), { timeoutMs: t.compose, description: 'the Cc field', doc });
        }
        for (const addr of ccList) {
          await enterRecipient(doc, compose, 'Cc', addr, t.recipient);
          confirmed.add(addr.toLowerCase());
        }
      }
      const refresh = () => { const live = G.findCompose(doc); if (live && live.root === compose.root) Object.assign(compose, live); };
      refresh();
      await enterSubject(doc, compose, job.subject, t.recipient);
      refresh();
      OEA.setEditableText(compose.body, job.body);
      if (!OEA.squash(OEA.extractText(compose.body)).includes(OEA.squash(job.body))) throw new OEAError('BODY_FAILED', 'Could not populate the message body.');
      if (!OEA.bodyMatchesLines(compose.body, job.body)) {
        throw new OEAError('BODY_FORMAT_FAILED', 'Gmail\'s editor did not keep the message\'s line breaks (they came out merged into fewer lines). Nothing was sent.');
      }

      refresh();
      await ctx.markStage(job.id, 'validating');
      const problems = validateComposed(doc, job, compose, confirmed);
      if (problems.length) throw new OEAError('VALIDATION_FAILED', `Message validation failed: ${problems.join(' ')}`);

      // Point of no return. Persist first; if that fails we abort without sending.
      await ctx.markStage(job.id, 'send_clicking');
      guard.processed.add(job.id);
      sendClicked = true;
      G.pressButton(compose.send); // exactly once

      await ctx.markStage(job.id, 'verifying');
      const verdict = await OEA.waitFor(() => {
        if (G.countCompose(doc) === 0) return { kind: 'closed' };
        const err = G.findSendError(doc);
        if (err) return { kind: 'error', text: err };
        if (G.findSentToast(doc)) return { kind: 'toast' };
        return null;
      }, { timeoutMs: t.verify, description: 'Gmail to confirm the send', doc });

      if (verdict.kind === 'error') {
        return { outcome: 'uncertain', sendClicked, code: 'SEND_ERROR_SHOWN', message: `Gmail reported a problem after Send: "${verdict.text}". The email was not resent.` };
      }
      const evidence = verdict.kind === 'closed' ? 'compose window closed after Send' : 'Gmail showed "Message sent"';
      return { outcome: 'sent', sendClicked, code: 'OK', evidence, message: 'Gmail accepted the message.' };
    } catch (e) {
      const code = e instanceof OEAError ? e.code : 'UNEXPECTED';
      const message = e instanceof OEAError ? e.message : `Unexpected error: ${e && e.message}`;
      if (sendClicked) {
        return { outcome: 'uncertain', sendClicked, code: code === 'TIMEOUT' ? 'NO_CONFIRMATION' : code,
          message: code === 'TIMEOUT' ? 'Send was clicked but Gmail gave no confirmation in time; the message may or may not have been sent. It was not resent.' : message };
      }
      return { outcome: 'failed', sendClicked, code, message };
    } finally {
      guard.current = null;
    }
  }

  OEA.activeProvider = { name: 'gmail', detectReadiness: G.detectReadiness, processJob };
})(typeof globalThis !== 'undefined' ? globalThis : window);
