// Yahoo Mail provider: compose -> validate -> send -> verify for one job. Classic content script (globalThis.OEA).
// Same safety rules as the Outlook flow:
//  * Send is clicked at most once per job; progress is persisted BEFORE the click.
//  * Any failure after the click is reported as 'uncertain', never as a plain failure.
//  * A job is 'sent' only if Yahoo gave positive evidence (compose closed / "Your message has been sent." toast).
// Sent-folder confirmation is not implemented for Yahoo (its message-list markup has not been captured), so the
// "Confirm in Sent Items" option is ignored here.
(function (root) {
  'use strict';
  const OEA = (root.OEA = root.OEA || {});
  const { OEAError } = OEA;
  const Y = OEA.yahoo;

  async function enterRecipient(doc, compose, name, email, timeoutMs) {
    const field = Y.findField(compose, name);
    if (!field) throw new OEAError('RECIPIENT_FIELD_MISSING', `Could not find the ${name} field.`);
    field.focus();
    // Attempt 1: type the address and press Enter. Attempt 2: a ';' terminator, which Yahoo also treats as "done".
    const attempts = [
      (f) => { OEA.setInputValue(f, email); OEA.pressKey(f, 'Enter', 13); },
      (f) => { OEA.setInputValue(f, `${email};`); },
    ];
    for (const attempt of attempts) {
      const f = Y.findField(compose, name);
      if (!f) throw new OEAError('RECIPIENT_FIELD_MISSING', `The ${name} field disappeared while entering the recipient.`);
      f.focus();
      attempt(f);
      try {
        await OEA.waitFor(() => {
          const err = Y.findRecipientError(doc);
          if (err) throw new OEAError('RECIPIENT_REJECTED', `Yahoo did not accept the recipient: ${err}`);
          const live = Y.findField(compose, name); // re-resolved: the page may re-render the row
          return live ? Y.findRecipientPill(compose, email, live) : null;
        }, { timeoutMs: Math.ceil(timeoutMs / 2), description: 'the recipient to resolve', doc });
        return;
      } catch (e) {
        if (e.code !== 'TIMEOUT') throw e;
      }
    }
    throw new OEAError('RECIPIENT_FAILED', `Yahoo did not turn the address into a recipient in the ${name} field.`);
  }

  /** Yahoo ignores typing while its page is not focused (the compose title never picks the subject up). */
  async function waitForPageFocus(doc, timeoutMs) {
    if (typeof doc.hasFocus !== 'function') return;
    try {
      await OEA.waitFor(() => doc.hasFocus(), { timeoutMs, description: 'the Yahoo tab to have keyboard focus', doc });
    } catch (e) {
      if (e.code !== 'TIMEOUT') throw e;
      throw new OEAError('PAGE_NOT_FOCUSED', 'The Yahoo Mail window is not in the foreground. Yahoo only accepts typing while its window has focus, so keep that window in front and the extension popup closed while it sends. Nothing was sent.');
    }
  }

  /**
   * Yahoo only counts the subject once its own state has it, which shows up as the compose title changing. Try an
   * editing command first, then real per-character key events, and give each a moment to register.
   */
  async function enterSubject(doc, compose, subject, timeoutMs) {
    const methods = [(i) => Y.typeIntoInput(i, subject), (i) => Y.typeIntoInputByKeys(i, subject)];
    for (const method of methods) {
      const input = Y.findCompose(doc) && Y.findCompose(doc).subject;
      if (!input) throw new OEAError('SUBJECT_FAILED', 'The Subject field disappeared.');
      method(input);
      try {
        await OEA.waitFor(() => input.value === subject && Y.subjectRegistered(compose, subject), { timeoutMs: Math.min(timeoutMs, 3000), description: 'Yahoo to register the subject', doc });
        return;
      } catch (e) {
        if (e.code !== 'TIMEOUT') throw e;
      }
    }
    throw new OEAError('SUBJECT_FAILED', 'Yahoo did not register the subject (the compose title did not change). Nothing was sent.');
  }

  function validateComposed(doc, job, compose, confirmed) {
    const problems = [];
    const cc = (job.cc || []).map((a) => a.trim().toLowerCase());
    const to = job.to.trim().toLowerCase();
    if (Y.countCompose(doc) !== 1) problems.push('Expected exactly one open compose window.');
    const live = Y.findCompose(doc);
    if (!live || live.root !== compose.root) problems.push('The compose window changed during filling.');
    for (const a of [to, ...cc]) if (!confirmed.has(a)) problems.push('A recipient was not confirmed as a chip.');
    // While the rows are expanded the chips must be exactly the intended ones (an extra suggestion chip means a wrong pick).
    const toField = Y.findField(compose, 'To');
    const ccField = Y.findField(compose, 'Cc');
    if (toField && Y.pillAddresses(compose, toField).join(',') !== to) problems.push('The To field does not contain exactly the intended recipient.');
    if (ccField && [...Y.pillAddresses(compose, ccField)].sort().join(',') !== [...cc].sort().join(',')) problems.push('The Cc field does not contain exactly the intended recipients.');
    // Once Yahoo collapses the rows it shows "<first> +N more"; check it agrees with what we entered.
    if (!toField) {
      const sum = Y.readRecipientSummary(compose);
      if (!sum) problems.push('Could not read the recipients after Yahoo collapsed them.');
      else if (sum.first !== to) problems.push('The first recipient shown does not match.');
      else if (sum.more && sum.more !== cc.length) problems.push('The number of recipients shown does not match.');
    }
    if ((compose.subject.value || '') !== job.subject) problems.push('Subject does not match.');
    // The header title mirrors the subject once Yahoo's own state has it; a field that only LOOKS filled leaves "New message".
    if (!Y.subjectRegistered(compose, job.subject)) problems.push('Yahoo did not register the subject (the compose title did not change).');
    if (!OEA.squash(OEA.extractText(compose.body)).includes(OEA.squash(job.body))) problems.push('Body does not match.');
    else if (!OEA.bodyMatchesLines(compose.body, job.body)) problems.push('Body line breaks do not match.');
    if (compose.send.disabled || compose.send.getAttribute('aria-disabled') === 'true') problems.push('The Send button is disabled.');
    return problems;
  }

  /**
   * @param job {id,to,cc?,subject,body}
   * @param ctx {doc, loc, guard, markStage(jobId,stage):Promise, timeouts?}
   * @returns {{outcome:'sent'|'failed'|'uncertain', sendClicked:boolean, code?:string, message:string, evidence?:string}}
   */
  async function processJob(job, ctx) {
    const doc = ctx.doc;
    const t = { focus: 10000, ...OEA.DEFAULT_TIMEOUTS, ...(ctx.timeouts || {}) };
    const guard = ctx.guard;
    let sendClicked = false;

    const bad = OEA.validateJobFields(job);
    if (bad.length) return { outcome: 'failed', sendClicked: false, code: 'INVALID_JOB', message: `Invalid job: ${bad.join(' ')}` };
    if (guard.processed.has(job.id) || guard.current === job.id) {
      return { outcome: 'uncertain', sendClicked: false, code: 'DUPLICATE_JOB', message: 'This job was already handled by this Yahoo tab; refusing to send again.' };
    }
    if (guard.current) return { outcome: 'failed', sendClicked: false, code: 'BUSY', message: 'Another job is being processed in this tab.' };

    guard.current = job.id;
    try {
      await ctx.markStage(job.id, 'started');

      await ctx.markStage(job.id, 'checking_ready');
      await OEA.waitFor(() => {
        const r = Y.detectReadiness(doc, ctx.loc);
        if (r.state === 'unexpected_page') throw new OEAError('UNEXPECTED_PAGE', r.message);
        return r.state === 'ready';
      }, { timeoutMs: t.ready, description: 'Yahoo Mail to finish loading', doc });

      await ctx.markStage(job.id, 'opening_compose');
      if (Y.countCompose(doc) > 0) throw new OEAError('COMPOSE_ALREADY_OPEN', 'A compose window is already open. Send or discard that draft, then retry.');
      const newMail = Y.findNewMail(doc);
      if (!newMail) throw new OEAError('NEW_MAIL_MISSING', 'Could not find the "New message" button.');
      newMail.click();
      const compose = await OEA.waitFor(() => {
        const c = Y.findCompose(doc);
        return c && Y.findField(c, 'To') ? c : null;
      }, { timeoutMs: t.compose, description: 'the compose window', doc });

      await ctx.markStage(job.id, 'filling');
      const confirmed = new Set();
      const to = job.to.trim();
      await enterRecipient(doc, compose, 'To', to, t.recipient);
      confirmed.add(to.toLowerCase());
      const ccList = (job.cc || []).map((a) => a.trim());
      if (ccList.length) {
        if (!Y.findField(compose, 'Cc')) {
          const toggle = Y.findCcToggle(compose);
          if (!toggle) throw new OEAError('CC_FIELD_MISSING', 'Could not find the "Cc Bcc" control in the compose window.');
          toggle.click();
          await OEA.waitFor(() => Y.findField(compose, 'Cc'), { timeoutMs: t.compose, description: 'the Cc field', doc });
        }
        for (const addr of ccList) {
          await enterRecipient(doc, compose, 'Cc', addr, t.recipient);
          confirmed.add(addr.toLowerCase());
        }
      }
      const refresh = () => { const live = Y.findCompose(doc); if (live && live.root === compose.root) Object.assign(compose, live); }; // the page may re-render fields
      await waitForPageFocus(doc, t.focus);
      refresh();
      await enterSubject(doc, compose, job.subject, t.recipient);
      refresh();
      OEA.setEditableText(compose.body, job.body);
      if (!OEA.squash(OEA.extractText(compose.body)).includes(OEA.squash(job.body))) throw new OEAError('BODY_FAILED', 'Could not populate the message body.');
      if (!OEA.bodyMatchesLines(compose.body, job.body)) {
        throw new OEAError('BODY_FORMAT_FAILED', 'Yahoo\'s editor did not keep the message\'s line breaks (they came out merged into fewer lines). Nothing was sent.');
      }

      refresh();
      await ctx.markStage(job.id, 'validating');
      const problems = validateComposed(doc, job, compose, confirmed);
      if (problems.length) throw new OEAError('VALIDATION_FAILED', `Message validation failed: ${problems.join(' ')}`);

      // Point of no return. Persist first; if that fails we abort without sending.
      await ctx.markStage(job.id, 'send_clicking');
      guard.processed.add(job.id);
      sendClicked = true;
      compose.send.click(); // exactly once

      await ctx.markStage(job.id, 'verifying');
      const verdict = await OEA.waitFor(() => {
        if (Y.countCompose(doc) === 0) return { kind: 'closed' };
        const err = Y.findSendError(doc);
        if (err) return { kind: 'error', text: err };
        if (Y.findSentToast(doc)) return { kind: 'toast' };
        return null;
      }, { timeoutMs: t.verify, description: 'Yahoo to confirm the send', doc });

      if (verdict.kind === 'error') {
        return { outcome: 'uncertain', sendClicked, code: 'SEND_ERROR_SHOWN', message: `Yahoo reported a problem after Send: "${verdict.text}". The email was not resent.` };
      }
      const evidence = verdict.kind === 'closed' ? 'compose window closed after Send' : 'Yahoo showed "Your message has been sent."';
      return { outcome: 'sent', sendClicked, code: 'OK', evidence, message: 'Yahoo accepted the message.' };
    } catch (e) {
      const code = e instanceof OEAError ? e.code : 'UNEXPECTED';
      const message = e instanceof OEAError ? e.message : `Unexpected error: ${e && e.message}`;
      if (sendClicked) {
        return { outcome: 'uncertain', sendClicked, code: code === 'TIMEOUT' ? 'NO_CONFIRMATION' : code,
          message: code === 'TIMEOUT' ? 'Send was clicked but Yahoo gave no confirmation in time; the message may or may not have been sent. It was not resent.' : message };
      }
      return { outcome: 'failed', sendClicked, code, message };
    } finally {
      guard.current = null;
    }
  }

  OEA.activeProvider = { name: 'yahoo', detectReadiness: Y.detectReadiness, processJob };
})(typeof globalThis !== 'undefined' ? globalThis : window);
