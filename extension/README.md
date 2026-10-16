# Outlook Email Automation (Chrome extension)

A Manifest V3 Chrome extension that takes email jobs from your backend API, composes each one in
**Outlook Web** (`https://outlook.live.com/mail/`) through the visible UI, clicks **Send**, verifies
the outcome, and reports the result back. Plain JavaScript/HTML/CSS, no build step.

It uses DOM interaction only: no Microsoft Graph, no SMTP, no third-party mail service. It does not
contain a backend, database, dashboard or queue; the backend is a separate component.

> **Status:** unit tests, a simulated-Outlook DOM test suite and a real-Chrome smoke test (against a
> *simulated* Outlook page and a mock backend) pass. **It has not been run against the real
> Outlook Web or a real backend, and no real email has been sent by this project.** Outlook's markup
> changes; use the [manual checklist](#manual-testing-checklist) before relying on it.

## Requirements

- Chrome (or Chromium-based) **120+**
- A signed-in Outlook Web tab at `https://outlook.live.com/mail/` (English UI)
- A backend implementing the [API contract](#backend-api-contract) (or adapt the paths in settings)
- Node 20+ only for running tests

## Install (unpacked)

1. `chrome://extensions` → enable **Developer mode**.
2. **Load unpacked** → select this `extension/` folder.
3. Pin the extension and open its popup.

## Configure

In the popup, under **Settings**:

| Field | Meaning |
|---|---|
| Backend API base URL | `https://…` (plain `http://` is accepted only for `localhost`/`127.0.0.1`). |
| API key / token | Stored in `chrome.storage.local`. Never logged, never sent anywhere except the backend. |
| Confirm each send in Sent Items | On by default. Stricter and adds ~10 s per email; see *How success is verified*. |
| Advanced → Auth header / scheme | Default `Authorization: Bearer <key>`. Clear the scheme to send the raw key (e.g. header `X-Api-Key`). |
| Advanced → claim/status/state paths & methods | **Placeholders**, replace with your real endpoints. |
| Advanced → Fetch every (min/max minutes) | The backend is asked for **one job at a random time between min and max minutes** (default 10 and 20) after the previous fetch. Not a way around Outlook's own limits. |
| Advanced → Status values (JSON) | Maps internal statuses to your backend's vocabulary. |

On **Save**, Chrome asks for permission to contact *only* the backend origin you entered
(`optional_host_permissions`). The extension has no broad host access.

## Run

1. Open `https://outlook.live.com/mail/` in a tab and sign in yourself. **The extension never opens
   tabs or touches authentication**; with no Outlook tab it shows *Waiting for Outlook*.
2. Tick **Automation enabled**, then press **Start**. **Stop** pauses. **Start fetches immediately**, then
   once every random 10–20 minutes. A one-minute watchdog alarm only recovers/flushes state and checks the
   clock; it never fetches early, so a suspended or restarted service worker is harmless.
3. Reload any Outlook tab that was open *before* installing the extension.

Popup statuses: *Not configured, Waiting for Outlook, Ready, Processing email, Email sent, Failed,
Needs review, Paused*. **Email sent** is shown only after Outlook gave positive evidence.

Automation **pauses itself** after a *needs review* result, an authentication failure, or 3 failures
in a row.

## Backend API contract

**Defaults**: Everything below is the default the extension
assumes (and which the bundled backend implements); all paths, methods and the auth header are
configurable. Job fields are fixed: `id`, `to`, `cc`, `subject`, `body` (+ optional `metadata`). Adapt `parseJobResponse` in `api.js` if your
shape differs.

**Claim next job**: `POST {base}/api/email-jobs/claim` (body `{}`)

- `200 {"job": {"id":"r5-1a2b3c4d5e6f","to":"a@b.com","cc":["c@d.com"],"subject":"…","body":"…"}}` or a bare job object.
  `cc` is optional (default `[]`): the first address is **To**, the others are **Cc**. The bundled
  [`backend/`](../backend/README.md) implements exactly this contract.
- `204` / `{"job": null}` → nothing to do
- The backend should *reserve* the job for this caller (so two clients can't get it). The claim is
  **never retried** by the extension.

**Report status**: `POST {base}/api/email-jobs/{id}/status`

```json
{ "jobId": "job_123", "status": "sent", "message": "Verified: compose window closed after Send", "timestamp": "2026-01-01T12:00:00.000Z" }
```

`status` ∈ `processing | sent | failed | needs_review` (remappable). This call must be **idempotent**
per job id (it is retried with backoff on network/5xx/429). `failed` is only ever sent for jobs
**provably not sent**; `needs_review` means the email may have been sent, so **do not auto-requeue it**.

**Job state (optional)**: `GET {base}/api/email-jobs/{id}` → `{"status":"…"}`. If a path is set,
the extension skips a claimed job whose backend state is already `sent`/`needs_review`.

Auth: header from settings, default `Authorization: Bearer <key>`. HTTP 401/403 → automation pauses.

## How duplicate sends are prevented

1. One job at a time per tick (in-memory flag + persisted `run` record).
2. In the Outlook tab: a per-job guard; `Send` is clicked at most once; the same id is refused afterwards.
3. Progress is written to `chrome.storage` **before** the Send click; if that write fails, nothing is sent.
4. Local completed-job map (last 500). A re-delivered job whose earlier outcome was `sent` or `needs_review`
   is **not** re-sent; its prior status is re-reported. A job that earlier **failed before sending** is retried.
   A job id is the row number plus a fingerprint of its content, so a cleared row gets the same id again.
5. Results are recorded locally *before* reporting. If reporting fails, the result stays pending and
   is delivered before any new job is claimed.
6. Anything uncertain (no confirmation, error banner, crash/worker restart after the click, channel
   lost) → `needs_review`, never an automatic retry. Local storage is supplementary; the backend must
   still enforce idempotency/reservation.

## How success is verified

1. After the single Send click the extension waits (bounded, 20 s) for the **compose window to close** or a
   **"message sent" notification**. An error banner or a timeout yields `needs_review`.
2. With **Confirm each send in Sent Items** on (the default), it then opens the **Sent Items** folder and waits
   up to 30 s for a message with the **same subject**, sent **within the last minute or so**, and, when the row shows
   an address rather than a contact's name, the **same recipient**. Only then is the job `sent`. Not found in time
   (or the folder control is missing) means `needs_review`, never an automatic resend. It then returns to the folder
   you were in. Turn the checkbox off to rely on step 1 only.

Limits of step 2: it needs the folder tree visible (a very narrow window can collapse it), reads the English
labels, and cannot compare recipients when Outlook shows a contact's display name.

## Project layout

```
manifest.json      MV3: popup, service worker, Outlook-only content script
background.js      service worker wiring (alarms, messages, tabs)
orchestrator.js    job lifecycle, recovery, pause rules (pure, injected deps)
api.js / config.js backend client, request building, validation, settings
state.js store.js logger.js   state machine & duplicate map, storage wrapper, redacting logger
outlook-dom.js     selectors, readiness, waits, text-only editor filling
outlook-flow.js    compose → validate → send → verify for one job
content.js         message bridge + progress persistence
popup.html/css/js  settings & status UI
tests/             node:test + jsdom; tests/e2e/chrome-smoke.mjs is a manual real-Chrome script
```

## Testing

```bash
npm install          # from the repo root (installs jsdom)
npm test             # unit/DOM tests (extension + backend)
npm run check        # syntax check of every extension file
```

Covered: settings validation, request construction, response parsing, bad/missing job data,
duplicate protection, state transitions, failure reporting, element detection, recipient/subject/body
validation, timeouts, uncertain send outcomes, crash recovery, manifest/permission/CSP checks.
UI tests use a **simulated** Outlook page (`tests/fixtures/outlook.js`), which proves the logic, not
that real Outlook's DOM matches.

Optional real-Chrome smoke runs (load the unpacked extension against a simulated Outlook page):

```bash
PLAYWRIGHT_DIR=<dir containing node_modules/playwright> CHROME_PATH=<chromium binary> \
  node extension/tests/e2e/chrome-smoke.mjs close   # mock backend; or: never (no confirmation → needs_review)
  node extension/tests/e2e/full-stack.mjs           # real backend code (in-memory sheet) + To/Cc check
```

### Manual testing checklist

Use a **test** Outlook account and a test backend queue with addresses you control.

- [ ] Load unpacked; popup opens; save invalid settings → errors; save valid → persists after reopening.
- [ ] Backend permission prompt appears for only your backend origin.
- [ ] No Outlook tab → *Waiting for Outlook*. Signed-out Outlook → waiting, no action taken.
- [ ] One job → New mail opens; **To** (not Cc/Bcc), subject and multi-line body are correct; with 2+ addresses, the extras appear under **Cc**; Send clicked once; mail arrives; backend shows `sent`.
- [ ] With Sent Items confirmation on: the status becomes `sent` only after the message appears in Sent Items, and Outlook returns to the folder you were in.
- [ ] Check blank-line fidelity and that no signature/formatting is damaged.
- [ ] Invalid recipient (e.g. nonexistent domain) → no send, `failed`.
- [ ] Leave a draft open → job fails with *compose window already open*, nothing sent.
- [ ] Block the status endpoint after a send → result stays pending; no second email; delivered once the endpoint returns.
- [ ] Reload the extension / kill the service worker mid-send → `needs_review`, never a resend.
- [ ] Re-deliver the same job id from the backend → not sent again.
- [ ] Wrong API key → automation pauses with a clear message.
- [ ] Rate limit (429) → honours `Retry-After`.
- [ ] Verify the real Outlook selectors in `outlook-dom.js` (`SELECTORS`) still match; adjust if not.

## Troubleshooting

| Symptom | Fix |
|---|---|
| *Waiting for Outlook* | Open and sign in at `outlook.live.com/mail/`; reload tabs opened before install. |
| *Outlook tab is not responding* | Reload the Outlook tab (content script not injected). |
| *Could not find the "New mail" button* / compose never appears | Outlook UI/language changed; update `SELECTORS`/`LABELS` in `outlook-dom.js`. |
| *Message was not found in Sent Items* (`needs_review`) | The message may still have been sent: open Sent Items and check. If Outlook is just slow, untick **Confirm each send in Sent Items**. |
| *Could not find the Sent Items folder* | The folder pane is collapsed or Outlook's labels changed; widen the window, or untick the Sent Items confirmation. |
| *Could not find the Cc control* | Outlook's Cc toggle changed; adjust `SELECTORS.cc` / `LABELS.ccToggle` in `outlook-dom.js`. |
| *Recipient did not resolve* | Outlook may resolve chips differently; adjust `enterRecipient`/`findRecipientPill`. |
| *Needs review* | Check **Sent Items**, then decide manually; press **Start** to resume. |
| *Duplicate delivery: already processed locally (sent / needs_review)* | You cleared a row's status to send it again, but the extension remembers sending it. If you really want to resend (after checking Sent Items), click **Clear job history** in the popup, then **Stop** and **Start**. |
| Cleared a `failed` row but nothing happens | Make sure no draft is open in Outlook, then **Stop** and **Start** (Start fetches immediately). |
| Backend errors | See *Recent errors* in the popup; DevTools → service worker console shows `[OEA]` logs. |

## Security

- Permissions: `storage`, `alarms`, host `https://outlook.live.com/*`; backend origin requested at runtime.
- Content script only on `https://outlook.live.com/mail/*`; it ignores messages not from this extension.
- API key only in `chrome.storage.local`; never logged. Logs hold job id, stage, result, category,
  timestamp and redacted text. No recipients, subjects or bodies. No mailbox data is read or sent.
- HTTPS enforced for the backend (HTTP only for localhost). No `eval`, no remote scripts, no `innerHTML`;
  job text is inserted as text nodes / `insertText` only.
- Any user with access to the browser profile can read the stored key; use a dedicated profile.

## Known limitations

- Only `outlook.live.com` (personal Outlook), English UI, one Outlook tab used at a time.
- Selectors/recipient-chip detection are heuristics validated only against a simulated page.
- Success evidence is the compose window closing or a sent toast, plus the optional Sent Items confirmation.
- Rows are sent strictly one at a time; Sent Items matching assumes at least about a minute between emails with the same subject.
- A sleeping computer/closed browser pauses processing; there is no job queue in the extension.
- Does not bypass Outlook sending limits, MFA, CAPTCHAs or other Microsoft challenges.
