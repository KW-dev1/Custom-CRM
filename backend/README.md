# Backend: email jobs from a Google Sheet

Small Node.js (Express) API that turns rows of a Google Sheet into email jobs for the Chrome
extension and writes the results back to the sheet.

> **Status:** unit and HTTP tests pass against an **in-memory sheet** and mocked Google HTTP. The
> Google OAuth and Sheets calls have **not been run against live Google** (the build environment
> cannot reach it). Do the [first-run checklist](#first-run-checklist) before relying on it.

## Sheet layout

| Column | Meaning |
|---|---|
| **G** | Email addresses: a JSON array `["a@x.com","b@y.org"]`, or a list separated by commas / semicolons / spaces / new lines. The **first valid address is To**, the rest are **Cc**. Duplicates are removed. |
| **H** | Email subject (one line). |
| **I** | Message content (line breaks are kept). |
| **J** | Status. Starts **empty** → `pending` when handed out → `sent`, `failed` or `needs_review`. |
| K (optional) | `COL_NOTE`: timestamp + reason for every status change. |

Row 1 is treated as a header (`HEADER_ROWS`). Column letters, status words and the header count are
configurable.

### Status rules

- Only rows whose **J is empty** (and that have data) are ever handed out, top to bottom.
- Handing out a row writes `pending` **first**. A `pending` row is **never retried automatically**.
- `sent`: the extension saw Outlook accept the message.
- `failed`: provably **not** sent (e.g. recipient wouldn't resolve). Safe to retry: clear J.
- `needs_review`: Send was clicked but not confirmed. **Check Outlook's Sent Items** first, then clear
  J to retry or set it to `sent`.
- A row with a malformed address, or empty subject/content, is marked `failed` with the reason in the note
  column (we refuse to send to a partly-parsed recipient list).
- A final status is never overwritten by a different one, and any status you type by hand is respected.
- Stuck on `pending` after a crash? Clear J to requeue (after checking Sent Items if unsure).

## Setup

1. **Google Cloud**: create a project, enable the **Google Sheets API**, configure the OAuth consent
   screen, then create an **OAuth client ID of type "Desktop app"**.
2. **Publish the OAuth app** ("In production" on the consent screen; it can stay unverified for personal
   use). **While it is in "Testing" mode Google expires refresh tokens after 7 days** and the backend
   will stop working with an `invalid_grant` error.
3. The Google user you authorise must have **edit access** to the sheet.
4. Configure:
   ```bash
   cp backend/.env.example backend/.env      # then edit it (never commit .env)
   ```
   Fill `API_KEYS` (a long random string; the extension's API key), `GOOGLE_CLIENT_ID`,
   `GOOGLE_CLIENT_SECRET`. The sheet URL is already set to your sample sheet.
5. Get the refresh token once (needs a browser, run on your own machine):
   ```bash
   npm install
   npm run auth --workspace backend      # open the printed URL, approve, copy GOOGLE_REFRESH_TOKEN into .env
   ```
6. Run:
   ```bash
   npm run start:backend
   curl http://127.0.0.1:3000/health
   ```

### Connecting the extension

- In the extension popup set **Backend API base URL** to this server and **API key** to one of `API_KEYS`
  (default auth header `Authorization: Bearer <key>`, which matches this API; no advanced changes needed).
- The extension requires **HTTPS** except for `http://localhost`/`127.0.0.1`. Either run the backend on the
  same computer as Chrome and use `http://localhost:3000`, or put it behind a TLS reverse proxy / tunnel
  (Caddy, nginx, Cloudflare Tunnel...) and use the `https://` URL.
- The extension asks for one job every random 10–20 minutes.

## API

All `/api` calls need `Authorization: Bearer <key>` (or `X-Api-Key`).

| Call | Result |
|---|---|
| `POST /api/email-jobs/claim` | `200 {"job": {id, to, cc[], subject, body, metadata:{row}}}` or `204` when nothing is waiting. Marks the row `pending`. `429` + `Retry-After` if called more often than `MIN_CLAIM_INTERVAL_SEC`. |
| `POST /api/email-jobs/:id/status` | Body `{"status": "processing\|sent\|failed\|needs_review", "message"?}` → `200 {id,status,applied}`. Idempotent. `409` if the row already has a different final status, `404` if the row was deleted/edited. |
| `GET /api/email-jobs/:id` | `200 {id,status}`: `empty|pending|sent|failed|needs_review|other`. |
| `GET /health` | `200 {ok:true}`, no auth, no Google call. |

Errors are JSON: `{"error":{"code","message"}}`. Google outages return `502`.

A job id looks like `r5-1a2b3c4d5e6f` (row + fingerprint of G/H/I). If the row was edited after being
handed out, or deleted, the result is refused (404) rather than written to the wrong row; if the sheet was
sorted, the row is found again by its content.

## Configuration

See `backend/.env.example`. Notable: `MIN_CLAIM_INTERVAL_SEC` (default 300, a server-side floor on how
often jobs are released), `MAX_CC` (default 20), `COL_*`, `STATUS_*`, `HEADER_ROWS`.

## Safety and limitations

- Run **exactly one** backend instance per sheet. Claims are serialised with an in-process lock; Google
  Sheets has no compare-and-set, so two instances could hand out the same row.
- If a job is handed out but the response is lost, the row stays `pending` (by design: never auto-resent).
- API keys are compared in constant time; repeated bad keys are throttled. The key and Google secrets
  live only in environment variables. Logs contain row numbers and ids, never addresses or message text.
- Sheet writes use `RAW` input, so cell contents are never evaluated as formulas.
- Edits to column G/H/I of a row that is already `pending` make its eventual result unrecordable (404).
- Sheets API quota is ~60 writes/min per user; at one email per 10–20 minutes this is irrelevant.

## First-run checklist

1. `npm run auth --workspace backend` succeeds and you have a refresh token.
2. Start the backend; add one test row (your own address in G, subject, content; J empty).
3. `curl -X POST -H "Authorization: Bearer <key>" http://127.0.0.1:3000/api/email-jobs/claim` returns the
   job and J becomes `pending`. A second call returns `204`.
4. Report: `curl -X POST -H "Authorization: Bearer <key>" -H 'content-type: application/json' -d '{"status":"sent"}' http://127.0.0.1:3000/api/email-jobs/<id>/status` → J becomes `sent`.
5. Reset J and run the real flow with the extension (see the extension README checklist).
6. Test a row with 3 addresses: first is To, other two are Cc.

## Tests

```bash
npm run test:backend
```
