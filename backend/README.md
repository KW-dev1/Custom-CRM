# Backend: email jobs from a local Excel file (or Google Sheet)

Small Node.js (Express) API that turns rows of a spreadsheet into email jobs for the Chrome
extension and writes the results back to the status column. The default data source is a **local
`.xlsx` file** on the same computer; Google Sheets is an optional alternative (`DATA_SOURCE=google`).

> **Status:** the Excel source has been exercised end to end (real `.xlsx` files, including a copy of your
> `contact_out.xlsx`, through the real HTTP API). The **Google** source has only been tested against mocked
> HTTP, never live Google. Neither has been run against real Outlook yet.

## Quick start (local Excel file)

```bash
npm install
mkdir -p backend/data && cp /path/to/contact_out.xlsx backend/data/     # kept out of git
cp backend/.env.example backend/.env   # set API_KEYS (long random string); XLSX_PATH is already ./data/contact_out.xlsx
npm run start:backend                  # listens on http://127.0.0.1:3000
```

In the extension popup use **Backend API base URL** `http://localhost:3000` and the same key. Fill in
columns **H (Subject)** and **I (Message)** for each row you want sent, save and **close Excel**.

### Working with the Excel file

- **Close the file in Excel** (or at least save and close) while the backend runs. Excel locks open files;
  the backend detects the lock, refuses to write, and the job is retried later rather than lost or sent without a record.
- The backend re-reads the file every time, so anything you saved is picked up on the next fetch.
  Unsaved changes in an open Excel window are invisible to it.
- A one-time backup copy (`contact_out.xlsx.<timestamp>.bak`) is made before the first write of each run.
- Writes touch only the status (and optional note) cell; other cells and formatting are preserved.
  Very exotic Excel features (macros, pivot tables, embedded images) may not survive a save. Keep your master copy elsewhere.
- Rows whose **subject or message is still empty are skipped untouched** until you fill them in. Rows with a
  malformed email address are marked `failed` with the reason in the optional note column (`COL_NOTE`).

## Sheet layout

| Column | Meaning |
|---|---|
| **G** | Email addresses: a JSON array `["a@x.com","b@y.org"]`, or a list separated by commas / semicolons / spaces / new lines. The **first valid address is To**, the rest are **Cc**. Duplicates are removed. |
| **H** | Email subject (one line). |
| **I** | Message content (line breaks are kept). |
| **J** | Status. Starts **empty** → `pending` when handed out → `sent`, `failed` or `needs_review`. |
| K (optional) | `COL_NOTE`: timestamp + reason for every status change. |
| **L** | `COL_SENDER`: the extension **profile** (Outlook account) that claimed the row, e.g. `outlook-A`. Written when the row is claimed; never changed afterwards. Set `COL_SENDER=` to disable. |

Row 1 is treated as a header (`HEADER_ROWS`). Column letters, status words and the header count are
configurable.

### Status rules

- Only rows whose **J is empty** (and that have data) are ever handed out, top to bottom.
- Handing out a row writes `pending` **first**. A `pending` row is **never retried automatically**.
- `sent`: the extension saw Outlook accept the message.
- `failed`: provably **not** sent (e.g. recipient wouldn't resolve). Safe to retry: clear J.
- `needs_review`: Send was clicked but not confirmed. **Check Outlook's Sent Items** first, then clear
  J to retry or set it to `sent`.
- A row with a malformed email address is marked `failed` with the reason in the note column (we refuse to
  send to a partly-parsed recipient list). A row with no address, subject or message **yet** is just skipped.
- A final status is never overwritten by a different one, and any status you type by hand is respected.
- Stuck on `pending` after a crash? Clear J to requeue (after checking Sent Items if unsure).

## Setup for Google Sheets (optional, `DATA_SOURCE=google`)

Best when you add rows daily: edit the sheet any time while the backend runs (no file locking), and statuses appear live.

**1. Make a native Google Sheet.** If you start from an `.xlsx`, open it in Google Sheets and choose **File > Save as Google Sheets**
(the tab must not show an `.XLSX` badge, because the Sheets API cannot edit Office-mode files).

**2. Create a service account** (free, inside your own Google account; no company account needed):

1. `console.cloud.google.com` > create a **project**.
2. **APIs & Services > Library** > **Google Sheets API** > **Enable**.
3. **IAM & Admin > Service Accounts > Create service account**; any name; skip the optional steps.
4. Open it > **Keys > Add key > Create new key > JSON**. Save the file as `backend/secrets/service-account.json`
   (that folder is git-ignored; treat the file like a password).
5. Copy the service account's **email address** (looks like `name@project.iam.gserviceaccount.com`).
6. In your Google Sheet click **Share**, paste that email, give it **Editor**.

**3. Configure `backend/.env`:**

```ini
DATA_SOURCE=google
SHEET_URL=<the sheet's browser URL, including gid=...>
GOOGLE_SERVICE_ACCOUNT_FILE=./secrets/service-account.json
COL_NOTE=K
```

**4. Check the connection (read-only, changes nothing):**

```bash
npm run check:source
```

It confirms the login, shows which header sits in each configured column and counts the rows. If a step is wrong it says which
(for example: key file missing, wrong kind of file, or which email to share the sheet with).

**5. Start** with `npm run start:backend`; the startup log prints `shareSheetWith` (the service account's email).
Then do the first-run checklist below with one test row.

*Alternative login as your own Google user (OAuth):* create an OAuth client of type **Desktop app**, **publish** the app (in "Testing" mode
tokens expire after 7 days), set `GOOGLE_CLIENT_ID` / `GOOGLE_CLIENT_SECRET`, run `npm run auth --workspace backend` once, and put the printed
`GOOGLE_REFRESH_TOKEN` in `.env` (use this instead of `GOOGLE_SERVICE_ACCOUNT_FILE`).

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
| `POST /api/email-jobs/claim` (header `X-Client-Id: outlook-A`, optional) | `200 {"job": {id, to, cc[], subject, body, metadata:{row}}}` or `204` when nothing is waiting. Marks the row `pending`. `429` + `Retry-After` if the same profile calls more often than `MIN_CLAIM_INTERVAL_SEC`. |
| `POST /api/email-jobs/:id/status` | Body `{"status": "processing\|sent\|failed\|needs_review", "message"?}` → `200 {id,status,applied}`. Idempotent. `409` if the row already has a different final status or was claimed by another profile (`WRONG_CLIENT`), `404` if the row was deleted/edited. |
| `GET /api/email-jobs/:id` | `200 {id,status}`: `empty|pending|sent|failed|needs_review|other`. |
| `GET /health` | `200 {ok:true}`, no auth, no Google call. |

Errors are JSON: `{"error":{"code","message"}}`. Google outages return `502`.

A job id looks like `r5-1a2b3c4d5e6f` (row + fingerprint of G/H/I). If the row was edited after being
handed out, or deleted, the result is refused (404) rather than written to the wrong row; if the sheet was
sorted, the row is found again by its content.

## Configuration

See `backend/.env.example`. Notable: `MIN_CLAIM_INTERVAL_SEC` (default 300, a server-side floor on how
often jobs are released), `MAX_CC` (default 20), `COL_*`, `STATUS_*`, `HEADER_ROWS`.

## Several Outlook accounts (one Chrome profile each)

Install the extension in every Chrome profile and give each one a **Profile name** in its popup (`outlook-A`,
`outlook-B`, ...). All profiles point at the same backend URL and use the same API key.

- **No double sends:** a row is marked `pending` (under a lock) *before* it is handed out, so two profiles asking
  at the same moment always get different rows. Verified with three real browser profiles pressing Start together.
- **First come, first served:** whichever profile asks next gets the next empty row.
- **Recorded in the sheet:** column L shows which profile claimed each row.
- **Per-profile pacing:** `MIN_CLAIM_INTERVAL_SEC` applies to each profile separately, so accounts don't throttle each other.
- **Only the claimer reports:** a result is accepted only from the profile that claimed the row (`409 WRONG_CLIENT`
  otherwise). Don't rename a profile while it has a job in flight.
- Profile names are 1-40 letters, digits, spaces, dots, dashes or underscores. **Use labels, not email addresses.**
- The header is `X-Client-Id`. Requests without it count as the profile `default`.
- **Stuck rows:** if a profile's Chrome closes mid-job, its row stays `pending` (never auto-retried). Filter column J for
  `pending` and read column L / the note timestamp to see who and when. After checking that account's Sent Items,
  clear J (and L) to requeue the row.
- Each Chrome profile must stay running with Outlook open and signed in to *its* account.

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
