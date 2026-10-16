# User guide

A step-by-step guide to installing, running and operating the system, from nothing to a daily routine.
If you only want to try it once, follow **Parts 1–6** and stop.

> **Reality check.** The pieces have been tested thoroughly in simulation, and the core path (compose, To/Cc,
> Send, appears in Sent Items) has been confirmed on a real Outlook Web account. The **Sent Items check** and
> **several accounts at once** have not yet run against real Outlook. Start small and watch the first emails.

## Contents

1. [What this does](#1-what-this-does)
2. [What you need](#2-what-you-need)
3. [Install the code](#3-install-the-code)
4. [Prepare your spreadsheet](#4-prepare-your-spreadsheet)
5. [Configure and start the backend](#5-configure-and-start-the-backend)
6. [Install and configure the Chrome extension](#6-install-and-configure-the-chrome-extension)
7. [Your first test (do this before anything real)](#7-your-first-test)
8. [Daily operation](#8-daily-operation)
9. [Several Outlook accounts](#9-several-outlook-accounts)
10. [Reference: statuses, settings and messages](#10-reference)
11. [Troubleshooting](#11-troubleshooting)
12. [Safety, limits and good practice](#12-safety-limits-and-good-practice)
13. [Updating, stopping, uninstalling](#13-updating-stopping-uninstalling)
14. [FAQ](#14-faq)
15. [Advanced](#15-advanced)

---

## 1. What this does

You keep your contacts and messages in an Excel file. The system sends one email at a time through **Outlook Web**,
using your real signed-in browser session, and writes the result back into the file.

```
 Excel file (G emails, H subject, I message, J status)
        ^  reads rows, writes status
        |
 Backend (Node.js, runs on your PC)  <----- HTTP (localhost) ----->  Chrome extension
                                                                          |
                                                                  drives Outlook Web
                                                                  (New mail > To/Cc > Subject >
                                                                   Body > Send > confirm)
```

Each cycle:

1. The extension asks the backend for **one** job at a random time between **10 and 20 minutes** after the previous ask
   (the very first ask happens when you press **Start**).
2. The backend takes the **first row whose Status (J) is empty**, marks it `pending`, and returns it.
3. The extension opens a new mail in Outlook Web, enters the recipients (first address = **To**, the rest = **Cc**),
   the subject and the message, and clicks **Send**.
4. It verifies the send (the compose window closes, then the message appears in **Sent Items**).
5. It reports back and the backend writes `sent` (or `failed` / `needs_review`) into column J.

It does **not** use Microsoft Graph, SMTP or any mail service. It automates the Outlook Web page you are signed in to.

## 2. What you need

| Requirement | Notes |
|---|---|
| A Windows, macOS or Linux computer that stays on | The backend and Chrome run on it. Sleep mode pauses everything. |
| **Node.js 22 or newer** | Check with `node -v`. Get it from nodejs.org. |
| **Git** | Windows: "Git for Windows" gives you *Git Bash*. |
| **Google Chrome 120+** | Or another Chromium browser that supports Manifest V3 extensions. |
| A **personal Outlook.com account** | Signed in at `https://outlook.live.com/mail/`, **English** interface. |
| Your contacts as an **.xlsx** file | Layout described in Part 4. |

Not supported: Microsoft 365 / work accounts on other domains, non-English Outlook, desktop Outlook.

## 3. Install the code

Open a terminal (Git Bash on Windows) in the folder where you keep projects:

```bash
git clone https://github.com/KW-dev1/Custom-CRM
cd Custom-CRM
npm install
```

Folders you will use:

| Path | What it is |
|---|---|
| `backend/` | The Node.js server. Your settings go in `backend/.env`. |
| `backend/data/` | Put your Excel file here (**never committed to git**). |
| `extension/` | The Chrome extension you load into the browser. |
| `docs/` | This guide. |

Optional check that everything works on your machine: `npm test` (takes about a minute).
At any time, `npm run check:source` is a **read-only** check of your data source (login, sheet/file, columns, row counts).

## 4. Prepare your spreadsheet

Put your file in `backend/data/`, for example `backend/data/contacts.xlsx`.

### Columns

Row **1 is a header** row (any names). Data starts on row 2. Only these columns matter:

| Column | Header (suggested) | You fill it? | Content |
|---|---|---|---|
| **G** | Email | Yes | The recipients, **first = To, others = Cc** (see formats below). |
| **H** | Subject | Yes | One line of text. |
| **I** | Message | Yes | The email body. Line breaks (Alt+Enter in a cell) are kept. Plain text only. |
| **J** | Status | **Leave empty** | The system writes `pending`, `sent`, `failed` or `needs_review`. |
| K *(optional)* | Note | No | A timestamp and reason for every change, if you set `COL_NOTE=K`. |
| **L** | Sent by | No | The extension profile that claimed the row (e.g. `outlook-A`). |

Other columns (company, name, LinkedIn and so on) can stay; they are ignored. Column letters can be changed in `.env`.

### Email address formats in column G

All of these work:

```
alice@example.com
alice@example.com, bob@example.com, carol@example.com
alice@example.com; bob@example.com
["alice@example.com","bob@example.com"]
Alice Smith <alice@example.com>, bob@example.com
```

- The **first valid address is To**; every other address is **Cc** (up to 20).
- Duplicates are removed (case-insensitive).
- A trailing comma or space (`alice@example.com, `) is fine.
- **A malformed address anywhere in the cell** (for example `bob@` or `a@@b.com`) makes the row `failed` with a
  reason in the note column. The system refuses to send to a partly understood recipient list.

### Writing multi-line messages in column I

The message is **plain text**; line breaks and blank lines are kept. Bold, colours, bullets, HTML, images and attachments are not supported.

**Easiest way: write the whole message on ONE line and mark the breaks.** Type `\n` (a backslash followed by the letter n)
wherever you want a line break, and `\n\n` for a blank line between paragraphs. `<br>` works the same way. Spaces around a marker are ignored.
This avoids every cell-editing and pasting problem, and works in Google Sheets and Excel.

| You type in the cell | The email shows |
|---|---|
| `Hi Ann,\n\nQuick note about Acme.\n\nThanks,\nKeith` | `Hi Ann,`<br>*(blank line)*<br>`Quick note about Acme.`<br>*(blank line)*<br>`Thanks,`<br>`Keith` |
| `Hi Ann,<br><br>Quick note.<br>Keith` | `Hi Ann,`<br>*(blank line)*<br>`Quick note.`<br>`Keith` |

Only that exact backslash-n (or `<br>`) is a marker; a lone `\` or `/` is left alone. The marker is **not** applied to the Subject, which stays one line.
The other ways below (real line breaks inside the cell) also work.

**Google Sheets**

- **Pasting a message:** **double-click the cell** (or select it and press **Enter**) so the cursor is *inside* the cell, then paste.
  If you paste while the cell is only *selected*, Google Sheets splits the text at every line break and spreads it over **many rows**.
  That leaves your message in pieces and creates junk rows; if it happens, undo (Ctrl+Z) and paste again in edit mode.
- **Typing a message:** press **Ctrl+Enter** (Windows) to start a new line inside the cell; two in a row make a blank line between paragraphs.
- **Checking it:** click the cell and look at the formula bar (or turn on *Format > Wrapping > Wrap*). The whole message, line breaks included,
  must be in that one cell. Never leave fragments of a message in the rows below it.
- Writing the message in Notepad first and pasting it in edit mode is the most reliable way.

**Excel:** press **Alt+Enter** for a new line inside a cell.

### What the backend does with each row

| Row state | What happens |
|---|---|
| J empty, G/H/I all filled | Eligible. Handed out top to bottom. |
| J empty, subject **or** message **or** address still empty | **Skipped, left untouched**: treated as "still being drafted". It is picked up once you fill it in. |
| J empty, malformed email address | Marked `failed` (reason in the note column). |
| J has any value (`pending`, `sent`, `failed`, `needs_review`, or anything you typed) | Never touched. |

### Rules for working with the file

- **Close the file in Excel** while the backend runs. Excel locks open files and the backend will not write over a
  locked file. (It then reports an error and the job is retried later; nothing is lost.)
- Edits you **saved** are picked up on the next fetch. Unsaved edits in an open Excel window are invisible.
- The backend makes a one-time **backup copy** (`yourfile.xlsx.<timestamp>.bak`) before its first write each run, and changes
  only the status, note and sender cells. Keep your own master copy anyway.
- Don't sort or delete rows while an email is `pending`.
- To **send a row again**, see [Requeue and resend](#requeue-and-resend).

## 5. Configure and start the backend

### Create your settings file

```bash
cp backend/.env.example backend/.env        # Windows (cmd): copy backend\.env.example backend\.env
```

Edit `backend/.env`. The minimum you must set:

```ini
API_KEYS=pick-a-long-random-string-at-least-16-characters
DATA_SOURCE=xlsx
XLSX_PATH=./data/contacts.xlsx
COL_NOTE=K
```

(`./data/...` is relative to the `backend` folder. An absolute path such as `D:\Projects\...\contacts.xlsx` also works.)

### All backend settings

| Setting | Default | Meaning |
|---|---|---|
| `API_KEYS` | *(required)* | The shared secret the extension sends. 16+ characters. Several keys separated by commas are allowed (handy for rotating). |
| `PORT` / `HOST` | `3000` / `127.0.0.1` | Where it listens. `127.0.0.1` means "this computer only", which is the safe choice. |
| `DATA_SOURCE` | `xlsx` | `xlsx` (local Excel file) or `google` (Google Sheets, see Part 15). |
| `GOOGLE_SERVICE_ACCOUNT_FILE` | *(google only)* | Path to the service-account JSON key. Alternative: OAuth values (Part 15). |
| `XLSX_PATH` | *(required for xlsx)* | Path to the `.xlsx` file. |
| `XLSX_SHEET` | first sheet | Worksheet name, if your file has several. |
| `HEADER_ROWS` | `1` | How many header rows to skip. |
| `COL_EMAILS` `COL_SUBJECT` `COL_BODY` `COL_STATUS` | `G` `H` `I` `J` | Column letters. |
| `COL_NOTE` | *(off)* | Column for timestamped notes, e.g. `K`. Recommended. |
| `COL_SENDER` | `L` | Column recording which profile claimed the row. Empty turns it off. |
| `MIN_CLAIM_INTERVAL_SEC` | `300` | A server-side floor: the same profile cannot be given jobs closer together than this. For testing set `0`. |
| `MAX_CC` | `20` | Maximum Cc recipients per email. |
| `STATUS_PENDING` `STATUS_SENT` `STATUS_FAILED` `STATUS_NEEDS_REVIEW` | `pending` `sent` `failed` `needs_review` | The words written into the status column. |

### Start it

```bash
npm run start:backend
```

You should see lines like `"event":"data_source"` and `"event":"listening"`. **Leave this terminal open**: closing it stops the backend.

Check it is alive: open `http://localhost:3000/health` in a browser. It should show `{"ok":true}`.

What you will see in this terminal while it runs (JSON lines, never containing email addresses or message text):

| Log event | Meaning |
|---|---|
| `claimed` | A row was handed to an extension (`row`, `client`, `cc` count). |
| `status_written` | A result was written to the sheet. |
| `row_invalid` | A row with malformed data was marked `failed`. |
| `rows_incomplete` | Rows skipped because subject/message/address are still empty. |
| `wrong_client` | A profile tried to report a row claimed by another profile (refused). |
| `data_source_error` | A problem reading or writing the spreadsheet, e.g. "Excel file not found" or "file is open in Excel" (or a Google error if you use Google Sheets). |

## 6. Install and configure the Chrome extension

### Install (once per Chrome profile)

1. Open `chrome://extensions`.
2. Turn on **Developer mode** (top right).
3. Click **Load unpacked** and select the **`extension`** folder inside the project.
4. Click the puzzle-piece icon and **pin** "Outlook Email Automation".

### Configure

Open the extension popup and fill in **Settings**:

| Field | What to enter |
|---|---|
| **Backend API base URL** | `http://localhost:3000` (backend on the same computer). A remote backend must use `https://`. |
| **Profile name** | A short label for this Chrome profile / Outlook account, e.g. `outlook-A`. Use a **different** name in every profile. **Labels only, no email addresses.** Optional with a single account. |
| **API key / token** | Exactly the `API_KEYS` value from `backend/.env`. |
| **Confirm each send in Sent Items** | On by default. After Send, the extension opens Sent Items and requires the message to be listed there before it reports `sent`. Adds about 10 seconds per email. |
| *Advanced* | Leave as is. These are the API paths, auth header, **Fetch every min/max minutes** (10 and 20) and the status words. |

Click **Save settings**. Chrome asks permission for the extension to contact `localhost`: **allow** it.
"Saved." appears under the button.

### Popup controls and what they mean

| Control | Does |
|---|---|
| **Automation enabled** (tick box) | Master switch. When off, nothing runs. |
| **Start** | Runs now (first fetch is immediate), then every random 10–20 minutes. |
| **Stop** | Pauses. A job already in flight finishes its safety steps. |
| **Last result** | The most recent job and its verified outcome. |
| **Recent errors** | The last few problems with a plain explanation. |
| **Clear job history** | Makes the extension forget which jobs it already processed (see [Requeue and resend](#requeue-and-resend)). |

The coloured badge at the top shows the state:

| Badge | Meaning |
|---|---|
| Not configured | A required setting is missing or invalid; the message under the title says which. |
| Waiting for Outlook | No Outlook tab, it is signed out, still loading, or the tab needs a reload. |
| Ready | Idle; the message says when the next check happens. |
| Processing email | A send is in progress. |
| Email sent | The last email was **verified** as sent. (Never shown unless verified.) |
| Failed | The last job failed **before sending**; the reason is shown. |
| Needs review | Send was clicked but the result could not be confirmed. **Check Sent Items.** Automation paused. |
| Paused | Stopped by you, disabled, or paused automatically after a problem. |

## 7. Your first test

Do this with **your own addresses**, never with your real contact list. It sends **one real email**.

1. Make a separate file `backend/data/test.xlsx`: a header row plus **one** row with two of your own addresses
   in G (`me@outlook.com, other-address-of-mine@gmail.com`), a subject in H, a two-paragraph message in I, J empty.
   Close Excel.
2. In `backend/.env` set `XLSX_PATH=./data/test.xlsx` and `MIN_CLAIM_INTERVAL_SEC=0`. Start the backend.
3. In Chrome, open `https://outlook.live.com/mail/`, sign in, and make sure **no draft is open**.
4. In the popup: save the settings, tick **Automation enabled**, click **Start**.
5. Watch: Outlook opens a new mail, the first address becomes **To**, the second **Cc**, the subject and body fill in,
   Send is clicked, Outlook visits **Sent Items** and returns to where you were. The popup shows **Email sent**.
6. Check: the email arrived (To and Cc both), the line breaks are right, and the file shows `sent` in J with a note in K
   and your profile name in L. (Open the file only after the backend finishes.)

If anything fails, jump to [Troubleshooting](#11-troubleshooting).

## 8. Daily operation

### Pre-flight checklist

- [ ] The backend is running (terminal open, `/health` works).
- [ ] The Excel file is **closed**; H and I are filled for the rows you want sent; J is empty for them.
- [ ] Chrome is running, with an Outlook tab open, **signed in**, **no draft open**, and the tab not minimized or sleeping.
- [ ] The popup shows **Automation enabled** ticked; press **Start**.
- [ ] `MIN_CLAIM_INTERVAL_SEC` is back to `300` (or higher) for real use.

### What to expect

- The first email is sent right after **Start**. After that, one email every random **10–20 minutes** (roughly 3–6 per hour per account).
- Each send takes about **10–40 seconds**. During that time Outlook is being driven by the extension: **don't type in that tab.**
- The extension is event-driven and survives Chrome putting its background worker to sleep; it resumes on its own.
- Rows are processed **top to bottom** among those with empty J.

### Watching it

- **The popup** (badge, last result, recent errors).
- **The Excel file** (columns J, K, L). Close and reopen it to refresh; or filter on J.
- **The backend terminal** (live log lines).
- **Outlook's Sent Items.**

### Automatic safety stops

Automation **pauses itself** (badge: Paused or Needs review) when:

- an email ends as **`needs_review`** (unconfirmed);
- the backend rejects the API key;
- **3 failures in a row** occur;
- the backend permanently refuses a result.

It never resends an email on its own after any of these.

### Requeue and resend

| Situation | What to do |
|---|---|
| Row is **`failed`** (never sent) | Fix the cause (e.g. a draft was open, a bad address). Close Excel, **clear J** (and K/L) of that row, save. In the popup click **Stop**, then **Start**. It is retried automatically. |
| Row is **`needs_review`** | First open **Outlook Sent Items** and check whether it really went. If it **did**, set J to `sent` yourself. If it **did not**, clear J (and K/L), click **Clear job history** in the popup, then **Stop** and **Start**. |
| Row is **`pending`** for a long time | The Chrome profile that took it was probably closed mid-job. Column L says which profile. Check that account's Sent Items, then handle it as above. |
| You **want to send an already-sent row again** | Clear J (and K/L), click **Clear job history** in the popup, then **Stop** and **Start**. (The history exists so a job is never sent twice by accident.) |

### Pause or stop

Click **Stop** (or untick **Automation enabled**) to pause. To stop everything, also press `Ctrl+C` in the backend terminal.

## 9. Several Outlook accounts

Use one **Chrome profile per Outlook account**, each profile signed in to its own account.

1. In each Chrome profile install the extension (Part 6) and open its own Outlook tab.
2. Give each popup the **same backend URL and API key** but a **different Profile name** (`outlook-A`, `outlook-B`, ...).
3. Press **Start** in each. They work in parallel against the same backend and file.

How it behaves:

- A row is marked `pending` **before** it is handed out, under a lock, so two profiles never get the same row.
- **First come, first served**: whichever profile asks next gets the next empty row.
- Column **L** records which profile claimed each row.
- Pacing is **per profile**: each account sends one email per random 10–20 minutes on its own.
- A result is accepted only from the profile that claimed the row. **Don't rename a profile while it has a job in flight.**
- All profiles need Chrome running with their Outlook tab open and signed in. One backend serves all of them.

## 10. Reference

### Status column (J)

| Value | Meaning |
|---|---|
| *(empty)* | Waiting to be sent. |
| `pending` | Handed to an extension; the email is being sent, or something interrupted it (see above). |
| `sent` | **Verified** sent (compose closed and, if enabled, listed in Sent Items). |
| `failed` | **Provably not sent** (failed before Send, or a malformed row). Safe to requeue. |
| `needs_review` | Send was clicked but the outcome is unclear. **Check Sent Items before retrying.** |

### Extension messages and what to do

| Message / code | Meaning | What to do |
|---|---|---|
| Waiting for Outlook | No Outlook tab, signed out, or wrong page. | Open and sign in at `outlook.live.com/mail/`; reload tabs opened before installing. |
| Outlook tab is not responding | The page was opened before the extension was installed/reloaded. | Reload the Outlook tab. |
| `COMPOSE_ALREADY_OPEN` | A draft is already open. | Close or discard it. |
| `NEW_MAIL_MISSING` / `CC_FIELD_MISSING` | A button was not found. | Outlook's layout changed or the window is very narrow; see Troubleshooting. |
| `RECIPIENT_FAILED` / `RECIPIENT_REJECTED` | Outlook did not turn the typed address into a recipient. | Check the address; send the console log if it repeats. |
| `BODY_FORMAT_FAILED` | Outlook's editor did not keep the message's line breaks (they came out merged into fewer lines). **Nothing was sent.** The extension tries real paragraphs, then line breaks, then building lines directly, and reads the editor back after each; this means none worked. | Check the message cell really contains line breaks (see *Writing multi-line messages*). If it does, send the maintainer the cell text and a screenshot of the compose window. |
| `VALIDATION_FAILED` | The filled-in message did not match the job. Nothing was sent. | Retry; report if it repeats. |
| `SEND_ERROR_SHOWN` | Outlook showed an error after Send. | **Check Sent Items** (needs review). |
| `NO_CONFIRMATION` | Send clicked, no confirmation in time. | **Check Sent Items** (needs review). |
| `SENT_ITEMS_NOT_FOUND` | Compose closed but the message was not matched in Sent Items in 30 s. The message ends with counts of what the check saw (rows listed, with the subject, with the recipient, within the time window). | **Check Sent Items.** If the email is there, the check is giving a false alarm: untick *Confirm each send in Sent Items* and send the counts to the maintainer. |
| `SENT_ITEMS_UNAVAILABLE` | The Sent Items folder entry was not visible. | Widen the Outlook window so the folder list shows, or untick the confirmation. |
| *Duplicate delivery ... already processed locally (sent / needs_review)* | The extension remembers sending this row. | If you really want to resend: **Clear job history**. |
| *The backend refused the result for the last job* | Row claimed by another profile, edited, or removed. Automation paused; email **not** resent. | Check the sheet, then **Start**. |

### Backend HTTP errors (shown in the popup)

| Code | Meaning |
|---|---|
| 401 | The API key does not match `API_KEYS`. |
| 429 | Claimed too soon (`MIN_CLAIM_INTERVAL_SEC`) or too many wrong keys. Waits and retries. |
| 502 | The backend could not read/write the file (not found, open in Excel, ...). The backend terminal says which. |
| 409 `CONFLICT` / `WRONG_CLIENT` | The row already has a different final status, or belongs to another profile. |

## 11. Troubleshooting

| Symptom | Likely cause and fix |
|---|---|
| Popup says **Not configured** even though the form is filled | You did not click **Save settings**. Click it and allow the permission prompt. |
| `Backend server error (HTTP 502)` | See the backend terminal. Usually `Excel file not found` (check the **exact** file name and that Windows is not hiding a second `.xlsx`), or the file is open in Excel. |
| `Backend rejected the credentials (HTTP 401)` | The popup key differs from `API_KEYS`. Copy it exactly. |
| Backend says `XLSX_PATH is required` although you use Google Sheets | `DATA_SOURCE=google` is missing from `backend/.env` (or has a `#` in front, or spaces around `=`). Edits to `.env` must be **added** to the file, not replace it. `npm run check:source` prints `Data source: ...` on its first line to confirm which source is active. |
| `check:source` says **Google rejected the service account key** | The message now says why. Usually: **the Windows clock is off** (Settings > Time & language > Date & time > *Sync now*, and turn on *Set time automatically*); or the key file was **edited** or is from a **deleted key** (create a NEW JSON key in Google Cloud and use the downloaded file unchanged); or the service account was deleted/disabled. |
| `Could not reach the backend` | The backend is not running, or the URL/port is wrong. |
| Nothing happens after **Start** | Outlook tab missing or signed out; no row has H, I and G filled with J empty; or the next-fetch time has not arrived (the message says when). |
| Row marked `failed` instantly (a fraction of a second) | A draft was already open, or **Duplicate delivery** from a previous failure (fixed: failed rows now retry; update to the latest code). |
| The whole message arrives on ONE line | The cell holds one line with no break markers: type `\n` (backslash n) where you want breaks, or `<br>`. Otherwise the lines may have been split into rows below the cell, or Outlook merged them. The extension now reads the editor back and refuses to send a message whose lines came out merged (`BODY_FORMAT_FAILED`). |
| The message arrives in pieces, or rows below it contain fragments | The text was pasted while the cell was only selected, so Google Sheets split it into rows. Delete the fragments, put the whole message back into ONE cell (double-click the cell first, then paste). |
| Email composes but `RECIPIENT_FAILED` | Update the extension to the latest version (older versions mis-detected recipient chips) and reload it. |
| `needs_review` although the email arrived | Outlook was slow to list it in Sent Items. Set J to `sent` by hand; consider unticking the Sent Items confirmation. |
| The backend log says `file is open in Excel` | Close Excel (the lock file `~$yourfile.xlsx` must disappear). |
| After updating, behaviour is unchanged | Reload the extension in `chrome://extensions` **and** reload the Outlook tab; restart the backend. |
| Windows: `cp` or `npm` not found | Use **Git Bash**, and make sure Node.js is installed (`node -v`). |

**Where to look for more detail**

- Backend: the terminal output.
- Extension background: `chrome://extensions` > the extension > **service worker** link (Console). Lines start with `[OEA]`.
- Outlook page: press F12 on the Outlook tab, Console, enable the **Verbose** level, filter on `OEA`.

When reporting a problem, include the popup message, the backend log lines around it, and the Console lines.

## 12. Safety, limits and good practice

- **Account protection.** Microsoft can throttle or lock accounts that send many similar emails automatically. The 10–20 minute
  spacing is deliberately conservative; do not shorten it much. Start with a handful of rows, check Sent Items, and watch the
  account for warnings. Respect recipients and local anti-spam rules. You are responsible for what you send.
- **No silent double sends.** A row is `pending` before it is released; a sent or unconfirmed job is never retried
  automatically; the extension keeps a local record of processed jobs; unconfirmed results always go to `needs_review`.
- **Your data stays local.** The Excel file never leaves your computer. The backend listens on `127.0.0.1` only. The API key
  lives in `backend/.env` and in the extension's local storage. Logs never contain addresses or message text.
- **Keep secrets out of git.** `.env`, spreadsheets, CSVs and backups are git-ignored. Never share `.env`.
- **The Outlook tab must stay open and signed in** for every profile that is running. Keep it in a visible, non-minimized window.
  (Behaviour with the tab in a background/hidden state has not been verified.)
- **Don't use Outlook in that tab while it runs.** Use another window for your own mail.
- **English Outlook only**, personal Outlook.com accounts, one message at a time, no attachments, plain text bodies,
  up to 20 Cc recipients, no Bcc.
- **Templates/merge fields are not supported.** Write each subject and message in the spreadsheet.
- The system does not bypass Microsoft sign-in, MFA, CAPTCHAs or Outlook's own sending limits.

## 13. Updating, stopping, uninstalling

**Update** to the newest version:

```bash
cd Custom-CRM
git checkout main
git pull
npm install
```

Then: restart the backend (Ctrl+C, `npm run start:backend`), click the **reload** icon for the extension in
`chrome://extensions`, and reload the Outlook tab.

**Stop**: click **Stop** in the popup; press `Ctrl+C` in the backend terminal.

**Uninstall**: remove the extension in `chrome://extensions`; delete the project folder. Your spreadsheet and its `.bak`
backups are ordinary files you can keep or delete.

## 14. FAQ

**Can I use a work/school (Microsoft 365) account?** Not tested or supported. Only `outlook.live.com` personal accounts.

**Can it send attachments, Bcc, HTML or merge fields?** No. Plain text, To + Cc, written per row.

**Does Chrome have to stay open?** Yes, with the Outlook tab open and signed in. The backend must also be running.

**What if the computer sleeps or restarts?** Processing pauses. If it slept mid-send, that row may be left `pending` or end
as `needs_review`; check Sent Items.

**Can I change the 10–20 minute pacing?** Yes: *Advanced > Fetch every, min/max minutes* in the popup. Faster pacing raises
the risk of account throttling.

**Why do some rows stay empty and untouched?** They are missing an address, subject or message. Fill them in and they are picked up.

**Why was a row marked `failed` with no explanation?** Set `COL_NOTE=K` in `.env` so the reason is written next to the status.

**Can two people use it?** Run one backend per spreadsheet. Several Chrome profiles (accounts) can share it; two backends on
one file are not safe.

**How do I run it on another PC or a server?** The extension requires `https://` for a non-local backend, so use a TLS
reverse proxy or tunnel in front of the backend, and keep the file on the backend's machine. Local use is simpler.

## 15. Advanced

### Google Sheets instead of Excel

Recommended if you add rows every day: you can edit the sheet while the backend runs (no "close the file" rule), and the
Status column fills in live. The data layout (columns G, H, I, J, row rules, statuses) is identical to Excel.

**Step 1: a native Google Sheet.** Create the sheet in Google Sheets, or open your `.xlsx` there and choose
**File > Save as Google Sheets**. (A tab that still shows an `.XLSX` badge cannot be edited by the backend.)

**Step 2: a service account** (a free robot login inside your own Google account; a personal Gmail is enough):

1. Go to `console.cloud.google.com` and create a **project**.
2. **APIs & Services > Library**: search **Google Sheets API** and click **Enable**.
3. **IAM & Admin > Service Accounts > Create service account**: any name, skip the optional steps, Done.
4. Open the service account > **Keys > Add key > Create new key > JSON**. A file downloads. Move it to
   `backend/secrets/service-account.json`. **Keep it secret** (the folder is git-ignored).
5. Copy the service account's **email address** (`something@your-project.iam.gserviceaccount.com`).
6. In your Google Sheet click **Share**, paste that email, choose **Editor**, and send.

**Step 3: settings.** In `backend/.env`:

```ini
DATA_SOURCE=google
SHEET_URL=https://docs.google.com/spreadsheets/d/<your sheet id>/edit?gid=<tab id>
GOOGLE_SERVICE_ACCOUNT_FILE=./secrets/service-account.json
COL_NOTE=K
```

Paste the sheet URL exactly as it appears in your browser; the `gid=` part selects the tab.

**Step 4: check it (read-only, changes nothing):**

```bash
npm run check:source
```

You should see `[ok] Google login works`, `[ok] Found the sheet/file and the right tab`, the header names for columns G, H, I, J,
and counts of your rows. Typical problems it names: key file not found, wrong kind of file downloaded, "does <email> have edit
access" (share the sheet with that email), or a wrong sheet link.

**Step 5: run.** `npm run start:backend`, then follow Parts 6 and 7 of this guide. Try one test row with your own addresses first.

Notes: keep Excel-only rules out of your head (no locking here). Do not edit a row while it is `pending`: its content is how the
backend recognises it. The 7-day token expiry applies only to the OAuth alternative (see `backend/README.md`), not to a service account.
Details and the OAuth alternative are in [`backend/README.md`](../backend/README.md).

### API contract

The extension and backend speak a small HTTP API (claim, report status, job state, health). It is documented in
[`backend/README.md`](../backend/README.md#api) and [`extension/README.md`](../extension/README.md#backend-api-contract); the
paths, auth header and status words are configurable in the popup's *Advanced* section.

### Tests

```bash
npm test             # unit, DOM and HTTP tests (extension + backend)
npm run check        # syntax check
```

Optional real-Chromium scripts (need Playwright and a Chromium binary; see the headers of the files):

```bash
node extension/tests/e2e/full-stack.mjs       # one profile: backend + extension + simulated Outlook
node extension/tests/e2e/multi-profile.mjs 3  # three profiles at once
```

These use a **simulated** Outlook page. The manual checklist in [`extension/README.md`](../extension/README.md) covers real Outlook.

### Project structure

```
backend/        Node.js API (Express) + Excel/Google data sources
  src/          config, rows (parsing), jobs (claim/report rules), xlsx-store, app (HTTP), index
extension/      Chrome extension (Manifest V3, no build step)
  background.js, orchestrator.js   scheduling, recovery, safety rules
  outlook-dom.js, outlook-flow.js  all Outlook page automation (selectors live in SELECTORS)
  popup.*                          settings UI
docs/           this guide
scripts/        icon generator, syntax check
```

When Outlook changes its page, the selectors in `extension/outlook-dom.js` (`SELECTORS` and `LABELS`) are the one place to adjust.
