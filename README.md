# Custom CRM

> **New here? Read the [User guide](docs/USER_GUIDE.md)**: install, set up the spreadsheet, run, test, daily use and troubleshooting, step by step.

A custom system, built in parts:

| Part | Stack | Status |
|---|---|---|
| Chrome extension that sends queued emails through Outlook Web | MV3, plain JS | Implemented: see [`extension/README.md`](extension/README.md) |
| Backend API (hands out one row at a time, records results) | Node.js + Express | Implemented: see [`backend/README.md`](backend/README.md) |
| Data | Local Excel file by default (G emails, H subject, I message, J status); Google Sheets optional | Implemented via the backend |

```bash
npm install
npm test        # extension + backend tests
npm run check   # syntax check
```

## Multiple Outlook accounts

Run one Chrome profile per Outlook account, each with the extension and its own **Profile name**, all pointing at the
same backend. Rows are handed out first come, first served, never twice, and column L records which account sent each row.

## How it fits together

1. The extension asks the backend for a job at a random time every 10–20 minutes.
2. The backend takes the first row whose status (J) is empty, marks it `pending`, and returns it:
   first address in G = **To**, the rest = **Cc**, H = subject, I = content.
3. The extension composes and sends the email in Outlook Web, verifies it, and reports back.
4. The backend writes `sent` (or `failed` / `needs_review`) to column J.

Neither half has been run against the real Outlook Web or live Google yet; see the checklists in each README.
