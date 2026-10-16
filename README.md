# Custom CRM

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

## How it fits together

1. The extension asks the backend for a job at a random time every 10–20 minutes.
2. The backend takes the first row whose status (J) is empty, marks it `pending`, and returns it:
   first address in G = **To**, the rest = **Cc**, H = subject, I = content.
3. The extension composes and sends the email in Outlook Web, verifies it, and reports back.
4. The backend writes `sent` (or `failed` / `needs_review`) to column J.

Neither half has been run against the real Outlook Web or live Google yet; see the checklists in each README.
