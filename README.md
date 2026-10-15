# Custom CRM

A custom system, built in parts:

| Part | Stack | Status |
|---|---|---|
| Chrome extension that sends queued emails through Outlook Web | MV3, plain JS | Implemented: see [`extension/README.md`](extension/README.md) |
| Backend API | Node.js | Not started |
| Database | Google Sheets | Not started |

```bash
npm install
npm test        # extension unit + DOM tests
npm run check   # syntax check
```
