// Read-only connection check:  npm run check:source
// Verifies the login, finds the sheet/file, shows which header sits in each configured column, and
// summarises the rows (counts only, no personal data). It never writes anything.
import { loadConfig, ConfigError } from '../src/config.js';
import { createGoogleAuth, createServiceAccountAuth, loadServiceAccountKey, GoogleAuthError } from '../src/google-auth.js';
import { createSheetsClient } from '../src/sheets-client.js';
import { createXlsxStore } from '../src/xlsx-store.js';
import { summarizeRows } from '../src/inspect.js';

const ok = (m) => console.log(`  [ok]   ${m}`);
const warn = (m) => console.log(`  [warn] ${m}`);
const fail = (m) => { console.error(`  [FAIL] ${m}`); process.exitCode = 1; };

let cfg;
try { cfg = loadConfig(); }
catch (e) {
  if (e instanceof ConfigError) { console.error(e.message); process.exit(1); }
  throw e;
}

console.log(`\nData source: ${cfg.source.type === 'xlsx' ? `local Excel file (${cfg.source.path})` : 'Google Sheets'}`);
let store;
try {
  if (cfg.source.type === 'google') {
    const auth = cfg.google.serviceAccountFile
      ? createServiceAccountAuth(await loadServiceAccountKey(cfg.google.serviceAccountFile))
      : createGoogleAuth(cfg.google);
    await auth.getAccessToken();
    ok(cfg.google.serviceAccountFile
      ? `Google login works (service account ${auth.identity}). The sheet must be shared with this address as Editor.`
      : 'Google login works (OAuth user).');
    store = createSheetsClient({ cfg, auth });
  } else {
    store = createXlsxStore({ cfg });
  }

  const header = await store.readHeader();
  ok('Found the sheet/file and the right tab.');
  console.log('\n  Column mapping (check the header names make sense):');
  for (const [k, col] of Object.entries(cfg.columns)) {
    if (!col) continue;
    const h = header[k];
    console.log(`    ${col}  ${k.padEnd(8)} header: ${h ? `"${h}"` : '(empty)'}`);
  }
  const expect = { emails: /mail/i, subject: /subject/i, body: /message|content|body/i, status: /status/i };
  for (const [k, re] of Object.entries(expect)) {
    const h = header[k];
    if (h !== undefined && h !== '' && !re.test(h)) warn(`Column ${cfg.columns[k]} is "${h}", which does not look like ${k}. Check COL_* settings if this is wrong.`);
  }

  const rows = await store.readRows();
  const s = summarizeRows(rows, cfg);
  console.log('\n  Rows below the header:');
  console.log(`    ${s.ready} ready to send (status empty, address + subject + message filled)`);
  console.log(`    ${s.incomplete} skipped: still being drafted (something missing)`);
  console.log(`    ${s.invalid.length} would be marked failed (malformed data)${s.invalid.length ? `: ${s.invalid.slice(0, 5).map((i) => `row ${i.row}`).join(', ')}` : ''}`);
  console.log(`    ${s.pending} pending, ${s.sent} sent, ${s.failed} failed, ${s.needs_review} needs_review, ${s.other} other status`);
  console.log(`    ${s.blank} blank`);
  if (s.pending) warn(`${s.pending} row(s) are "pending": a profile took them and never reported. Check Sent Items before clearing them.`);
  if (s.ready === 0) warn('Nothing is ready to send yet: fill in address, subject and message, and leave the status cell empty.');
  console.log('\nRead-only check finished. Nothing was changed.\n');
} catch (e) {
  fail(e instanceof GoogleAuthError || e.name === 'SheetsError' ? e.message : `Unexpected error: ${e.message}`);
  console.error('');
}
