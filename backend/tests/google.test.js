import test from 'node:test';
import assert from 'node:assert/strict';
import { createGoogleAuth, GoogleAuthError } from '../src/google-auth.js';
import { createSheetsClient, SheetsError } from '../src/sheets-client.js';
import { loadConfig } from '../src/config.js';

const cfg = (extra = {}) => loadConfig({
  COL_SENDER: '', ...extra,
  DATA_SOURCE: 'google', API_KEYS: 'abcdefghijklmnop1234', SHEET_URL: 'https://docs.google.com/spreadsheets/d/SHEETID/edit?gid=42',
  GOOGLE_CLIENT_ID: 'cid', GOOGLE_CLIENT_SECRET: 'sec', GOOGLE_REFRESH_TOKEN: 'rt',
});
const json = (status, body) => ({ ok: status >= 200 && status < 300, status, json: async () => body });

test('auth: exchanges refresh token, caches until near expiry, refreshes after', async () => {
  let calls = 0; let t = 0;
  const auth = createGoogleAuth({ clientId: 'c', clientSecret: 's', refreshToken: 'r', now: () => t,
    fetchImpl: async (url, init) => { calls++; assert.match(String(init.body), /grant_type=refresh_token/); return json(200, { access_token: `tok${calls}`, expires_in: 3600 }); } });
  assert.equal(await auth.getAccessToken(), 'tok1');
  assert.equal(await auth.getAccessToken(), 'tok1');
  t = 3600 * 1000;
  assert.equal(await auth.getAccessToken(), 'tok2');
  assert.equal(calls, 2);
});
test('auth: concurrent callers share one refresh', async () => {
  let calls = 0;
  const auth = createGoogleAuth({ clientId: 'c', clientSecret: 's', refreshToken: 'r', fetchImpl: async () => { calls++; return json(200, { access_token: 'x', expires_in: 3600 }); } });
  await Promise.all([auth.getAccessToken(), auth.getAccessToken(), auth.getAccessToken()]);
  assert.equal(calls, 1);
});
test('auth: invalid_grant gives an actionable, token-free error', async () => {
  const auth = createGoogleAuth({ clientId: 'c', clientSecret: 's', refreshToken: 'SECRET-RT', fetchImpl: async () => json(400, { error: 'invalid_grant' }) });
  await assert.rejects(auth.getAccessToken(), (e) => e instanceof GoogleAuthError && e.revoked && /Testing/.test(e.message) && !e.message.includes('SECRET-RT'));
});
test('auth: network failure', async () => {
  const auth = createGoogleAuth({ clientId: 'c', clientSecret: 's', refreshToken: 'r', fetchImpl: async () => { throw new Error('x'); } });
  await assert.rejects(auth.getAccessToken(), GoogleAuthError);
});

function sheetsWith(handler, env = {}) {
  const seen = [];
  const auth = { getAccessToken: async () => 'T', invalidate() { seen.push('invalidate'); } };
  const client = createSheetsClient({ cfg: cfg(env), auth, sleep: async () => {}, fetchImpl: async (url, init = {}) => { seen.push({ url: decodeURIComponent(url), method: init.method || 'GET', body: init.body && JSON.parse(init.body), auth: init.headers.authorization }); return handler(seen.at(-1), seen); } });
  return { client, seen };
}
const meta = { sheets: [{ properties: { sheetId: 0, title: 'Other' } }, { properties: { sheetId: 42, title: "Bob's Leads" } }] };

test('sheets: resolves the tab title from the gid, quotes it, reads G:J after the header', async () => {
  const { client, seen } = sheetsWith(async (r) => r.url.includes('fields=') ? json(200, meta) : json(200, { values: [['a@x.com', 'S1', 'B1', ''], [], ['', '', '', ''], ['c@x.com', 'S3', 'B3']] }));
  const rows = await client.readRows();
  assert.ok(seen[1].url.includes("'Bob''s Leads'!G2:J"));
  assert.equal(seen[1].auth, 'Bearer T');
  assert.deepEqual(rows.map((r) => r.rowNumber), [2, 3, 4, 5], 'blank rows keep numbering correct');
  assert.equal(rows[0].emails, 'a@x.com');
  assert.equal(rows[3].status, '', 'short rows are padded');
  assert.equal(rows[1].subject, '');
});
test('sheets: unknown gid is a clear error', async () => {
  const { client } = sheetsWith(async () => json(200, { sheets: [] }));
  await assert.rejects(client.readRows(), /gid 42/);
});
test('sheets: readRow targets exactly one row; refuses header rows', async () => {
  const { client, seen } = sheetsWith(async (r) => r.url.includes('fields=') ? json(200, meta) : json(200, { values: [['a@x.com', 's', 'b', 'pending']] }));
  const row = await client.readRow(7);
  assert.ok(seen[1].url.includes('!G7:J7'));
  assert.equal(row.status, 'pending');
  await assert.rejects(client.readRow(1), SheetsError);
});
test('sheets: writeStatus uses RAW input and only the status cell (plus note if configured)', async () => {
  const { client, seen } = sheetsWith(async (r) => r.url.includes('fields=') ? json(200, meta) : json(200, {}));
  await client.writeStatus(9, '=HYPERLINK("x")');
  const w = seen.at(-1);
  assert.equal(w.method, 'POST');
  assert.equal(w.body.valueInputOption, 'RAW');
  assert.deepEqual(w.body.data, [{ range: "'Bob''s Leads'!J9", values: [['=HYPERLINK("x")']] }]);
  await assert.rejects(client.writeStatus(1, 'x'), /header/);
});
test('sheets: 401 triggers one token refresh and retry; 403/404 give hints', async () => {
  let n = 0;
  const { client, seen } = sheetsWith(async (r) => r.url.includes('fields=') ? (n++ === 0 ? json(401, {}) : json(200, meta)) : json(200, { values: [] }));
  await client.readRows();
  assert.ok(seen.includes('invalidate'));
  const f = sheetsWith(async () => json(403, { error: { message: 'denied' } }));
  await assert.rejects(f.client.readRows(), /edit access/);
});
test('sheets: transient GET errors are retried, writes are not', async () => {
  let g = 0;
  const a = sheetsWith(async (r) => r.url.includes('fields=') ? json(200, meta) : (g++ < 2 ? json(503, {}) : json(200, { values: [] })));
  await a.client.readRows();
  assert.equal(g, 3);
  let w = 0;
  const b = sheetsWith(async (r) => r.url.includes('fields=') ? json(200, meta) : (w++, json(503, {})));
  await assert.rejects(b.client.writeStatus(3, 'sent'), (e) => e.retryable);
  assert.equal(w, 1);
});

test('sheets: sender is written in the same batch as the status and read from its column', async () => {
  const { client, seen } = sheetsWith(async (r) => r.url.includes('fields=') ? json(200, meta) : json(200, { values: [['a@x.com', 's', 'b', 'pending', '', 'outlook-A']] }), { COL_SENDER: 'L' });
  await client.writeStatus(9, 'pending', undefined, { sender: 'outlook-A' });
  const w = seen.at(-1);
  assert.deepEqual(w.body.data.map((d) => d.range), ["'Bob''s Leads'!J9", "'Bob''s Leads'!L9"]);
  const row = await client.readRow(9);
  assert.equal(row.sender, 'outlook-A');
});
