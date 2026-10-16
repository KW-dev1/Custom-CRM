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

// ---------------- service account ----------------
import { generateKeyPairSync, createVerify } from 'node:crypto';
import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createServiceAccountAuth, loadServiceAccountKey, SHEETS_SCOPE, TOKEN_URL } from '../src/google-auth.js';

const rsa = generateKeyPairSync('rsa', { modulusLength: 2048 });
const privatePem = rsa.privateKey.export({ type: 'pkcs8', format: 'pem' });
const SA = 'outlook-bot@my-project.iam.gserviceaccount.com';
const decode = (p) => JSON.parse(Buffer.from(p, 'base64url').toString());

test('service account: sends a correctly signed JWT to Google and returns the access token', async () => {
  let seen;
  const auth = createServiceAccountAuth({ clientEmail: SA, privateKey: privatePem, now: () => 1_700_000_000_000,
    fetchImpl: async (url, init) => { seen = { url, body: new URLSearchParams(String(init.body)) }; return json(200, { access_token: 'SA-TOKEN', expires_in: 3600 }); } });
  assert.equal(await auth.getAccessToken(), 'SA-TOKEN');
  assert.equal(auth.identity, SA);
  assert.equal(seen.url, TOKEN_URL);
  assert.equal(seen.body.get('grant_type'), 'urn:ietf:params:oauth:grant-type:jwt-bearer');
  const [h, c, sig] = seen.body.get('assertion').split('.');
  assert.deepEqual(decode(h), { alg: 'RS256', typ: 'JWT' });
  assert.deepEqual(decode(c), { iss: SA, scope: SHEETS_SCOPE, aud: TOKEN_URL, iat: 1_700_000_000, exp: 1_700_003_600 });
  const ok = createVerify('RSA-SHA256').update(`${h}.${c}`).verify(rsa.publicKey, Buffer.from(sig, 'base64url'));
  assert.equal(ok, true, 'the signature must verify with the matching public key');
});
test('service account: token is cached, refreshed after expiry, concurrent callers share one request', async () => {
  let calls = 0; let t = 0;
  const auth = createServiceAccountAuth({ clientEmail: SA, privateKey: privatePem, now: () => t,
    fetchImpl: async () => { calls++; return json(200, { access_token: `t${calls}`, expires_in: 3600 }); } });
  await Promise.all([auth.getAccessToken(), auth.getAccessToken(), auth.getAccessToken()]);
  assert.equal(calls, 1);
  assert.equal(await auth.getAccessToken(), 't1');
  t = 3600 * 1000;
  assert.equal(await auth.getAccessToken(), 't2');
});
test('service account: invalid_grant gives an actionable error without leaking the key', async () => {
  const auth = createServiceAccountAuth({ clientEmail: SA, privateKey: privatePem, fetchImpl: async () => json(400, { error: 'invalid_grant', error_description: 'Invalid JWT Signature.' }) });
  await assert.rejects(auth.getAccessToken(), (e) => e instanceof GoogleAuthError && e.revoked && /clock|deleted/.test(e.message) && !e.message.includes('PRIVATE KEY'));
});
test('service account: unusable private key and network failure are reported cleanly', async () => {
  const bad = createServiceAccountAuth({ clientEmail: SA, privateKey: 'not a key', fetchImpl: async () => json(200, {}) });
  await assert.rejects(bad.getAccessToken(), (e) => /private key/.test(e.message) && !e.message.includes('not a key'));
  const down = createServiceAccountAuth({ clientEmail: SA, privateKey: privatePem, fetchImpl: async () => { throw new Error('x'); } });
  await assert.rejects(down.getAccessToken(), /Could not reach Google/);
});
test('loadServiceAccountKey: valid key file', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'sa-'));
  const f = join(dir, 'k.json');
  await writeFile(f, JSON.stringify({ type: 'service_account', client_email: SA, private_key: privatePem, project_id: 'p' }));
  assert.deepEqual(await loadServiceAccountKey(f), { clientEmail: SA, privateKey: privatePem });
});
test('loadServiceAccountKey: missing file, bad JSON, wrong kind of file, missing fields', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'sa-'));
  await assert.rejects(loadServiceAccountKey(join(dir, 'nope.json')), /not found/);
  const bad = join(dir, 'bad.json'); await writeFile(bad, '{oops');
  await assert.rejects(loadServiceAccountKey(bad), /not valid JSON/);
  const oauth = join(dir, 'oauth.json'); await writeFile(oauth, JSON.stringify({ installed: { client_id: 'x', client_secret: 'SECRET-VALUE' } }));
  await assert.rejects(loadServiceAccountKey(oauth), (e) => /not a service account key/.test(e.message) && !e.message.includes('SECRET-VALUE'));
  const partial = join(dir, 'partial.json'); await writeFile(partial, JSON.stringify({ type: 'service_account', client_email: SA }));
  await assert.rejects(loadServiceAccountKey(partial), /not a service account key/);
});
test('sheets: a 403 tells you exactly which email to share the sheet with', async () => {
  const auth = { getAccessToken: async () => 'T', invalidate() {}, identity: SA };
  const client = createSheetsClient({ cfg: cfg(), auth, sleep: async () => {}, fetchImpl: async () => json(403, { error: { message: 'The caller does not have permission' } }) });
  await assert.rejects(client.readRows(), (e) => e.message.includes(SA) && /Editor/.test(e.message));
});

test('sheets: readHeader reads the header row cells of the configured columns', async () => {
  const { client, seen } = sheetsWith(async (r) => r.url.includes('fields=') ? json(200, meta) : json(200, { values: [['Email', 'Subject', 'Message', 'Status']] }));
  const h = await client.readHeader();
  assert.ok(seen.at(-1).url.includes('!G1:J1'));
  assert.deepEqual([h.emails, h.subject, h.body, h.status], ['Email', 'Subject', 'Message', 'Status']);
});

// ---- invalid_grant explanations ----
import { explainInvalidGrant } from '../src/google-auth.js';

test('explainInvalidGrant: a measured clock skew is reported in minutes, ahead or behind', () => {
  assert.match(explainInvalidGrant('anything', 600), /about 10 minute\(s\) ahead of Google/);
  assert.match(explainInvalidGrant('', -300), /about 5 minute\(s\) behind Google/);
  assert.match(explainInvalidGrant('', 600), /Sync the Windows clock/);
});
test('explainInvalidGrant: small skew falls through to what Google said', () => {
  assert.match(explainInvalidGrant('Invalid JWT Signature.', 3), /does not match Google's records.*NEW JSON key/s);
  assert.match(explainInvalidGrant('Invalid JWT: Token must be a short-lived token (60 minutes) and in a reasonable timeframe.', null), /clock is wrong/);
  assert.match(explainInvalidGrant('Invalid grant: account not found', null), /not found or is disabled/);
  assert.match(explainInvalidGrant('', null), /deleted or disabled.*clock/s);
  assert.match(explainInvalidGrant('Invalid JWT Signature.', 3), /Google says: "Invalid JWT Signature\."/);
});
test('service account: the real error path uses the Date header to detect a bad clock', async () => {
  const googleNow = Date.parse('2026-10-16T12:00:00Z');
  const auth = createServiceAccountAuth({ clientEmail: SA, privateKey: privatePem, now: () => googleNow + 25 * 60 * 1000,
    fetchImpl: async () => ({ ok: false, status: 400, headers: { get: (k) => (k === 'date' ? new Date(googleNow).toUTCString() : null) }, json: async () => ({ error: 'invalid_grant', error_description: 'Invalid JWT: Token must be a short-lived token' }) }) });
  await assert.rejects(auth.getAccessToken(), (e) => e.revoked && /about 25 minute\(s\) ahead/.test(e.message));
});
