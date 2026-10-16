// Google login for the Sheets API, two ways (plain fetch + node:crypto, no Google SDK):
//  * service account (recommended): a robot account with a JSON key; you share the sheet with its email;
//  * OAuth2 refresh token: acts as your own Google user.
// Both expose { getAccessToken(), invalidate(), identity }.

import { createSign } from 'node:crypto';
import { readFile } from 'node:fs/promises';

export class GoogleAuthError extends Error {
  constructor(message, { revoked = false, status = null } = {}) {
    super(message);
    this.name = 'GoogleAuthError';
    this.revoked = revoked;
    this.status = status;
  }
}

export const TOKEN_URL = 'https://oauth2.googleapis.com/token';
export const SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets';

export function createGoogleAuth({ clientId, clientSecret, refreshToken, fetchImpl = globalThis.fetch, now = Date.now }) {
  let cached = null; // { token, expiresAt }
  let inflight = null;

  async function refresh() {
    let res;
    try {
      res = await fetchImpl(TOKEN_URL, {
        method: 'POST',
        headers: { 'content-type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({ client_id: clientId, client_secret: clientSecret, refresh_token: refreshToken, grant_type: 'refresh_token' }),
      });
    } catch {
      throw new GoogleAuthError('Could not reach Google to refresh the access token.');
    }
    let data = null;
    try { data = await res.json(); } catch { /* handled below */ }
    if (!res.ok || !data || !data.access_token) {
      const code = data && data.error;
      if (code === 'invalid_grant') {
        throw new GoogleAuthError('Google rejected the refresh token (expired or revoked). Re-run `npm run auth --workspace backend`. If your OAuth app is in "Testing" mode, tokens expire after 7 days: publish the app.', { revoked: true, status: res.status });
      }
      throw new GoogleAuthError(`Google token endpoint error${code ? `: ${code}` : ''} (HTTP ${res.status}).`, { status: res.status });
    }
    cached = { token: data.access_token, expiresAt: now() + Math.max(0, (Number(data.expires_in) || 3600) - 60) * 1000 };
    return cached.token;
  }

  return {
    identity: 'the Google account you authorised',
    async getAccessToken() {
      if (cached && now() < cached.expiresAt) return cached.token;
      inflight = inflight || refresh().finally(() => { inflight = null; });
      return inflight;
    },
    invalidate() { cached = null; },
  };
}

// ---------------------------------------------------------------------------------------------
// Service account
// ---------------------------------------------------------------------------------------------

/** Read and validate the JSON key downloaded from Google Cloud. Error messages never include the key. */
export async function loadServiceAccountKey(path) {
  let text;
  try { text = await readFile(path, 'utf8'); }
  catch { throw new GoogleAuthError(`Service account key file not found or unreadable: ${path}`); }
  let key;
  try { key = JSON.parse(text); }
  catch { throw new GoogleAuthError('The service account key file is not valid JSON. Download a fresh JSON key from Google Cloud (Service Accounts > Keys > Add key > JSON).'); }
  if (!key || key.type !== 'service_account' || !key.client_email || !key.private_key) {
    throw new GoogleAuthError('That file is not a service account key: it needs "type": "service_account", "client_email" and "private_key". (An OAuth client file, with "installed" or "web" in it, is the wrong download.)');
  }
  return { clientEmail: String(key.client_email), privateKey: String(key.private_key) };
}

const b64u = (v) => Buffer.from(v).toString('base64url');

export function createServiceAccountAuth({ clientEmail, privateKey, fetchImpl = globalThis.fetch, now = Date.now }) {
  let cached = null;
  let inflight = null;

  function assertion() {
    const iat = Math.floor(now() / 1000);
    const head = b64u(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
    const claims = b64u(JSON.stringify({ iss: clientEmail, scope: SHEETS_SCOPE, aud: TOKEN_URL, iat, exp: iat + 3600 }));
    try {
      const sig = createSign('RSA-SHA256').update(`${head}.${claims}`).sign(privateKey);
      return `${head}.${claims}.${b64u(sig)}`;
    } catch {
      throw new GoogleAuthError('The private key in the service account file could not be used to sign. Download a fresh JSON key.');
    }
  }

  async function refresh() {
    let res;
    try {
      res = await fetchImpl(TOKEN_URL, {
        method: 'POST',
        headers: { 'content-type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({ grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: assertion() }),
      });
    } catch (e) {
      if (e instanceof GoogleAuthError) throw e;
      throw new GoogleAuthError('Could not reach Google to get an access token.');
    }
    let data = null;
    try { data = await res.json(); } catch { /* handled below */ }
    if (!res.ok || !data || !data.access_token) {
      const code = data && data.error;
      if (code === 'invalid_grant') {
        throw new GoogleAuthError('Google rejected the service account key. It may have been deleted or disabled in Google Cloud, or this computer\'s clock is wrong. Create a new JSON key and check the system time.', { revoked: true, status: res.status });
      }
      throw new GoogleAuthError(`Google token endpoint error${code ? `: ${code}` : ''} (HTTP ${res.status}).`, { status: res.status });
    }
    cached = { token: data.access_token, expiresAt: now() + Math.max(0, (Number(data.expires_in) || 3600) - 60) * 1000 };
    return cached.token;
  }

  return {
    identity: clientEmail,
    async getAccessToken() {
      if (cached && now() < cached.expiresAt) return cached.token;
      inflight = inflight || refresh().finally(() => { inflight = null; });
      return inflight;
    },
    invalidate() { cached = null; },
  };
}
