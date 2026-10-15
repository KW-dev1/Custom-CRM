// OAuth2 refresh-token flow (acts as your Google user). Uses plain fetch; no Google SDK.

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
    async getAccessToken() {
      if (cached && now() < cached.expiresAt) return cached.token;
      inflight = inflight || refresh().finally(() => { inflight = null; });
      return inflight;
    },
    invalidate() { cached = null; },
  };
}
