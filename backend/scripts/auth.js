// One-time helper: obtain a Google OAuth refresh token for the Sheets scope.
// Run on a machine with a browser:  npm run auth --workspace backend
// Needs GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET (OAuth client of type "Desktop app").
import http from 'node:http';
import { randomBytes, createHash } from 'node:crypto';
import { TOKEN_URL, SHEETS_SCOPE } from '../src/google-auth.js';

const { GOOGLE_CLIENT_ID: clientId, GOOGLE_CLIENT_SECRET: clientSecret } = process.env;
if (!clientId || !clientSecret) {
  console.error('Set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET (in backend/.env or the environment) first.');
  process.exit(1);
}
const PORT = 53682;
const redirect = `http://127.0.0.1:${PORT}/callback`;
const verifier = randomBytes(32).toString('base64url');
const challenge = createHash('sha256').update(verifier).digest('base64url');
const state = randomBytes(16).toString('hex');

const url = new URL('https://accounts.google.com/o/oauth2/v2/auth');
url.search = new URLSearchParams({
  client_id: clientId, redirect_uri: redirect, response_type: 'code', scope: SHEETS_SCOPE,
  access_type: 'offline', prompt: 'consent', state, code_challenge: challenge, code_challenge_method: 'S256',
}).toString();

const server = http.createServer(async (req, res) => {
  const u = new URL(req.url, redirect);
  if (u.pathname !== '/callback') { res.statusCode = 404; return res.end(); }
  const done = (msg, code = 0) => { res.end(msg); server.close(); setTimeout(() => process.exit(code), 100); };
  if (u.searchParams.get('state') !== state) return done('State mismatch. Aborted.', 1);
  const code = u.searchParams.get('code');
  if (!code) return done(`Authorization failed: ${u.searchParams.get('error') || 'no code'}`, 1);
  const r = await fetch(TOKEN_URL, {
    method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ code, client_id: clientId, client_secret: clientSecret, redirect_uri: redirect, grant_type: 'authorization_code', code_verifier: verifier }),
  });
  const data = await r.json().catch(() => ({}));
  if (!data.refresh_token) {
    console.error('No refresh token returned:', data.error_description || data.error || r.status);
    return done('No refresh token returned. See the terminal.', 1);
  }
  console.log('\nAdd this to backend/.env (keep it secret):\n\nGOOGLE_REFRESH_TOKEN=' + data.refresh_token + '\n');
  done('Done. You can close this tab and return to the terminal.');
});
server.listen(PORT, '127.0.0.1', () => {
  console.log('Open this URL in your browser and approve access to Google Sheets:\n\n' + url + '\n');
});
