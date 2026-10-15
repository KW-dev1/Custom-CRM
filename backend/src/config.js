// Environment-based configuration. Secrets only come from the environment, never from source.

export class ConfigError extends Error {
  constructor(errors) {
    super(`Invalid configuration:\n - ${errors.join('\n - ')}`);
    this.name = 'ConfigError';
    this.errors = errors;
  }
}

const COL_RE = /^[A-Z]{1,2}$/;
export const colIndex = (letters) => [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;

/** Accepts a full sheet URL and returns { id, gid } (either may be null). */
export function parseSheetUrl(url) {
  const m = /\/spreadsheets\/d\/([A-Za-z0-9_-]+)/.exec(url || '');
  const g = /[#&?]gid=(\d+)/.exec(url || '');
  return { id: m ? m[1] : null, gid: g ? g[1] : null };
}

export function loadConfig(env = process.env) {
  const errors = [];
  const get = (k, d = '') => String(env[k] ?? d).trim();
  const need = (k) => { const v = get(k); if (!v) errors.push(`${k} is required`); return v; };
  const int = (k, d, min, max) => {
    const raw = get(k, String(d));
    const n = Number(raw);
    if (!Number.isInteger(n) || n < min || n > max) { errors.push(`${k} must be an integer between ${min} and ${max}`); return d; }
    return n;
  };
  const col = (k, d, optional = false) => {
    const v = get(k, d).toUpperCase();
    if (!v && optional) return null;
    if (!COL_RE.test(v)) { errors.push(`${k} must be a column letter (A-ZZ)`); return d; }
    return v;
  };

  const keys = (get('API_KEYS') || get('API_KEY')).split(',').map((s) => s.trim()).filter(Boolean);
  if (!keys.length) errors.push('API_KEYS (or API_KEY) is required');
  if (keys.some((k) => k.length < 16)) errors.push('Every API key must be at least 16 characters');

  const fromUrl = parseSheetUrl(get('SHEET_URL'));
  const sheetId = get('SHEET_ID') || fromUrl.id;
  const gid = get('SHEET_GID') || fromUrl.gid;
  const tab = get('SHEET_TAB');
  if (!sheetId) errors.push('SHEET_URL or SHEET_ID is required');
  else if (!/^[A-Za-z0-9_-]+$/.test(sheetId)) errors.push('SHEET_ID has invalid characters');
  if (!gid && !tab) errors.push('SHEET_GID (or a SHEET_URL containing gid=) or SHEET_TAB is required');
  if (gid && !/^\d+$/.test(gid)) errors.push('SHEET_GID must be numeric');

  const columns = {
    emails: col('COL_EMAILS', 'G'), subject: col('COL_SUBJECT', 'H'), body: col('COL_BODY', 'I'),
    status: col('COL_STATUS', 'J'), note: col('COL_NOTE', '', true),
  };
  const used = Object.values(columns).filter(Boolean);
  if (new Set(used).size !== used.length) errors.push('Column settings must all be different');

  const cfg = {
    port: int('PORT', 3000, 1, 65535),
    host: get('HOST', '127.0.0.1'),
    apiKeys: keys,
    google: { clientId: need('GOOGLE_CLIENT_ID'), clientSecret: need('GOOGLE_CLIENT_SECRET'), refreshToken: need('GOOGLE_REFRESH_TOKEN') },
    sheet: { id: sheetId, gid: gid || null, tab: tab || null, headerRows: int('HEADER_ROWS', 1, 0, 100) },
    columns,
    // Defence in depth: refuse job claims closer together than this, whatever the client does.
    minClaimIntervalSec: int('MIN_CLAIM_INTERVAL_SEC', 300, 0, 86400),
    maxCcRecipients: int('MAX_CC', 20, 0, 100),
    statusLabels: {
      pending: get('STATUS_PENDING', 'pending'), sent: get('STATUS_SENT', 'sent'),
      failed: get('STATUS_FAILED', 'failed'), needs_review: get('STATUS_NEEDS_REVIEW', 'needs_review'),
    },
  };
  if (new Set(Object.values(cfg.statusLabels).map((s) => s.toLowerCase())).size !== 4 || Object.values(cfg.statusLabels).some((s) => !s)) {
    errors.push('Status labels must be four distinct non-empty values');
  }
  if (errors.length) throw new ConfigError(errors);
  return cfg;
}
