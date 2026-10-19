// Settings defaults, normalisation and validation. Pure module (no chrome.* calls).

export const OUTLOOK_MAIL_PREFIX = 'https://outlook.live.com/mail/';
export const OUTLOOK_TAB_PATTERN = 'https://outlook.live.com/mail/*';

// Mail providers the extension can drive. Each Chrome profile uses ONE provider and one account.
export const PROVIDERS = Object.freeze({
  outlook: Object.freeze({ label: 'Outlook Web', openUrl: 'outlook.live.com/mail', tabPattern: OUTLOOK_TAB_PATTERN }),
  yahoo: Object.freeze({ label: 'Yahoo Mail', openUrl: 'mail.yahoo.com', tabPattern: 'https://mail.yahoo.com/*', needsFocus: true }),
});
export const DEFAULT_PROVIDER = 'outlook';

// Canonical internal statuses. Mapped to the backend vocabulary via settings.statusValues.
export const STATUS = Object.freeze({
  PROCESSING: 'processing',
  SENT: 'sent',
  FAILED: 'failed',
  NEEDS_REVIEW: 'needs_review',
});

// Default backend vocabulary: canonical status -> value sent to the backend (overridable).
export const DEFAULT_STATUS_VALUES = Object.freeze(Object.fromEntries(Object.values(STATUS).map((v) => [v, v])));

// Statuses shown in the popup.
export const UI_STATUS = Object.freeze({
  NOT_CONFIGURED: 'not_configured',
  WAITING_OUTLOOK: 'waiting_outlook',
  READY: 'ready',
  PROCESSING: 'processing',
  SENT: 'sent',
  FAILED: 'failed',
  NEEDS_REVIEW: 'needs_review',
  PAUSED: 'paused',
});

// PLACEHOLDERS: the real backend contract was not available when this was written.
// Every path/method/header below is configurable in the popup (Advanced).
export const DEFAULT_SETTINGS = Object.freeze({
  apiBaseUrl: '',
  apiKey: '',
  provider: DEFAULT_PROVIDER, // 'outlook' | 'yahoo': which webmail this Chrome profile drives
  profileName: '', // names this Chrome profile / Outlook account to the backend (X-Client-Id); empty = "default"
  authHeader: 'Authorization',
  authScheme: 'Bearer', // set to '' to send the raw key in the header
  claimPath: '/api/email-jobs/claim',
  claimMethod: 'POST',
  statusPath: '/api/email-jobs/{id}/status',
  statusMethod: 'POST',
  jobStatePath: '', // optional, e.g. /api/email-jobs/{id}; empty = skip the pre-send state check
  statusValues: { ...DEFAULT_STATUS_VALUES },
  fetchMinMinutes: 10, // the backend is asked for one job at a random time between min and max minutes
  fetchMaxMinutes: 20,
  pauseAfterProblems: 3, // automation pauses only after this many failed / needs-review results IN A ROW (a sent email resets the count)
  verifySentItems: true, // after Send, confirm the message is listed in Sent Items (stricter, ~10 s per email)
  enabled: false,
});

export const PROFILE_NAME_RE = /^[A-Za-z0-9 ._-]{1,40}$/;

export function normalizeSettings(raw = {}) {
  const s = { ...DEFAULT_SETTINGS, ...raw };
  for (const k of ['apiBaseUrl', 'apiKey', 'profileName', 'authHeader', 'authScheme', 'claimPath', 'statusPath', 'jobStatePath']) {
    s[k] = typeof s[k] === 'string' ? s[k].trim() : '';
  }
  s.apiBaseUrl = s.apiBaseUrl.replace(/\/+$/, '');
  s.provider = typeof s.provider === 'string' && s.provider.trim() ? s.provider.trim().toLowerCase() : DEFAULT_PROVIDER;
  s.claimMethod = String(s.claimMethod || 'POST').toUpperCase();
  s.statusMethod = String(s.statusMethod || 'POST').toUpperCase();
  s.statusValues = { ...DEFAULT_STATUS_VALUES, ...(raw.statusValues || {}) };
  s.fetchMinMinutes = Number(s.fetchMinMinutes);
  s.fetchMaxMinutes = Number(s.fetchMaxMinutes);
  s.enabled = Boolean(s.enabled);
  s.verifySentItems = s.verifySentItems !== false;
  s.pauseAfterProblems = Number(s.pauseAfterProblems);
  return s;
}

export function isAllowedBackendUrl(u) {
  if (u.username || u.password) return false;
  if (u.protocol === 'https:') return true;
  // Plain HTTP is only tolerated for local development.
  return u.protocol === 'http:' && ['localhost', '127.0.0.1'].includes(u.hostname);
}

/** @returns {{ok: boolean, errors: Record<string,string>}} */
export function validateSettings(input) {
  const s = normalizeSettings(input);
  const errors = {};

  if (!s.apiBaseUrl) errors.apiBaseUrl = 'Backend API base URL is required.';
  else {
    try {
      const u = new URL(s.apiBaseUrl);
      if (!isAllowedBackendUrl(u)) errors.apiBaseUrl = 'Use an https:// URL without embedded credentials.';
    } catch {
      errors.apiBaseUrl = 'Backend API base URL is not a valid URL.';
    }
  }
  if (!s.apiKey) errors.apiKey = 'API key / token is required.';
  if (!Object.hasOwn(PROVIDERS, s.provider)) errors.provider = `Mail provider must be one of: ${Object.keys(PROVIDERS).join(', ')}.`;
  if (s.profileName && !PROFILE_NAME_RE.test(s.profileName)) errors.profileName = 'Profile name: 1-40 letters, digits, spaces, dots, dashes or underscores (no email addresses).';
  if (s.authHeader && !/^[A-Za-z0-9-]+$/.test(s.authHeader)) errors.authHeader = 'Invalid header name.';
  if (!s.authHeader) errors.authHeader = 'Auth header name is required.';
  if (!s.claimPath.startsWith('/')) errors.claimPath = 'Claim path must start with "/".';
  if (!s.statusPath.startsWith('/')) errors.statusPath = 'Status path must start with "/".';
  else if (!s.statusPath.includes('{id}')) errors.statusPath = 'Status path must contain the {id} placeholder.';
  if (s.jobStatePath && (!s.jobStatePath.startsWith('/') || !s.jobStatePath.includes('{id}'))) {
    errors.jobStatePath = 'Job state path must start with "/" and contain {id}.';
  }
  if (!['GET', 'POST'].includes(s.claimMethod)) errors.claimMethod = 'Claim method must be GET or POST.';
  if (!['POST', 'PUT', 'PATCH'].includes(s.statusMethod)) errors.statusMethod = 'Status method must be POST, PUT or PATCH.';
  const okMin = Number.isFinite(s.fetchMinMinutes) && s.fetchMinMinutes >= 1 && s.fetchMinMinutes <= 1440;
  const okMax = Number.isFinite(s.fetchMaxMinutes) && s.fetchMaxMinutes >= 1 && s.fetchMaxMinutes <= 1440;
  if (!okMin) errors.fetchMinMinutes = 'Minimum must be between 1 and 1440 minutes.';
  if (!okMax) errors.fetchMaxMinutes = 'Maximum must be between 1 and 1440 minutes.';
  if (okMin && okMax && s.fetchMinMinutes > s.fetchMaxMinutes) errors.fetchMaxMinutes = 'Maximum must be at least the minimum.';
  if (!Number.isInteger(s.pauseAfterProblems) || s.pauseAfterProblems < 1 || s.pauseAfterProblems > 20) {
    errors.pauseAfterProblems = 'Pause-after count must be a whole number from 1 to 20.';
  }
  for (const k of Object.values(STATUS)) {
    const v = s.statusValues[k];
    if (typeof v !== 'string' || !v.trim()) errors.statusValues = `Status value for "${k}" must be a non-empty string.`;
  }
  return { ok: Object.keys(errors).length === 0, errors };
}

/** Origin pattern for chrome.permissions.request, or null if the URL is unusable. */
export function backendOriginPattern(apiBaseUrl) {
  try {
    return `${new URL(apiBaseUrl).origin}/*`;
  } catch {
    return null;
  }
}
