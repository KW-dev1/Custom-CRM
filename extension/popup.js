// Popup UI. Uses textContent only (never innerHTML) so job data / errors cannot inject markup.
import { validateSettings, normalizeSettings, backendOriginPattern, DEFAULT_SETTINGS } from './config.js';
import { createStore, KEYS } from './store.js';

const store = createStore(chrome.storage.local);
const $ = (id) => document.getElementById(id);
const FIELDS = ['apiBaseUrl', 'apiKey', 'profileName', 'authHeader', 'authScheme', 'claimMethod', 'claimPath', 'statusMethod', 'statusPath', 'jobStatePath', 'fetchMinMinutes', 'fetchMaxMinutes', 'pauseAfterProblems'];

const LABELS = {
  not_configured: 'Not configured', waiting_outlook: 'Waiting for Outlook', ready: 'Ready',
  processing: 'Processing email', sent: 'Email sent', failed: 'Failed', needs_review: 'Needs review', paused: 'Paused',
};

function fmt(ts) {
  return ts ? new Date(ts).toLocaleString() : '';
}

async function renderStatus() {
  const [status, runtime, settings] = await Promise.all([store.getStatus(), store.getRuntime(), store.getSettings()]);
  let code = status ? status.code : 'not_configured';
  let msg = status ? status.message : 'Enter the backend URL and API key below.';
  if (!validateSettings(settings).ok) code = 'not_configured';
  else if (!settings.enabled || !runtime.running) { code = 'paused'; msg = settings.enabled ? 'Paused. Press Start to process jobs.' : 'Automation is disabled.'; }
  const badge = $('badge');
  badge.textContent = LABELS[code] || code;
  badge.dataset.status = code;
  $('statusMsg').textContent = msg;
  $('enabled').checked = settings.enabled;
  $('start').disabled = !settings.enabled || runtime.running;
  $('stop').disabled = !runtime.running;
}

async function renderResult() {
  const r = await store.getLastResult();
  const el = $('lastResult');
  if (!r) { el.textContent = 'No jobs processed yet.'; el.className = 'card muted'; return; }
  el.className = 'card';
  el.textContent = `${LABELS[r.status] || r.status} — job ${r.jobId} (${fmt(r.at)}). ${r.message}`;
}

async function renderErrors() {
  const errs = (await store.getList(KEYS.errors)).slice(-5).reverse();
  const ul = $('errors');
  ul.replaceChildren();
  if (!errs.length) { const li = document.createElement('li'); li.className = 'muted'; li.textContent = 'None.'; ul.append(li); return; }
  for (const e of errs) {
    const li = document.createElement('li');
    li.textContent = `${fmt(e.at)} · ${e.stage}${e.category ? ` · ${e.category}` : ''}${e.jobId ? ` · job ${e.jobId}` : ''}: ${e.message || 'No details.'}`;
    ul.append(li);
  }
}

async function loadForm() {
  const s = await store.getSettings();
  for (const f of FIELDS) $(f).value = s[f] ?? DEFAULT_SETTINGS[f] ?? '';
  $('statusValues').value = JSON.stringify(s.statusValues, null, 2);
  $('verifySentItems').checked = s.verifySentItems;
}

function readForm() {
  const raw = {};
  for (const f of FIELDS) raw[f] = $(f).value;
  raw.verifySentItems = $('verifySentItems').checked;
  let statusError = null;
  try { raw.statusValues = JSON.parse($('statusValues').value || '{}'); }
  catch { statusError = 'Status values must be valid JSON.'; }
  return { raw, statusError };
}

function showErrors(errors) {
  for (const f of [...FIELDS, 'statusValues']) $(`err-${f}`) && ($(`err-${f}`).textContent = errors[f] || '');
}

$('form').addEventListener('submit', async (ev) => {
  ev.preventDefault();
  const msg = $('saveMsg');
  msg.textContent = '';
  const { raw, statusError } = readForm();
  const current = await store.getSettings();
  const next = normalizeSettings({ ...raw, enabled: current.enabled });
  const { errors } = validateSettings(next);
  if (statusError) errors.statusValues = statusError;
  showErrors(errors);
  if (Object.keys(errors).length) { msg.textContent = 'Fix the highlighted fields.'; return; }

  // Ask only for the configured backend origin (no broad host access).
  const origin = backendOriginPattern(next.apiBaseUrl);
  msg.textContent = 'Checking permission to contact the backend…';
  try {
    if (!(await chrome.permissions.contains({ origins: [origin] })) && !(await chrome.permissions.request({ origins: [origin] }))) {
      msg.textContent = 'Permission to contact the backend was not granted.';
      return;
    }
  } catch (e) { msg.textContent = `Could not request permission: ${e.message}`; return; }

  await store.saveSettings(next);
  msg.textContent = 'Saved.';
  chrome.runtime.sendMessage({ type: 'OEA_TICK' }).catch(() => {});
});

$('enabled').addEventListener('change', async (ev) => {
  const s = await store.getSettings();
  await store.saveSettings({ ...s, enabled: ev.target.checked });
  chrome.runtime.sendMessage({ type: ev.target.checked ? 'OEA_TICK' : 'OEA_STOP' }).catch(() => {});
});
$('clearHistory').addEventListener('click', async () => {
  const msg = $('historyMsg');
  if (!confirm('Clear the list of jobs this extension has already processed?\n\nOnly do this if you checked Outlook\'s Sent Items and want a cleared row to be sent again.')) return;
  try {
    const r = await chrome.runtime.sendMessage({ type: 'OEA_CLEAR_HISTORY' });
    msg.textContent = r && r.ok ? 'Job history cleared.' : `Not cleared: ${(r && r.reason) || 'unknown error'}`;
  } catch (e) { msg.textContent = `Not cleared: ${e.message}`; }
});
$('start').addEventListener('click', () => chrome.runtime.sendMessage({ type: 'OEA_START' }).catch(() => {}));
$('stop').addEventListener('click', () => chrome.runtime.sendMessage({ type: 'OEA_STOP' }).catch(() => {}));

chrome.storage.onChanged.addListener(() => { renderStatus(); renderResult(); renderErrors(); });

loadForm();
renderStatus();
renderResult();
renderErrors();
