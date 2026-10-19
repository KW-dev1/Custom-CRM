// MV3 service worker: wires chrome.* APIs to the orchestrator. No long-lived in-memory state;
// processing is driven by alarms and popup messages.

import { PROVIDERS } from './config.js';
import * as api from './api.js';
import { createStore } from './store.js';
import { createLogger } from './logger.js';
import { createOrchestrator } from './orchestrator.js';

const store = createStore(chrome.storage.local);
let secrets = [];
const refreshSecrets = async () => { secrets = [(await store.getSettings()).apiKey]; };
refreshSecrets();
chrome.storage.onChanged.addListener((c) => { if (c.settings) refreshSecrets(); });

const logger = createLogger({ store, secrets: () => secrets });

const tabs = {
  async findMailTab(provider) {
    const found = await chrome.tabs.query({ url: (PROVIDERS[provider] || PROVIDERS.outlook).tabPattern });
    if (!found.length) return null;
    // Never open extra tabs; prefer the active one, otherwise the most recently used.
    return found.find((t) => t.active) || found.sort((a, b) => (b.lastAccessed || 0) - (a.lastAccessed || 0))[0];
  },
  send: (tabId, msg) => chrome.tabs.sendMessage(tabId, msg),
};

const orchestrator = createOrchestrator({
  store, api, tabs, logger,
  // Chrome >= 120 allows 30s alarms; shorter requests are clamped.
  scheduleNext: (sec) => chrome.alarms.create('oea-next', { when: Date.now() + Math.max(sec, 30) * 1000 }),
  // Pending extension API calls keep the worker alive while a long send is in flight.
  keepAlive: () => {
    const id = setInterval(() => chrome.runtime.getPlatformInfo(), 20000);
    return () => clearInterval(id);
  },
});

function ensurePollAlarm() {
  chrome.alarms.create('oea-poll', { periodInMinutes: 1 });
}

chrome.runtime.onInstalled.addListener(() => { ensurePollAlarm(); orchestrator.tick(); });
chrome.runtime.onStartup.addListener(() => { ensurePollAlarm(); orchestrator.tick(); });
chrome.alarms.onAlarm.addListener((a) => {
  if (a.name === 'oea-poll' || a.name === 'oea-next') orchestrator.tick();
});

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (sender.id !== chrome.runtime.id) return false; // ignore other extensions / pages
  if (msg && msg.type === 'OEA_CLEAR_HISTORY') {
    orchestrator.clearHistory().then(sendResponse, () => sendResponse({ ok: false, reason: 'Unexpected error.' }));
    return true; // async response
  }
  const run = {
    OEA_START: () => orchestrator.start(),
    OEA_STOP: () => orchestrator.stop(),
    OEA_TICK: () => orchestrator.tick(),
  }[msg && msg.type];
  if (!run) return false;
  ensurePollAlarm();
  run().catch(() => {});
  sendResponse({ ok: true });
  return false;
});
