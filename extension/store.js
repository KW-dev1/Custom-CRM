// Thin wrapper over a chrome.storage area (injectable for tests). All durable state lives here
// because the MV3 service worker can be suspended at any time.

import { normalizeSettings } from './config.js';

export const KEYS = {
  settings: 'settings', runtime: 'runtime', run: 'run', completed: 'completed',
  pendingReport: 'pendingReport', inflight: 'inflight', status: 'status',
  lastResult: 'lastResult', errors: 'errors', logs: 'logs',
};

export function createStore(area) {
  const get = async (k, d) => {
    const r = await area.get(k);
    return r[k] === undefined ? d : r[k];
  };
  const set = (k, v) => area.set({ [k]: v });
  const del = (k) => area.remove(k);
  return {
    getSettings: async () => normalizeSettings(await get(KEYS.settings, {})),
    saveSettings: (s) => set(KEYS.settings, s),
    getRuntime: () => get(KEYS.runtime, { running: false, consecutiveFailures: 0, lastFinishedAt: 0, nextClaimAt: 0 }),
    patchRuntime: async (p) => set(KEYS.runtime, { ...(await get(KEYS.runtime, { running: false, consecutiveFailures: 0, lastFinishedAt: 0, nextClaimAt: 0 })), ...p }),
    getRun: () => get(KEYS.run, null),
    setRun: (r) => set(KEYS.run, r),
    clearRun: () => del(KEYS.run),
    getCompletedMap: () => get(KEYS.completed, {}),
    setCompletedMap: (m) => set(KEYS.completed, m),
    getPendingReport: () => get(KEYS.pendingReport, null),
    setPendingReport: (r) => set(KEYS.pendingReport, r),
    clearPendingReport: () => del(KEYS.pendingReport),
    getInflight: () => get(KEYS.inflight, null),
    clearInflight: () => del(KEYS.inflight),
    getStatus: () => get(KEYS.status, null),
    setStatus: (s) => set(KEYS.status, s),
    getLastResult: () => get(KEYS.lastResult, null),
    setLastResult: (r) => set(KEYS.lastResult, r),
    getList: (k) => get(k, []),
    setList: (k, v) => set(k, v),
  };
}
