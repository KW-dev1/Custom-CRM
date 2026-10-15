// Minimal Google Sheets v4 client for one tab. Returns plain row objects; the job service depends
// only on this small interface (readRows / readRow / writeStatus), so it can be faked in tests.

import { colIndex } from './config.js';

export class SheetsError extends Error {
  constructor(message, { status = null, retryable = false } = {}) {
    super(message);
    this.name = 'SheetsError';
    this.status = status;
    this.retryable = retryable;
  }
}

const API = 'https://sheets.googleapis.com/v4/spreadsheets';
const quoteTab = (t) => `'${String(t).replace(/'/g, "''")}'`;

export function createSheetsClient({ cfg, auth, fetchImpl = globalThis.fetch, sleep = (ms) => new Promise((r) => setTimeout(r, ms)) }) {
  const { columns } = cfg;
  const letters = Object.values(columns).filter(Boolean);
  const minIdx = Math.min(...letters.map(colIndex));
  const maxIdx = Math.max(...letters.map(colIndex));
  const letter = (i) => { let s = ''; for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
  const first = letter(minIdx);
  const last = letter(maxIdx);
  const at = (colLetters) => colIndex(colLetters) - minIdx;
  let tabTitle = cfg.sheet.tab || null;

  async function request(path, { method = 'GET', body } = {}, attempt = 0) {
    const token = await auth.getAccessToken();
    let res;
    try {
      res = await fetchImpl(`${API}/${cfg.sheet.id}${path}`, {
        method,
        headers: { authorization: `Bearer ${token}`, ...(body ? { 'content-type': 'application/json' } : {}) },
        body: body ? JSON.stringify(body) : undefined,
      });
    } catch {
      throw new SheetsError('Could not reach the Google Sheets API.', { retryable: true });
    }
    if (res.status === 401 && attempt === 0) { auth.invalidate(); return request(path, { method, body }, 1); }
    if ((res.status === 429 || res.status >= 500) && attempt < 2 && method === 'GET') {
      await sleep(500 * 2 ** attempt);
      return request(path, { method, body }, attempt + 1);
    }
    if (!res.ok) {
      let detail = '';
      try { detail = (await res.json()).error?.message || ''; } catch { /* ignore */ }
      const hint = res.status === 403 ? ' (does this Google user have edit access to the sheet?)' : res.status === 404 ? ' (check SHEET_ID / SHEET_GID)' : '';
      throw new SheetsError(`Sheets API HTTP ${res.status}${detail ? `: ${detail}` : ''}${hint}`, { status: res.status, retryable: res.status === 429 || res.status >= 500 });
    }
    return res.json();
  }

  async function getTab() {
    if (tabTitle) return tabTitle;
    const meta = await request('?fields=sheets.properties(sheetId,title)');
    const sheets = (meta.sheets || []).map((s) => s.properties);
    const hit = sheets.find((p) => String(p.sheetId) === String(cfg.sheet.gid));
    if (!hit) throw new SheetsError(`No tab with gid ${cfg.sheet.gid} found in the spreadsheet.`, { status: 404 });
    tabTitle = hit.title;
    return tabTitle;
  }

  const toRow = (cells, rowNumber) => {
    const c = (col) => (col ? String(cells[at(col)] ?? '') : '');
    return { rowNumber, emails: c(columns.emails), subject: c(columns.subject), body: c(columns.body), status: c(columns.status), note: c(columns.note) };
  };

  return {
    /** All data rows below the header, including blank ones (so row numbers stay correct). */
    async readRows() {
      const tab = await getTab();
      const start = cfg.sheet.headerRows + 1;
      const range = encodeURIComponent(`${quoteTab(tab)}!${first}${start}:${last}`);
      const data = await request(`/values/${range}?majorDimension=ROWS&valueRenderOption=FORMATTED_VALUE`);
      return (data.values || []).map((cells, i) => toRow(cells, start + i));
    },
    async readRow(rowNumber) {
      if (!Number.isInteger(rowNumber) || rowNumber <= cfg.sheet.headerRows) throw new SheetsError('Invalid row number.');
      const tab = await getTab();
      const range = encodeURIComponent(`${quoteTab(tab)}!${first}${rowNumber}:${last}${rowNumber}`);
      const data = await request(`/values/${range}?majorDimension=ROWS&valueRenderOption=FORMATTED_VALUE`);
      return toRow((data.values && data.values[0]) || [], rowNumber);
    },
    /** Write the status (and optional note) cell(s). RAW input: values are never parsed as formulas. */
    async writeStatus(rowNumber, status, note) {
      if (!Number.isInteger(rowNumber) || rowNumber <= cfg.sheet.headerRows) throw new SheetsError('Refusing to write to a header or invalid row.');
      const tab = await getTab();
      const data = [{ range: `${quoteTab(tab)}!${columns.status}${rowNumber}`, values: [[status]] }];
      if (columns.note && note !== undefined) data.push({ range: `${quoteTab(tab)}!${columns.note}${rowNumber}`, values: [[note]] });
      await request('/values:batchUpdate', { method: 'POST', body: { valueInputOption: 'RAW', data } });
    },
  };
}
