import { loadConfig } from '../src/config.js';

export const testConfig = (extra = {}) => loadConfig({
  API_KEYS: 'abcdefghijklmnop1234', XLSX_PATH: '/tmp/fake-leads.xlsx',
  MIN_CLAIM_INTERVAL_SEC: '0', COL_NOTE: 'K', ...extra,
});

/** In-memory stand-in for the Sheets client (rows start at sheet row 2). */
export class FakeSheet {
  constructor(rows = []) {
    this.rows = rows.map((r, i) => ({ rowNumber: i + 2, emails: '', subject: '', body: '', status: '', note: '', ...r }));
    this.writes = [];
    this.failWrites = false;
    this.delay = 0;
  }
  async #wait() { if (this.delay) await new Promise((r) => setTimeout(r, this.delay)); }
  async readRows() { await this.#wait(); return this.rows.map((r) => ({ ...r })); }
  async readRow(n) { await this.#wait(); return { ...(this.rows.find((r) => r.rowNumber === n) || { rowNumber: n, emails: '', subject: '', body: '', status: '', note: '' }) }; }
  async writeStatus(n, status, note) {
    await this.#wait();
    if (this.failWrites) throw new Error('sheet write failed');
    this.writes.push({ row: n, status, note });
    const r = this.rows.find((x) => x.rowNumber === n);
    if (r) { r.status = status; if (note !== undefined) r.note = note; }
  }
}
export const lead = (n, extra = {}) => ({ emails: `["lead${n}@x.com","cc${n}@y.org"]`, subject: `Subject ${n}`, body: `Body ${n}`, ...extra });
