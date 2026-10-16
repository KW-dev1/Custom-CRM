// Local Excel (.xlsx) data source. Same small interface as the Google Sheets client
// (readRows / readRow / writeStatus), so the job service does not care where rows live.
//
// Safety:
//  * every operation re-reads the file from disk, so edits you saved in Excel are always seen;
//  * writes refuse to run while the workbook is open in Excel/LibreOffice (lock file present),
//    because saving over an open file would lose data or fail;
//  * a write replaces only the status (and optional note) cell, saves to a temp file and renames
//    it over the original, then re-reads the cell to confirm it really landed;
//  * a one-time backup copy of the file is made before the first write of each process.

import ExcelJS from 'exceljs';
import { access, copyFile, rename, rm } from 'node:fs/promises';
import { basename, dirname, join } from 'node:path';
import { SheetsError } from './sheets-client.js';

export function cellText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join('');
    if ('result' in v) return cellText(v.result); // formula: use the cached value
    if ('text' in v) return cellText(v.text);     // hyperlink
    if ('error' in v) return '';
  }
  return String(v);
}

/** Names of the "someone has this open" marker files Excel / LibreOffice create next to a workbook. */
export function lockFileNames(file) {
  const name = basename(file);
  // Excel truncates the first two characters for longer names, so check both spellings.
  return [`~$${name}`, `~$${name.slice(2)}`, `.~lock.${name}#`];
}

export function createXlsxStore({ cfg, now = Date.now }) {
  const file = cfg.source.path;
  const { columns } = cfg;
  const headerRows = cfg.sheet.headerRows;
  let backedUp = false;

  async function load() {
    try { await access(file); } catch { throw new SheetsError(`Excel file not found: ${file}`, { status: 404 }); }
    const wb = new ExcelJS.Workbook();
    try { await wb.xlsx.readFile(file); }
    catch (e) { throw new SheetsError(`Could not read the Excel file (${e.message}). Is it a valid .xlsx?`, { retryable: true }); }
    const ws = cfg.source.sheetName ? wb.getWorksheet(cfg.source.sheetName) : wb.worksheets[0];
    if (!ws) throw new SheetsError(`Worksheet "${cfg.source.sheetName || '(first)'}" not found in the Excel file.`, { status: 404 });
    return { wb, ws };
  }
  const text = (ws, row, col) => (col ? cellText(ws.getCell(`${col}${row}`).value) : '');
  const toRow = (ws, n) => ({
    rowNumber: n, emails: text(ws, n, columns.emails), subject: text(ws, n, columns.subject),
    body: text(ws, n, columns.body), status: text(ws, n, columns.status), note: text(ws, n, columns.note),
    sender: text(ws, n, columns.sender),
  });

  async function assertNotOpenElsewhere() {
    for (const name of lockFileNames(file)) {
      try {
        await access(join(dirname(file), name));
        throw new SheetsError(`The file is open in Excel (found ${name}). Close it so the result can be recorded; the job will be retried.`, { retryable: true });
      } catch (e) {
        if (e instanceof SheetsError) throw e;
      }
    }
  }

  return {
    async readRows() {
      const { ws } = await load();
      const out = [];
      for (let n = headerRows + 1; n <= ws.rowCount; n++) out.push(toRow(ws, n));
      return out;
    },
    async readRow(n) {
      if (!Number.isInteger(n) || n <= headerRows) throw new SheetsError('Invalid row number.');
      const { ws } = await load();
      return toRow(ws, n);
    },
    /** extra.sender: name of the extension profile that claimed the row (written when a sender column is configured). */
    async writeStatus(n, status, note, extra = {}) {
      if (!Number.isInteger(n) || n <= headerRows) throw new SheetsError('Refusing to write to a header or invalid row.');
      await assertNotOpenElsewhere();
      if (!backedUp) {
        const stamp = new Date(now()).toISOString().replace(/[:.]/g, '-');
        await copyFile(file, `${file}.${stamp}.bak`).catch((e) => { throw new SheetsError(`Could not create a backup copy: ${e.message}`, { retryable: true }); });
        backedUp = true;
      }
      const { wb, ws } = await load();
      ws.getCell(`${columns.status}${n}`).value = String(status); // plain string: never a formula
      if (columns.note && note !== undefined) ws.getCell(`${columns.note}${n}`).value = String(note);
      if (columns.sender && extra.sender !== undefined) ws.getCell(`${columns.sender}${n}`).value = String(extra.sender);
      const tmp = `${file}.${process.pid}.tmp`;
      try {
        await wb.xlsx.writeFile(tmp);
        await rename(tmp, file);
      } catch (e) {
        await rm(tmp, { force: true });
        throw new SheetsError(`Could not save the Excel file (${e.code || e.message}). Is it open in another program?`, { retryable: true });
      }
      const check = await load();
      if (text(check.ws, n, columns.status) !== String(status)) {
        throw new SheetsError('The status was written but could not be read back; refusing to continue.', { retryable: true });
      }
    },
  };
}
