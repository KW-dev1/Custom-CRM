import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { mkdtemp, writeFile, readdir, access } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createXlsxStore, cellText, lockFileNames } from '../src/xlsx-store.js';
import { createJobService } from '../src/jobs.js';
import { SheetsError } from '../src/sheets-client.js';
import { loadConfig } from '../src/config.js';

async function makeBook({ extra = {} } = {}) {
  const dir = await mkdtemp(join(tmpdir(), 'oea-xlsx-'));
  const file = join(dir, 'contact_out.xlsx');
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('sheet1');
  ws.addRow(['Company', 'Website', 'Role', 'First', 'Last', 'LinkedIn', 'Email', 'Subject', 'Message', 'Status']);
  ws.addRow(['Acme', 'https://acme.io', 'CEO', 'Ann', 'A', 'li', 'ann@acme.io, bob@acme.io', 'Hello Ann', 'Line one\nLine two', null]);
  ws.addRow(['Beta', 'https://beta.io', 'CTO', 'Bo', 'B', 'li', 'solo@beta.io, ', 'Hi Bo', 'Body Bo', null]);
  ws.addRow([]); // blank row
  ws.addRow(['Gamma', 'g', 'VP', 'Gi', 'G', 'li', 'gi@gamma.io', '', '', null]); // empty subject/body
  ws.getCell('A2').font = { bold: true };
  ws.getCell('B2').numFmt = '@';
  if (extra.rich) ws.getCell('H3').value = { richText: [{ text: 'Hi ' }, { text: 'Bo', font: { bold: true } }] };
  if (extra.link) ws.getCell('G3').value = { text: 'solo@beta.io', hyperlink: 'mailto:solo@beta.io' };
  if (extra.formula) ws.getCell('H3').value = { formula: '"Hi "&D3', result: 'Hi Bo' };
  const second = wb.addWorksheet('other');
  second.addRow(['x']);
  await wb.xlsx.writeFile(file);
  return { dir, file };
}
const cfgFor = (file, env = {}) => loadConfig({
  DATA_SOURCE: 'xlsx', XLSX_PATH: file, API_KEYS: 'abcdefghijklmnop1234', MIN_CLAIM_INTERVAL_SEC: '0', ...env,
});
const reopen = async (file) => { const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(file); return wb.getWorksheet('sheet1'); };

test('cellText normalises rich text, hyperlinks, formulas, numbers, dates, errors', () => {
  assert.equal(cellText(null), '');
  assert.equal(cellText(42), '42');
  assert.equal(cellText({ richText: [{ text: 'a' }, { text: 'b' }] }), 'ab');
  assert.equal(cellText({ text: 'x@y.com', hyperlink: 'mailto:x@y.com' }), 'x@y.com');
  assert.equal(cellText({ formula: 'A1', result: 'val' }), 'val');
  assert.equal(cellText({ error: '#N/A' }), '');
  assert.match(cellText(new Date(0)), /1970/);
});
test('lock-file names cover Excel (both spellings) and LibreOffice', () => {
  assert.deepEqual(lockFileNames('/d/contact_out.xlsx'), ['~$contact_out.xlsx', '~$ntact_out.xlsx', '.~lock.contact_out.xlsx#']);
});

test('reads rows below the header with correct numbers, including blanks', async () => {
  const { file } = await makeBook();
  const rows = await createXlsxStore({ cfg: cfgFor(file) }).readRows();
  assert.deepEqual(rows.map((r) => r.rowNumber), [2, 3, 4, 5]);
  assert.equal(rows[0].emails, 'ann@acme.io, bob@acme.io');
  assert.equal(rows[0].body, 'Line one\nLine two');
  assert.equal(rows[0].status, '');
  assert.equal(rows[2].emails, '');
});
test('rich text, hyperlink and formula cells are read as text', async () => {
  for (const extra of [{ rich: true }, { link: true }, { formula: true }]) {
    const { file } = await makeBook({ extra });
    const r = await createXlsxStore({ cfg: cfgFor(file) }).readRow(3);
    assert.ok(r.subject === 'Hi Bo' && r.emails.startsWith('solo@beta.io'), JSON.stringify(extra));
  }
});
test('sheet can be chosen by name', async () => {
  const { file } = await makeBook();
  const rows = await createXlsxStore({ cfg: cfgFor(file, { XLSX_SHEET: 'other' }) }).readRows();
  assert.equal(rows.length, 0);
  await assert.rejects(createXlsxStore({ cfg: cfgFor(file, { XLSX_SHEET: 'nope' }) }).readRows(), /not found/);
});
test('missing file gives a clear error', async () => {
  await assert.rejects(createXlsxStore({ cfg: cfgFor('/nonexistent/x.xlsx') }).readRows(), (e) => e instanceof SheetsError && /not found/.test(e.message));
});

test('writeStatus changes only the status cell, keeps other cells and formatting, leaves no temp file', async () => {
  const { dir, file } = await makeBook();
  const store = createXlsxStore({ cfg: cfgFor(file) });
  await store.writeStatus(2, 'pending');
  const ws = await reopen(file);
  assert.equal(ws.getCell('J2').value, 'pending');
  assert.equal(ws.getCell('G2').value, 'ann@acme.io, bob@acme.io');
  assert.equal(ws.getCell('I2').value, 'Line one\nLine two');
  assert.equal(ws.getCell('A2').font.bold, true);
  assert.equal(ws.getCell('B2').numFmt, '@');
  assert.equal(ws.getCell('J3').value, null);
  const files = await readdir(dir);
  assert.ok(!files.some((f) => f.endsWith('.tmp')), 'no temp file left');
});
test('a backup copy is made once, before the first write', async () => {
  const { dir, file } = await makeBook();
  const store = createXlsxStore({ cfg: cfgFor(file) });
  await store.writeStatus(2, 'pending');
  await store.writeStatus(3, 'pending');
  const baks = (await readdir(dir)).filter((f) => f.endsWith('.bak'));
  assert.equal(baks.length, 1);
  const orig = await reopen(join(dir, baks[0]));
  assert.equal(orig.getCell('J2').value, null, 'backup holds the pre-write state');
});
test('note column is written when configured', async () => {
  const { file } = await makeBook();
  await createXlsxStore({ cfg: cfgFor(file, { COL_NOTE: 'K' }) }).writeStatus(2, 'sent', '2026 verified');
  assert.equal((await reopen(file)).getCell('K2').value, '2026 verified');
});
test('values are stored as plain text, never formulas', async () => {
  const { file } = await makeBook();
  await createXlsxStore({ cfg: cfgFor(file) }).writeStatus(2, '=HYPERLINK("http://x")');
  const v = (await reopen(file)).getCell('J2').value;
  assert.equal(typeof v, 'string');
});
test('refuses to write while the workbook is open in Excel (lock file)', async () => {
  const { dir, file } = await makeBook();
  await writeFile(join(dir, '~$ntact_out.xlsx'), 'lock');
  const store = createXlsxStore({ cfg: cfgFor(file) });
  await assert.rejects(store.writeStatus(2, 'pending'), (e) => e instanceof SheetsError && e.retryable && /open in Excel/.test(e.message));
  assert.equal((await reopen(file)).getCell('J2').value, null);
  assert.ok((await readdir(dir)).every((f) => !f.endsWith('.bak')), 'no backup/write attempted either');
});
test('refuses header rows and invalid rows', async () => {
  const { file } = await makeBook();
  const store = createXlsxStore({ cfg: cfgFor(file) });
  await assert.rejects(store.writeStatus(1, 'x'), /header/);
  await assert.rejects(store.readRow(0), SheetsError);
});
test('edits saved in Excel between operations are picked up (no stale cache)', async () => {
  const { file } = await makeBook();
  const store = createXlsxStore({ cfg: cfgFor(file) });
  assert.equal((await store.readRow(2)).subject, 'Hello Ann');
  const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(file);
  wb.getWorksheet('sheet1').getCell('H2').value = 'Edited in Excel';
  await wb.xlsx.writeFile(file);
  assert.equal((await store.readRow(2)).subject, 'Edited in Excel');
  await store.writeStatus(3, 'pending');
  assert.equal((await reopen(file)).getCell('H2').value, 'Edited in Excel', 'the edit survives our write');
});
test('writes serialised through the job service land correctly on the real file', async () => {
  const { file } = await makeBook();
  const cfg = cfgFor(file, { COL_NOTE: 'K' });
  const svc = createJobService({ sheet: createXlsxStore({ cfg }), cfg });
  const a = await svc.claimNext();
  assert.equal(a.to, 'ann@acme.io');
  assert.deepEqual(a.cc, ['bob@acme.io']);
  assert.equal((await reopen(file)).getCell('J2').value, 'pending');
  const b = await svc.claimNext();
  assert.equal(b.to, 'solo@beta.io');
  assert.deepEqual(b.cc, []);
  const c = await svc.claimNext(); // Gamma has no subject/message: marked failed, then queue is empty
  assert.equal(c, null);
  assert.equal((await reopen(file)).getCell('J5').value, 'failed');
  await svc.report(a.id, 'sent', 'Verified: compose window closed');
  const ws = await reopen(file);
  assert.equal(ws.getCell('J2').value, 'sent');
  assert.match(ws.getCell('K2').value, /Verified/);
  assert.equal(ws.getCell('J3').value, 'pending');
  await assert.rejects(access(`${file}.tmp`));
});
