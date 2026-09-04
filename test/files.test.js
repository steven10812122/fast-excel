const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { listExcelFiles, loadWorkbook } = require('../src/engine/files');
const { processLoadedWorkbook } = require('../src/engine/match');

function tmpDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'fastexcel-files-'));
}

test('listExcelFiles ignores lock files, dotfiles, and non-excel files', () => {
  const dir = tmpDir();
  fs.writeFileSync(path.join(dir, 'real.xlsx'), '');
  fs.writeFileSync(path.join(dir, '~$real.xlsx'), ''); // Excel's own lock file
  fs.writeFileSync(path.join(dir, '.DS_Store'), '');
  fs.writeFileSync(path.join(dir, 'notes.txt'), '');
  fs.writeFileSync(path.join(dir, 'legacy.xls'), '');

  const found = listExcelFiles(dir, false).map((f) => path.basename(f)).sort();
  assert.deepEqual(found, ['legacy.xls', 'real.xlsx']);
});

test('listExcelFiles only descends into subfolders when recursive is true', () => {
  const dir = tmpDir();
  fs.mkdirSync(path.join(dir, 'sub'));
  fs.writeFileSync(path.join(dir, 'top.xlsx'), '');
  fs.writeFileSync(path.join(dir, 'sub', 'nested.xlsx'), '');

  assert.deepEqual(
    listExcelFiles(dir, false).map((f) => path.basename(f)),
    ['top.xlsx']
  );
  assert.deepEqual(
    listExcelFiles(dir, true).map((f) => path.basename(f)).sort(),
    ['nested.xlsx', 'top.xlsx']
  );
});

test('loadWorkbook reads a real .xlsx file from disk', async () => {
  const dir = tmpDir();
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell('A1').value = '營收';
  ws.getCell('B1').value = 42;
  const file = path.join(dir, 'a.xlsx');
  await wb.xlsx.writeFile(file);

  const loaded = await loadWorkbook(file);
  assert.equal(loaded.getWorksheet('Sheet1').getCell('A1').value, '營收');
});

test('loadWorkbook converts a genuine legacy .xls file and it matches normally', async () => {
  const dir = tmpDir();
  const legacy = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['營收', 100, 200, 300],
  ]);
  XLSX.utils.book_append_sheet(legacy, sheet, 'Sheet1');
  const file = path.join(dir, 'old.xls');
  XLSX.writeFile(legacy, file, { bookType: 'biff8' });

  const workbook = await loadWorkbook(file);
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const result = processLoadedWorkbook(workbook, file, recipe);
  assert.equal(result.fields['營收'].matched, true);
  assert.deepEqual(
    result.fields['營收'].block.cells[0].map((c) => c.value),
    [100, 200, 300]
  );
});
