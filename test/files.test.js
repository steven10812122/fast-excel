const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { listDataFiles, loadWorkbook } = require('../src/engine/files');
const { processLoadedWorkbook } = require('../src/engine/match');

function tmpDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'fastexcel-files-'));
}

test('listDataFiles includes .csv, ignores lock files, dotfiles, and unrelated files', () => {
  const dir = tmpDir();
  fs.writeFileSync(path.join(dir, 'real.xlsx'), '');
  fs.writeFileSync(path.join(dir, '~$real.xlsx'), ''); // Excel's own lock file
  fs.writeFileSync(path.join(dir, '.DS_Store'), '');
  fs.writeFileSync(path.join(dir, 'notes.txt'), '');
  fs.writeFileSync(path.join(dir, 'legacy.xls'), '');
  fs.writeFileSync(path.join(dir, 'export.csv'), '');

  const found = listDataFiles(dir, false).map((f) => path.basename(f)).sort();
  assert.deepEqual(found, ['export.csv', 'legacy.xls', 'real.xlsx']);
});

test('listDataFiles only descends into subfolders when recursive is true', () => {
  const dir = tmpDir();
  fs.mkdirSync(path.join(dir, 'sub'));
  fs.writeFileSync(path.join(dir, 'top.xlsx'), '');
  fs.writeFileSync(path.join(dir, 'sub', 'nested.xlsx'), '');

  assert.deepEqual(
    listDataFiles(dir, false).map((f) => path.basename(f)),
    ['top.xlsx']
  );
  assert.deepEqual(
    listDataFiles(dir, true).map((f) => path.basename(f)).sort(),
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

test('loadWorkbook reads a .csv file and it matches normally, including a blank field', async () => {
  const dir = tmpDir();
  const file = path.join(dir, 'export.csv');
  fs.writeFileSync(
    file,
    '項目,一月,二月,三月\n營收,100,,300\n',
    'utf-8'
  );

  const workbook = await loadWorkbook(file);
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const result = processLoadedWorkbook(workbook, file, recipe);
  assert.equal(result.fields['營收'].matched, true);
  // row 1's header line (一月/二月/三月) still vouches for the gap at
  // 二月 exactly like it would in a real .xlsx -- same engine, same rule.
  assert.deepEqual(
    result.fields['營收'].block.cells[0].map((c) => c.value),
    [100, '', 300]
  );
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
