const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const ExcelJS = require('exceljs');
const { exportConsolidated } = require('../src/engine/export');

function block(rows) {
  return {
    rows: rows.length,
    cols: rows[0].length,
    cells: rows.map((r) => r.map((v) => ({ address: 'A1', value: v }))),
  };
}

test('a single-row field with varying widths spreads into numbered columns', async () => {
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const results = [
    { file: 'a.xlsx', fields: { 營收: { matched: true, block: block([[100, 200, 300]]) } } },
    { file: 'b.xlsx', fields: { 營收: { matched: true, block: block([[10, 20]]) } } },
    { file: 'c.xlsx', fields: { 營收: { matched: false } } },
  ];

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'fastexcel-export-'));
  const outPath = path.join(dir, 'out.xlsx');
  await exportConsolidated(results, recipe, outPath);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(outPath);
  const sheet = wb.getWorksheet('Consolidated');

  assert.deepEqual(sheet.getRow(1).values.slice(1), ['Source file', '營收 1', '營收 2', '營收 3']);
  assert.deepEqual(sheet.getRow(2).values.slice(1), ['a.xlsx', 100, 200, 300]);
  assert.deepEqual(sheet.getRow(3).values.slice(1), ['b.xlsx', 10, 20, '']);
});

test('a single scalar field keeps a plain column with no numbered suffix', async () => {
  const recipe = { fields: [{ name: '支出', keywords: ['支出'] }] };
  const results = [{ file: 'a.xlsx', fields: { 支出: { matched: true, block: block([[500]]) } } }];

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'fastexcel-export-'));
  const outPath = path.join(dir, 'out.xlsx');
  await exportConsolidated(results, recipe, outPath);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(outPath);
  const sheet = wb.getWorksheet('Consolidated');
  assert.deepEqual(sheet.getRow(1).values.slice(1), ['Source file', '支出']);
  assert.equal(sheet.getRow(2).getCell(2).value, 500);
});

test('a genuine 2D block falls back to one joined-text column', async () => {
  const recipe = { fields: [{ name: '部門預算', keywords: ['部門預算'] }] };
  const results = [
    {
      file: 'a.xlsx',
      fields: { 部門預算: { matched: true, block: block([[1, 2], [3, 4]]) } },
    },
  ];

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'fastexcel-export-'));
  const outPath = path.join(dir, 'out.xlsx');
  await exportConsolidated(results, recipe, outPath);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(outPath);
  const sheet = wb.getWorksheet('Consolidated');
  assert.deepEqual(sheet.getRow(1).values.slice(1), ['Source file', '部門預算']);
  assert.equal(sheet.getRow(2).getCell(2).value, '1, 2 | 3, 4');
});

test('the source-file column label follows the caller-supplied language', async () => {
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const results = [{ file: 'a.xlsx', fields: { 營收: { matched: true, block: block([[1]]) } } }];

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'fastexcel-export-'));
  const outPath = path.join(dir, 'out.xlsx');
  await exportConsolidated(results, recipe, outPath, '來源檔案');

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(outPath);
  const sheet = wb.getWorksheet('Consolidated');
  assert.equal(sheet.getRow(1).getCell(1).value, '來源檔案');
});
