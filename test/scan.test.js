const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const ExcelJS = require('exceljs');
const { scanFiles } = require('../src/engine/scan');

function tmpDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'fastexcel-scan-'));
}

async function makeFile(dir, name, cells) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  for (const [addr, v] of Object.entries(cells)) ws.getCell(addr).value = v;
  const file = path.join(dir, name);
  await wb.xlsx.writeFile(file);
  return file;
}

const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };

test('scanFiles processes every file and reports the right count', async () => {
  const dir = tmpDir();
  const a = await makeFile(dir, 'a.xlsx', { A1: '營收', B1: 1 });
  const b = await makeFile(dir, 'b.xlsx', { A1: '營收', B1: 2 });

  const outcome = await scanFiles([a, b], recipe, undefined, {});
  assert.equal(outcome.fileCount, 2);
  assert.equal(outcome.results.length, 2);
  assert.equal(outcome.errors.length, 0);
  assert.equal(outcome.cancelled, false);
});

test('scanFiles collects a per-file error instead of throwing for the whole batch', async () => {
  const dir = tmpDir();
  const good = await makeFile(dir, 'good.xlsx', { A1: '營收', B1: 1 });
  const bad = path.join(dir, 'not-really-excel.xlsx');
  fs.writeFileSync(bad, 'this is not a zip file');

  const outcome = await scanFiles([good, bad], recipe, undefined, {});
  assert.equal(outcome.results.length, 1);
  assert.equal(outcome.errors.length, 1);
  assert.equal(outcome.errors[0].file, bad);
});

test('scanFiles stops early when shouldCancel becomes true mid-batch', async () => {
  const dir = tmpDir();
  const files = await Promise.all(
    [1, 2, 3, 4].map((n) => makeFile(dir, `f${n}.xlsx`, { A1: '營收', B1: n }))
  );

  let processed = 0;
  const outcome = await scanFiles(files, recipe, undefined, {
    shouldCancel: () => processed >= 2,
    onProgress: () => {
      processed++;
    },
  });

  assert.equal(outcome.cancelled, true);
  assert.ok(outcome.results.length < files.length, 'should stop before processing every file');
});

test('scanFiles reports progress with the right done/total counts', async () => {
  const dir = tmpDir();
  const files = await Promise.all(
    [1, 2, 3].map((n) => makeFile(dir, `g${n}.xlsx`, { A1: '營收', B1: n }))
  );

  const progressCalls = [];
  await scanFiles(files, recipe, undefined, {
    onProgress: (p) => progressCalls.push({ done: p.done, total: p.total }),
  });

  assert.deepEqual(progressCalls, [
    { done: 1, total: 3 },
    { done: 2, total: 3 },
    { done: 3, total: 3 },
  ]);
});
