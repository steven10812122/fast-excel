const { test } = require('node:test');
const assert = require('node:assert/strict');
const { buildWorkbook, setBorder } = require('./helpers');
const { processLoadedWorkbook, extractManual } = require('../src/engine/match');

function values(block) {
  return block.cells.map((row) => row.map((c) => c.value));
}

test('finds a row block and stops at the blank cell', () => {
  const wb = buildWorkbook({
    Sheet1: { A1: '營收', B1: 100, C1: 200, D1: 300 },
  });
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const r = processLoadedWorkbook(wb, 'a.xlsx', recipe);
  const f = r.fields['營收'];
  assert.equal(f.matched, true);
  assert.deepEqual(values(f.block), [[100, 200, 300]]);
});

test('finds a column block and stops at the blank cell', () => {
  const wb = buildWorkbook({
    Sheet1: { B2: '支出', B3: 10, B4: 20 },
  });
  const recipe = { fields: [{ name: '支出', keywords: ['支出'] }] };
  const r = processLoadedWorkbook(wb, 'b.xlsx', recipe);
  const f = r.fields['支出'];
  assert.equal(f.matched, true);
  assert.deepEqual(values(f.block), [[10], [20]]);
});

test('resolves a full 2D block under a corner label', () => {
  const wb = buildWorkbook({
    財務: {
      A1: '部門預算',
      B1: '一月', C1: '二月',
      B2: 111, C2: 222,
      B3: 333, C3: 444,
    },
  });
  const recipe = { fields: [{ name: '部門預算', keywords: ['部門預算'] }] };
  const r = processLoadedWorkbook(wb, 'c.xlsx', recipe);
  const f = r.fields['部門預算'];
  assert.equal(f.block.rows, 3);
  assert.equal(f.block.cols, 2);
  assert.deepEqual(values(f.block), [['一月', '二月'], [111, 222], [333, 444]]);
});

test('a medium/thick/double border stops extension even with no blank gap', () => {
  const wb = buildWorkbook({
    Sheet1: { A1: '營收', B1: 100, C1: 200, D1: 300, E1: '不相干資料', F1: '更多不相干資料' },
  });
  setBorder(wb, 'Sheet1', 'D1', { right: { style: 'medium' } });
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const r = processLoadedWorkbook(wb, 'd.xlsx', recipe);
  assert.deepEqual(values(r.fields['營收'].block), [[100, 200, 300]]);
});

test('a header/label line above the data stops extension when data has no gap', () => {
  const wb = buildWorkbook({
    Sheet1: { B1: 'Q1', C1: 'Q2', D1: 'Q3', A2: '營收', B2: 100, C2: 200, D2: 300, E2: 999 },
  });
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const r = processLoadedWorkbook(wb, 'e.xlsx', recipe);
  assert.deepEqual(values(r.fields['營收'].block), [[100, 200, 300]]);
});

test('tolerates one missing value mid-block when the header line still vouches for it', () => {
  const wb = buildWorkbook({
    Sheet1: { B1: 'Q1', C1: 'Q2', D1: 'Q3', E1: 'Q4', A2: '營收', B2: 100, D2: 300, E2: 400 },
    // C2 (Q2) is deliberately left blank -- one missing month, not the end of the table.
  });
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const r = processLoadedWorkbook(wb, 'f.xlsx', recipe);
  const f = r.fields['營收'];
  assert.equal(f.block.cols, 4);
  assert.deepEqual(values(f.block), [[100, '', 300, 400]]);
});

test('tolerates the missing value even when it is the very first one in the row', () => {
  // Regression: the tolerated gap sits at the same (row, col) the
  // perpendicular-axis height scan starts from. Previously the height
  // scan's own unconditional blank check fired at n=0 there and reported
  // height=0, nuking the whole block even though the row's width was
  // computed correctly.
  const wb = buildWorkbook({
    Sheet1: { B1: 'Q1', C1: 'Q2', D1: 'Q3', A2: '營收', C2: 200, D2: 300 },
    // B2 (Q1) is blank -- the very first data cell in the row.
  });
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const r = processLoadedWorkbook(wb, 'f2.xlsx', recipe);
  const f = r.fields['營收'];
  assert.equal(f.matched, true);
  assert.deepEqual(values(f.block), [['', 200, 300]]);
});

test('tolerates the missing value even when it is the very last one in the row', () => {
  // Regression: the lookahead used to require the *next* cell to have
  // real data to confirm a gap was genuinely isolated -- which a trailing
  // gap can never satisfy, since there is no next cell. When the header
  // line itself also stops right there (there was never going to be a
  // next column anyway), the gap should still be trusted as the final
  // entry instead of being dropped.
  const wb = buildWorkbook({
    Sheet1: { B1: 'Q1', C1: 'Q2', D1: 'Q3', A2: '營收', B2: 100, C2: 200 },
    // D2 (Q3) is blank -- the very last data cell in the row.
  });
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const r = processLoadedWorkbook(wb, 'f3.xlsx', recipe);
  const f = r.fields['營收'];
  assert.equal(f.matched, true);
  assert.deepEqual(values(f.block), [[100, 200, '']]);
});

test('two consecutive missing values still end the block', () => {
  const wb = buildWorkbook({
    Sheet1: {
      B1: 'Q1', C1: 'Q2', D1: 'Q3', E1: 'Q4',
      A2: '營收', B2: 100,
      // C2 and D2 both blank -- two gaps in a row, not tolerated.
      E2: 400,
    },
  });
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const r = processLoadedWorkbook(wb, 'g.xlsx', recipe);
  assert.deepEqual(values(r.fields['營收'].block), [[100]]);
});

test('a field stops exactly at another field\'s own anchor with no gap between them', () => {
  const wb = buildWorkbook({
    Sheet1: { A2: '營收', B2: 100, C2: 200, D2: 300, E2: '虧損', F2: 10, G2: 20 },
  });
  const recipe = {
    fields: [
      { name: '營收', keywords: ['營收'] },
      { name: '虧損', keywords: ['虧損'] },
    ],
  };
  const r = processLoadedWorkbook(wb, 'h.xlsx', recipe);
  assert.deepEqual(values(r.fields['營收'].block), [[100, 200, 300]]);
  assert.deepEqual(values(r.fields['虧損'].block), [[10, 20]]);
});

test('height growth does not swallow the next field\'s row underneath it', () => {
  const wb = buildWorkbook({
    Sheet1: { A2: '營收', B2: 100, C2: 200, A3: '支出', B3: 10, C3: 20 },
  });
  const recipe = {
    fields: [
      { name: '營收', keywords: ['營收'] },
      { name: '支出', keywords: ['支出'] },
    ],
  };
  const r = processLoadedWorkbook(wb, 'i.xlsx', recipe);
  assert.deepEqual(values(r.fields['營收'].block), [[100, 200]]);
  assert.deepEqual(values(r.fields['支出'].block), [[10, 20]]);
});

test('a merged, centered title fixes the block width regardless of the data row', () => {
  const wb = buildWorkbook(
    { Sheet1: { A1: '營收', A2: 1, B2: 2, C2: 3, D2: 4, A3: 5, B3: 6, C3: 7, D3: 8 } },
    { Sheet1: ['A1:D1'] }
  );
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const r = processLoadedWorkbook(wb, 'j.xlsx', recipe);
  const f = r.fields['營收'];
  assert.equal(f.block.rows, 2);
  assert.equal(f.block.cols, 4);
});

test('direction is auto-picked per anchor -- no per-field config needed', () => {
  const wb = buildWorkbook({
    Sheet1: {
      // 營收: label left of a row of values
      A1: '營收', B1: 1, C1: 2, D1: 3,
      // 費用: label above a column of values
      A3: '費用', A4: 9, A5: 8,
    },
  });
  const recipe = {
    fields: [
      { name: '營收', keywords: ['營收'] },
      { name: '費用', keywords: ['費用'] },
    ],
  };
  const r = processLoadedWorkbook(wb, 'k.xlsx', recipe);
  assert.deepEqual(values(r.fields['營收'].block), [[1, 2, 3]]);
  assert.deepEqual(values(r.fields['費用'].block), [[9], [8]]);
});

test('flags ambiguous when both directions found real data', () => {
  const wb = buildWorkbook({
    Sheet1: { A1: '營收', B1: 100, C1: 200, A2: 50 },
  });
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const r = processLoadedWorkbook(wb, 'l.xlsx', recipe);
  assert.equal(r.fields['營收'].ambiguous, true);
});

test('an unmatched field is reported as not matched, not thrown', () => {
  const wb = buildWorkbook({ Sheet1: { A1: '其他東西' } });
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const r = processLoadedWorkbook(wb, 'm.xlsx', recipe);
  assert.equal(r.fields['營收'].matched, false);
});

test('extractManual re-anchors at a hand-picked cell, bypassing keyword search', () => {
  const wb = buildWorkbook({
    // No cell here says "營收" at all -- a real "auto-match found nothing"
    // case that only a manual override can rescue.
    Sheet1: { A2: '本期實際數字', B2: 700, C2: 800 },
  });
  const ws = wb.getWorksheet('Sheet1');
  const fr = extractManual(ws, 'A2', new Set());
  assert.equal(fr.matched, true);
  assert.equal(fr.manual, true);
  assert.equal(fr.headerText, '本期實際數字');
  assert.deepEqual(fr.block.cells[0].map((c) => c.value), [700, 800]);
});

test('extractManual still stops at another field\'s existing anchor', () => {
  const wb = buildWorkbook({
    Sheet1: { A2: '本期', B2: 700, C2: 800, D2: '虧損', E2: 10 },
  });
  const ws = wb.getWorksheet('Sheet1');
  const fr = extractManual(ws, 'A2', new Set(['D2']));
  assert.deepEqual(fr.block.cells[0].map((c) => c.value), [700, 800]);
});

test('extractManual returns null for a garbage address', () => {
  const wb = buildWorkbook({ Sheet1: { A1: 'x' } });
  const ws = wb.getWorksheet('Sheet1');
  assert.equal(extractManual(ws, 'not-a-cell', new Set()), null);
});

test('a formula error surfaces as its error string, not [object Object] or a false match', () => {
  const wb = buildWorkbook({
    Sheet1: { A1: '營收', B1: 100, C1: { formula: '1/0', result: { error: '#DIV/0!' } }, D1: 300 },
  });
  const recipe = { fields: [{ name: '營收', keywords: ['營收'] }] };
  const r = processLoadedWorkbook(wb, 'n.xlsx', recipe);
  // the error cell is not blank, so extension continues through it, and
  // its value comes back as the plain error string, not the raw object.
  assert.deepEqual(values(r.fields['營收'].block), [[100, '#DIV/0!', 300]]);
});
