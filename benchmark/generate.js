// Generates a 24-file "two years of monthly reports" benchmark batch,
// deliberately cycling through every messy layout pattern the matching
// engine was built to handle -- position drift, merged titles, borders,
// zero-gap adjacent fields, sparse missing values, legacy .xls, and mixed
// English/Chinese keywords with a formula error thrown in. Ground truth
// is recorded at generation time, so correctness can be checked exactly
// afterward instead of eyeballed.

const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

const OUT_DIR = path.join(__dirname, 'files');
const TRUTH_PATH = path.join(__dirname, 'truth.json');

// Deterministic PRNG so the batch is reproducible across runs.
function mulberry32(seed) {
  return function () {
    seed |= 0;
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const rand = mulberry32(20240601);
const randInt = (min, max) => min + Math.floor(rand() * (max - min + 1));

const FIELDS = [
  { name: '營收', enName: 'Revenue', keywords: ['營收', '營業收入', 'Revenue'] },
  { name: '支出', enName: 'Cost', keywords: ['支出', '成本', 'Cost'] },
  { name: '毛利', enName: 'Gross Profit', keywords: ['毛利', 'Gross Profit'] },
  { name: '淨利', enName: 'Net Income', keywords: ['淨利', '稅後淨利', 'Net Income'] },
];

function colLetter(n) {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function randomValues(n) {
  return Array.from({ length: n }, () => randInt(50, 9999) * 100);
}

async function buildFile(index, month) {
  const style = index % 8;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  const truth = { file: '', fields: {} };
  const rowOffset = randInt(0, 2); // simulates "someone inserted a row this month"
  const colOffset = randInt(0, 2);

  const values3 = () => randomValues(3);

  if (style === 4) {
    // Two fields packed with zero gap -- only makes sense for exactly a
    // pair, so handle it as a special full-file case.
    const row = 2 + rowOffset;
    let col = 1 + colOffset;
    const [a, b] = FIELDS;
    ws.getCell(row, col).value = a.name;
    truth.fields[a.name] = { anchor: { row, col }, values: values3() };
    truth.fields[a.name].values.forEach((v, i) => (ws.getCell(row, col + 1 + i).value = v));
    col = col + 1 + truth.fields[a.name].values.length;
    ws.getCell(row, col).value = b.name;
    truth.fields[b.name] = { anchor: { row, col }, values: values3() };
    truth.fields[b.name].values.forEach((v, i) => (ws.getCell(row, col + 1 + i).value = v));
    // remaining two fields laid out normally, a couple rows below
    let r2 = row + 2;
    for (const f of FIELDS.slice(2)) {
      ws.getCell(r2, 1 + colOffset).value = f.name;
      const vals = values3();
      vals.forEach((v, i) => (ws.getCell(r2, 2 + colOffset + i).value = v));
      truth.fields[f.name] = { anchor: { row: r2, col: 1 + colOffset }, values: vals };
      r2 += 1;
    }
  } else {
    let row = 1 + rowOffset;
    for (const f of FIELDS) {
      const useEnglish = style === 7 && f === FIELDS[0];
      const label = useEnglish ? f.enName : f.name;
      const vals = values3();

      if (style === 2) {
        // Merged, centered title above a row of values.
        const startCol = 1 + colOffset;
        const endCol = startCol + vals.length - 1;
        ws.getCell(row, startCol).value = label;
        ws.mergeCells(row, startCol, row, endCol);
        ws.getCell(row, startCol).alignment = { horizontal: 'center' };
        vals.forEach((v, i) => (ws.getCell(row + 1, startCol + i).value = v));
        truth.fields[f.name] = { anchor: { row, col: startCol }, values: vals };
        row += 3;
      } else if (style === 3) {
        // Border-marked edge, unrelated data glued on right after.
        const startCol = 1 + colOffset;
        ws.getCell(row, startCol).value = label;
        vals.forEach((v, i) => (ws.getCell(row, startCol + 1 + i).value = v));
        const lastCol = startCol + vals.length;
        ws.getCell(row, lastCol).border = { right: { style: 'medium' } };
        ws.getCell(row, lastCol + 1).value = '備註代碼X' + randInt(100, 999);
        truth.fields[f.name] = { anchor: { row, col: startCol }, values: vals };
        row += 2;
      } else if (style === 5) {
        // Header row above, with one isolated missing value tolerated.
        const startCol = 1 + colOffset;
        const headerRow = row;
        const dataRow = row + 1;
        vals.forEach((_, i) => (ws.getCell(headerRow, startCol + 1 + i).value = `P${i + 1}`));
        ws.getCell(dataRow, startCol).value = label;
        const gapIndex = randInt(0, vals.length - 1);
        vals.forEach((v, i) => {
          if (i !== gapIndex) ws.getCell(dataRow, startCol + 1 + i).value = v;
        });
        // The gapped cell was genuinely never written -- the correct
        // extraction is an empty value there, not the unused random
        // number that was generated but deliberately not put on the
        // sheet.
        const expected = vals.map((v, i) => (i === gapIndex ? '' : v));
        truth.fields[f.name] = { anchor: { row: dataRow, col: startCol }, values: expected };
        row += 3;
      } else if (style === 6) {
        // Plain layout; whole file gets written out as .xls at the end.
        const startCol = 1 + colOffset;
        ws.getCell(row, startCol).value = label;
        vals.forEach((v, i) => (ws.getCell(row, startCol + 1 + i).value = v));
        truth.fields[f.name] = { anchor: { row, col: startCol }, values: vals };
        row += 2;
      } else if (style === 7 && f === FIELDS[0]) {
        // Mixed English label + a formula-error cell substituted in.
        const startCol = 1 + colOffset;
        ws.getCell(row, startCol).value = label;
        vals.forEach((v, i) => {
          if (i === 1) ws.getCell(row, startCol + 1 + i).value = { formula: '1/0', result: { error: '#DIV/0!' } };
          else ws.getCell(row, startCol + 1 + i).value = v;
        });
        const expected = vals.map((v, i) => (i === 1 ? '#DIV/0!' : v));
        truth.fields[f.name] = { anchor: { row, col: startCol }, values: expected };
        row += 2;
      } else {
        // style 0/1/7(other fields): plain row or column layout.
        const startCol = 1 + colOffset;
        if (style === 1) {
          ws.getCell(row, startCol).value = label;
          vals.forEach((v, i) => (ws.getCell(row + 1 + i, startCol).value = v));
          truth.fields[f.name] = { anchor: { row, col: startCol }, values: vals };
          row += vals.length + 2;
        } else {
          ws.getCell(row, startCol).value = label;
          vals.forEach((v, i) => (ws.getCell(row, startCol + 1 + i).value = v));
          truth.fields[f.name] = { anchor: { row, col: startCol }, values: vals };
          row += 2;
        }
      }
    }
  }

  const ext = style === 6 ? 'xls' : 'xlsx';
  const fileName = `${month}${style === 6 ? '_legacy' : ''}.${ext}`;
  const filePath = path.join(OUT_DIR, fileName);

  if (ext === 'xls') {
    const buf = await wb.xlsx.writeBuffer();
    const tmp = new ExcelJS.Workbook();
    await tmp.xlsx.load(buf);
    // round-trip through SheetJS to produce a genuine legacy .xls
    const wsData = [];
    tmp.getWorksheet('Sheet1').eachRow({ includeEmpty: true }, (r, rn) => {
      wsData[rn - 1] = [];
      r.eachCell({ includeEmpty: true }, (c, cn) => {
        wsData[rn - 1][cn - 1] = c.value;
      });
    });
    const legacyWb = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(legacyWb, sheet, 'Sheet1');
    XLSX.writeFile(legacyWb, filePath, { bookType: 'biff8' });
  } else {
    await wb.xlsx.writeFile(filePath);
  }

  truth.file = fileName;
  truth.style = style;
  return truth;
}

async function main() {
  fs.rmSync(OUT_DIR, { recursive: true, force: true });
  fs.mkdirSync(OUT_DIR, { recursive: true });

  const months = [];
  for (const year of [2024, 2025]) {
    for (let m = 1; m <= 12; m++) months.push(`${year}-${String(m).padStart(2, '0')}`);
  }

  const truths = [];
  for (let i = 0; i < months.length; i++) {
    truths.push(await buildFile(i, months[i]));
  }

  fs.writeFileSync(TRUTH_PATH, JSON.stringify({ fields: FIELDS, truths }, null, 2));
  console.log(`Generated ${truths.length} files in ${OUT_DIR}`);
  console.log(`Style distribution:`, truths.reduce((acc, t) => ((acc[t.style] = (acc[t.style] || 0) + 1), acc), {}));
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
