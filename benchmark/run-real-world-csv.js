// Runs the real matching engine against a real CSV: Taiwan Stock
// Exchange's official quarterly operating-results open dataset, ~950
// companies in one flat table -- a genuinely different real-world shape
// than the SEC benchmark (one wide table with many rows, instead of a few
// key line items per file), and it has no borders or merged cells at all
// since CSV can't carry them, so it only exercises 3 of the 5 boundary
// signals.

const path = require('path');
const fs = require('fs');
const { loadWorkbook } = require('../src/engine/files');
const { processLoadedWorkbook } = require('../src/engine/match');

const FILE = path.join(__dirname, 'real-world', 'tw-listed-companies.csv');

async function main() {
  if (!fs.existsSync(FILE)) {
    console.error('Run `node benchmark/fetch-real-world-csv.js` first.');
    process.exit(1);
  }

  const totalRows = fs.readFileSync(FILE, 'utf-8').trim().split('\n').length - 1;

  const recipe = {
    fields: [
      { name: '營業收入', keywords: ['營業收入'] },
      { name: '毛利率', keywords: ['毛利率'] },
      { name: '營業利益率', keywords: ['營業利益率'] },
      { name: '稅後純益率', keywords: ['稅後純益率'] },
    ],
  };

  const started = process.hrtime.bigint();
  const workbook = await loadWorkbook(FILE);
  const result = processLoadedWorkbook(workbook, FILE, recipe);
  const elapsedMs = Number(process.hrtime.bigint() - started) / 1e6;

  console.log(`=== Real CSV: ${totalRows} companies, ${recipe.fields.length} fields ===`);
  console.log(`Time: ${elapsedMs.toFixed(1)} ms\n`);

  for (const field of recipe.fields) {
    const fr = result.fields[field.name];
    if (fr && fr.matched) {
      console.log(
        `${field.name}: matched ${fr.block.rows} of ${totalRows} rows (header "${fr.headerText}", confidence ${fr.score})` +
          (fr.block.rows < totalRows ? ` -- stopped early, see README for why` : '')
      );
    } else {
      console.log(`${field.name}: NOT FOUND`);
    }
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
