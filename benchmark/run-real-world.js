// Runs the real matching engine against real SEC EDGAR 10-Q "Financial
// Report" Excel exports -- genuine multi-sheet (40-50+ sheets per file)
// public filings downloaded from sec.gov, not generated. Different
// companies use different terminology for the same line item ("Net
// sales" vs "Revenue" vs "Net Operating Revenues" vs "Revenues"), and
// each file has dozens of sheets to search across -- this is not a
// controlled test, results are reported as-is including any misses.

const path = require('path');
const { listDataFiles } = require('../src/engine/files');
const { scanFiles } = require('../src/engine/scan');

const FILES_DIR = path.join(__dirname, 'real-world');

const recipe = {
  fields: [
    { name: 'Revenue', keywords: ['Net sales', 'Revenue', 'Revenues', 'Net Operating Revenues', 'Total net sales', 'Total revenues'] },
    { name: 'Cost', keywords: ['Cost of sales', 'Cost of revenue', 'Cost of goods sold'] },
    { name: 'Gross Profit', keywords: ['Gross margin', 'Gross profit', 'Gross Profit'] },
    { name: 'Net Income', keywords: ['Net income', 'Net earnings'] },
  ],
};

async function main() {
  const files = listDataFiles(FILES_DIR, false);
  console.log(`Scanning ${files.length} real SEC EDGAR filings...\n`);

  const started = process.hrtime.bigint();
  const { results, errors } = await scanFiles(files, recipe, undefined, {});
  const elapsedMs = Number(process.hrtime.bigint() - started) / 1e6;

  let matched = 0;
  let total = 0;
  for (const result of results) {
    const fileName = path.basename(result.file);
    console.log(`--- ${fileName} ---`);
    for (const field of recipe.fields) {
      total++;
      const fr = result.fields[field.name];
      if (fr && fr.matched) {
        matched++;
        const vals = fr.block.cells[0].map((c) => c.value).join(', ');
        console.log(`  ${field.name}: [${vals}]  (sheet: "${fr.sheet}", header: "${fr.headerText}", conf: ${fr.score}${fr.ambiguous ? ', AMBIGUOUS' : ''})`);
      } else {
        console.log(`  ${field.name}: NOT FOUND`);
      }
    }
  }

  console.log(`\n=== Summary ===`);
  console.log(`Files:   ${files.length}`);
  console.log(`Errors:  ${errors.length}`);
  console.log(`Time:    ${elapsedMs.toFixed(0)} ms (${(elapsedMs / files.length).toFixed(0)} ms/file)`);
  console.log(`Matched: ${matched}/${total} (${((matched / total) * 100).toFixed(1)}%)`);
  if (errors.length) errors.forEach((e) => console.log('  error:', e.file, e.message));
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
