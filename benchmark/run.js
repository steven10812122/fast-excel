// Runs the real matching engine (the exact same code path the app uses)
// against the 24-file generated batch, times it, and checks every single
// field/file result against the ground truth recorded at generation
// time -- not eyeballed, actually diffed.

const path = require('path');
const fs = require('fs');
const { listExcelFiles } = require('../src/engine/files');
const { scanFiles } = require('../src/engine/scan');

const FILES_DIR = path.join(__dirname, 'files');
const TRUTH_PATH = path.join(__dirname, 'truth.json');

function flatten(block) {
  return block.cells.flat().map((c) => c.value);
}

async function main() {
  const { fields, truths } = JSON.parse(fs.readFileSync(TRUTH_PATH, 'utf-8'));
  const recipe = { fields: fields.map((f) => ({ name: f.name, keywords: f.keywords })) };
  const truthByFile = new Map(truths.map((t) => [t.file, t]));

  const files = listExcelFiles(FILES_DIR, false);

  const started = process.hrtime.bigint();
  const { results, errors } = await scanFiles(files, recipe, undefined, {});
  const elapsedMs = Number(process.hrtime.bigint() - started) / 1e6;

  let totalFields = 0;
  let correctFields = 0;
  const mismatches = [];

  for (const result of results) {
    const fileName = path.basename(result.file);
    const truth = truthByFile.get(fileName);
    for (const field of fields) {
      totalFields++;
      const fr = result.fields[field.name];
      const expected = truth.fields[field.name].values;
      const actual = fr && fr.matched ? flatten(fr.block) : null;
      const ok = actual && JSON.stringify(actual) === JSON.stringify(expected);
      if (ok) correctFields++;
      else mismatches.push({ file: fileName, field: field.name, expected, actual });
    }
  }

  console.log('=== Fast Excel benchmark ===');
  console.log(`Files:          ${files.length}`);
  console.log(`Fields/file:    ${fields.length}`);
  console.log(`Read errors:    ${errors.length}`);
  console.log(`Time:           ${elapsedMs.toFixed(1)} ms (${(elapsedMs / files.length).toFixed(1)} ms/file)`);
  console.log(`Fields checked: ${totalFields}`);
  console.log(`Correct:        ${correctFields} (${((correctFields / totalFields) * 100).toFixed(1)}%)`);

  if (mismatches.length) {
    console.log('\nMismatches:');
    for (const m of mismatches) {
      console.log(` ${m.file} / ${m.field}: expected ${JSON.stringify(m.expected)}, got ${JSON.stringify(m.actual)}`);
    }
    process.exit(1);
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
