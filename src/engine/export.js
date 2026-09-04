// Turns the array of per-file match results into one consolidated workbook.

const ExcelJS = require('exceljs');
const path = require('path');

// A matched field is a rows x cols block. When every file's match for a
// field is a single row (rows <= 1 everywhere -- the common "one label,
// several period values" shape), it's spread into real numbered output
// columns instead of a joined string, so the destination report can still
// do further math on the values directly. A field that ever comes back as
// a genuine 2D block (rows > 1 somewhere) can't fit that "one row per
// file" shape, so it falls back to a single joined-text column: each row
// joined with ", " and rows stacked with " | ".
function fieldColumnPlan(fieldName, results) {
  let maxCols = 1;
  let hasMultiRow = false;
  for (const result of results) {
    const fr = result.fields[fieldName];
    if (!fr || !fr.matched) continue;
    if (fr.block.rows > 1) hasMultiRow = true;
    maxCols = Math.max(maxCols, fr.block.cols);
  }
  if (hasMultiRow) return { spread: false, width: 1 };
  return { spread: maxCols > 1, width: maxCols };
}

function joinedBlockText(block) {
  return block.cells.map((row) => row.map((c) => c.value).join(', ')).join(' | ');
}

async function exportConsolidated(results, recipe, outputPath, sourceFileLabel = 'Source file') {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Consolidated');

  const plans = new Map(recipe.fields.map((f) => [f.name, fieldColumnPlan(f.name, results)]));

  const columns = [{ header: sourceFileLabel, key: '__file', width: 32 }];
  for (const field of recipe.fields) {
    const plan = plans.get(field.name);
    if (plan.spread) {
      for (let i = 1; i <= plan.width; i++) {
        columns.push({ header: `${field.name} ${i}`, key: `${field.name}__${i}`, width: 16 });
      }
    } else {
      columns.push({ header: field.name, key: field.name, width: 20 });
    }
  }
  sheet.columns = columns;
  sheet.getRow(1).font = { bold: true };

  for (const result of results) {
    const row = { __file: path.basename(result.file) };
    for (const field of recipe.fields) {
      const plan = plans.get(field.name);
      const fr = result.fields[field.name];
      if (plan.spread) {
        const values = fr && fr.matched ? fr.block.cells[0].map((c) => c.value) : [];
        for (let i = 0; i < plan.width; i++) {
          row[`${field.name}__${i + 1}`] = values[i] ?? '';
        }
      } else {
        row[field.name] =
          fr && fr.matched
            ? fr.block.rows === 1 && fr.block.cols === 1
              ? fr.block.cells[0][0].value
              : joinedBlockText(fr.block)
            : '';
      }
    }
    sheet.addRow(row);
  }

  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}

module.exports = { exportConsolidated, fieldColumnPlan, joinedBlockText };
