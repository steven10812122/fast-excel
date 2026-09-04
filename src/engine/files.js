// File discovery and workbook loading -- deliberately Electron-free so it
// can be unit tested with plain `node`, unlike main.js which also wires
// these into ipcMain/dialog.

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

function listExcelFiles(folderPath, recursive) {
  const found = [];
  function walk(dir) {
    for (const name of fs.readdirSync(dir)) {
      if (name.startsWith('~$') || name.startsWith('.')) continue; // Excel lock files, dotfiles
      const full = path.join(dir, name);
      const stat = fs.statSync(full);
      if (stat.isDirectory()) {
        if (recursive) walk(full);
        continue;
      }
      if (/\.xlsx?$/i.test(name)) found.push(full);
    }
  }
  walk(folderPath);
  return found;
}

// Legacy .xls files (Excel 97-2003 binary format) aren't readable by
// exceljs at all. Convert them to an in-memory .xlsx first via SheetJS,
// then hand that to exceljs as normal so every downstream rule (blank/
// border/label-line/merge detection) keeps working unchanged. Some
// formatting fidelity (older border styles in particular) can be lost in
// that round-trip -- .xls files fall back to blank/other-anchor detection
// more often than a native .xlsx would.
async function loadWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  if (/\.xls$/i.test(filePath)) {
    const legacy = XLSX.readFile(filePath, { cellStyles: true });
    const buffer = XLSX.write(legacy, { type: 'buffer', bookType: 'xlsx' });
    await workbook.xlsx.load(buffer);
  } else {
    await workbook.xlsx.readFile(filePath);
  }
  return workbook;
}

module.exports = { listExcelFiles, loadWorkbook };
