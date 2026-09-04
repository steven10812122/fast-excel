const ExcelJS = require('exceljs');

// Build an in-memory workbook from a plain description so tests don't
// touch disk. `sheets` is { sheetName: { 'A1': value, ... } }.
// `mergesBySheet` is { sheetName: ['A1:D1', ...] }.
function buildWorkbook(sheets, mergesBySheet = {}) {
  const wb = new ExcelJS.Workbook();
  for (const [sheetName, cells] of Object.entries(sheets)) {
    const ws = wb.addWorksheet(sheetName);
    for (const [address, value] of Object.entries(cells)) {
      ws.getCell(address).value = value;
    }
    for (const range of mergesBySheet[sheetName] || []) {
      ws.mergeCells(range);
    }
  }
  return wb;
}

function setBorder(wb, sheetName, address, border) {
  wb.getWorksheet(sheetName).getCell(address).border = border;
}

module.exports = { buildWorkbook, setBorder };
