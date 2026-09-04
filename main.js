const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { autoUpdater } = require('electron-updater');

const { processLoadedWorkbook, extractManual, findHeaderCell } = require('./src/engine/match');
const { exportConsolidated } = require('./src/engine/export');

function createWindow() {
  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.loadFile(path.join(__dirname, 'src/renderer/index.html'));
}

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });

  // Best-effort only: there's no GitHub release to check against until
  // this repo's package.json "publish" owner/repo is filled in and a
  // release is actually published, so this must never crash startup --
  // just silently skip if the feed isn't reachable or isn't set up yet.
  if (app.isPackaged) {
    autoUpdater.checkForUpdatesAndNotify().catch(() => {});
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

ipcMain.handle('dialog:select-folder', async () => {
  const result = await dialog.showOpenDialog({ properties: ['openDirectory'] });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

ipcMain.handle('dialog:select-save-path', async () => {
  const result = await dialog.showSaveDialog({
    defaultPath: 'consolidated.xlsx',
    filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }],
  });
  if (result.canceled || !result.filePath) return null;
  return result.filePath;
});

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

const cancelledScans = new Set();

ipcMain.handle('scan-folder', async (event, { folderPath, recipe, recursive, scanId, threshold }) => {
  const files = listExcelFiles(folderPath, recursive);
  const results = [];
  const errors = [];

  for (let i = 0; i < files.length; i++) {
    if (cancelledScans.has(scanId)) break;
    const filePath = files[i];
    try {
      const workbook = await loadWorkbook(filePath);
      results.push(processLoadedWorkbook(workbook, filePath, recipe, threshold));
    } catch (err) {
      errors.push({ file: filePath, message: err.message });
    }
    event.sender.send('scan-progress', {
      scanId,
      done: i + 1,
      total: files.length,
      file: path.basename(filePath),
    });
  }

  const wasCancelled = cancelledScans.has(scanId);
  cancelledScans.delete(scanId);
  return { results, errors, fileCount: files.length, cancelled: wasCancelled };
});

ipcMain.on('scan-cancel', (_event, scanId) => {
  cancelledScans.add(scanId);
});

ipcMain.handle('export-consolidated', async (_event, { results, recipe, outputPath }) => {
  await exportConsolidated(results, recipe, outputPath);
  return true;
});

// Manual override: re-extract one field, for one file, anchored at a
// cell the user picked by hand instead of the auto-matched keyword hit.
// Still respects every other field's anchor on that sheet as a boundary.
ipcMain.handle('rescan-field', async (_event, { filePath, sheetName, address, recipe, fieldName }) => {
  const workbook = await loadWorkbook(filePath);
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) return null;

  const otherAnchors = new Set();
  for (const field of recipe.fields) {
    if (field.name === fieldName) continue;
    const match = findHeaderCell(worksheet, field.keywords);
    if (match) otherAnchors.add(match.address);
  }

  return extractManual(worksheet, address, otherAnchors);
});

ipcMain.handle('recipe:save', async (_event, fields) => {
  const result = await dialog.showSaveDialog({
    defaultPath: 'fast-excel-recipe.json',
    filters: [{ name: 'Fast Excel 抓取設定', extensions: ['json'] }],
  });
  if (result.canceled || !result.filePath) return null;
  fs.writeFileSync(result.filePath, JSON.stringify(fields, null, 2), 'utf-8');
  return result.filePath;
});

ipcMain.handle('recipe:load', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [{ name: 'Fast Excel 抓取設定', extensions: ['json'] }],
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  const raw = fs.readFileSync(result.filePaths[0], 'utf-8');
  return JSON.parse(raw);
});
