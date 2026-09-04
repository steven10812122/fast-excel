const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { autoUpdater } = require('electron-updater');

const { extractManual, findHeaderCell } = require('./src/engine/match');
const { exportConsolidated } = require('./src/engine/export');
const { listExcelFiles, loadWorkbook } = require('./src/engine/files');
const { scanFiles } = require('./src/engine/scan');
const { validateRecipeFields } = require('./src/engine/recipe');

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
  // just log and move on if the feed isn't reachable or isn't set up yet.
  if (app.isPackaged) {
    autoUpdater.checkForUpdatesAndNotify().catch((err) => {
      console.error('autoUpdater check failed (non-fatal):', err.message);
    });
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

const cancelledScans = new Set();

ipcMain.handle('scan-folder', async (event, { folderPath, recipe, recursive, scanId, threshold }) => {
  const files = listExcelFiles(folderPath, recursive);
  const outcome = await scanFiles(files, recipe, threshold, {
    shouldCancel: () => cancelledScans.has(scanId),
    onProgress: ({ done, total, file }) => {
      event.sender.send('scan-progress', { scanId, done, total, file: path.basename(file) });
    },
  });
  cancelledScans.delete(scanId);
  return outcome;
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

  let parsed;
  try {
    parsed = JSON.parse(fs.readFileSync(result.filePaths[0], 'utf-8'));
  } catch (err) {
    return { valid: false, error: '這個檔案不是合法的 JSON 格式。' };
  }
  return validateRecipeFields(parsed);
});
