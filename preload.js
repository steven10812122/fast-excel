const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('fastExcel', {
  selectFolder: () => ipcRenderer.invoke('dialog:select-folder'),
  selectSavePath: () => ipcRenderer.invoke('dialog:select-save-path'),
  scanFolder: (folderPath, recipe, recursive, scanId, threshold) =>
    ipcRenderer.invoke('scan-folder', { folderPath, recipe, recursive, scanId, threshold }),
  cancelScan: (scanId) => ipcRenderer.send('scan-cancel', scanId),
  onScanProgress: (callback) => {
    const listener = (_event, data) => callback(data);
    ipcRenderer.on('scan-progress', listener);
    return () => ipcRenderer.removeListener('scan-progress', listener);
  },
  exportConsolidated: (results, recipe, outputPath) =>
    ipcRenderer.invoke('export-consolidated', { results, recipe, outputPath }),
  rescanField: (filePath, sheetName, address, recipe, fieldName) =>
    ipcRenderer.invoke('rescan-field', { filePath, sheetName, address, recipe, fieldName }),
  saveRecipe: (fields) => ipcRenderer.invoke('recipe:save', fields),
  loadRecipe: () => ipcRenderer.invoke('recipe:load'),
});
