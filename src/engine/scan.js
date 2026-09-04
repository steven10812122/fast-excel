// The scan loop itself, decoupled from Electron's IPC/event plumbing so
// it can be unit tested directly: main.js supplies `shouldCancel` (reads
// a Set keyed by scan id) and `onProgress` (sends an IPC event), but the
// loop's own behavior -- cancel mid-batch, collect per-file errors,
// report a final count -- is what actually needs regression protection.

const { loadWorkbook } = require('./files');
const { processLoadedWorkbook } = require('./match');

async function scanFiles(files, recipe, threshold, { shouldCancel, onProgress } = {}) {
  const results = [];
  const errors = [];

  for (let i = 0; i < files.length; i++) {
    if (shouldCancel && shouldCancel()) break;
    const filePath = files[i];
    try {
      const workbook = await loadWorkbook(filePath);
      results.push(processLoadedWorkbook(workbook, filePath, recipe, threshold));
    } catch (err) {
      errors.push({ file: filePath, message: err.message });
    }
    if (onProgress) onProgress({ done: i + 1, total: files.length, file: filePath });
  }

  return {
    results,
    errors,
    fileCount: files.length,
    cancelled: Boolean(shouldCancel && shouldCancel()),
  };
}

module.exports = { scanFiles };
