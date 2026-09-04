// Minimal i18n: a flat string table per language, a `t(key, ...args)`
// helper (calls the entry if it's a function, for strings that need
// interpolation), and a small localStorage-backed override on top of
// auto-detection from the OS/browser locale. No build step, no bundler --
// this is a plain script loaded before renderer.js, same as the rest of
// this app.

const STRINGS = {
  en: {
    tagline: "Match fields by keyword — it doesn't matter where the cells actually are.",
    recipeHeading: 'Field setup',
    importBtn: 'Import',
    importBtnTitle: 'Load a previously saved setup file',
    exportBtn: 'Export',
    exportBtnTitle: 'Save the current setup as a file to share or move to another computer',
    addFieldBtn: '+ Add field',
    fieldNameLabel: 'Output column name',
    fieldNamePlaceholder: 'e.g. Revenue',
    fieldKeywordsLabel: 'Keywords (comma-separated, synonyms supported)',
    fieldKeywordsPlaceholder: 'Revenue, Sales, Income',
    fieldHint:
      "Automatically figures out whether the data sits to the right of the label or below it, and stops at the first blank cell, a border, or another field's own match — no direction or cell count to set.",
    removeFieldBtn: 'Remove this field',
    thresholdLabel: 'Match strictness',
    selectFolderBtn: 'Choose folder',
    noFolderSelected: 'No folder selected',
    recursiveLabel: 'Include subfolders',
    scanBtn: 'Start scan',
    cancelBtn: 'Cancel',
    exportReportBtn: 'Export report',
    statusNeedField: 'Set up at least one field first (name + keywords) on the left.',
    statusNeedFolder: 'Choose a folder first.',
    statusScanning: 'Scanning…',
    statusCancelled: (n) => `Cancelled. Processed ${n} file(s).`,
    statusScanned: (found, ok) => `Found ${found} file(s), successfully processed ${ok}.`,
    statusScanFailedSuffix: (n) => ` ${n} file(s) failed to read (see the list below).`,
    statusScanFailed: (msg) => `Scan failed: ${msg}`,
    statusExporting: 'Exporting…',
    statusExported: (path) => `Exported to ${path}`,
    statusExportFailed: (msg) => `Export failed: ${msg}`,
    statusRecipeSaved: (path) => `Setup saved to ${path}`,
    statusRecipeSaveFailed: (msg) => `Save failed: ${msg}`,
    statusRecipeImported: 'Setup imported.',
    statusRecipeImportFailed: (msg) => `Import failed: ${msg}`,
    tableSourceFile: 'Source file',
    cellSkipped: 'Skipped',
    cellSkippedTitle: 'Marked to skip — click to change',
    cellNotFound: 'Not found',
    cellNotFoundTitle: 'No header matching the keywords was found in this file — click to specify one by hand',
    ambiguousNote: ' ⚠ Both directions found data; the direction was auto-picked — worth a check',
    cellTitle: (sheet, headerText, address, score, rows, cols, ambiguous) =>
      `Sheet: ${sheet} | Matched header: "${headerText}" (${address}) | Confidence: ${score} | Block size: ${rows} row(s) x ${cols} col(s)${ambiguous}`,
    previewTitle: 'Preview',
    previewNotFound:
      'No header matching the keywords was found in this file. You can specify a cell by hand below.',
    previewAmbiguous:
      " ⚠ Both directions found data — this was auto-picked, please confirm it's the block you want.",
    previewMeta: (sheet, headerText, address, score, rows, cols) =>
      `Sheet: ${sheet} | Matched header: "${headerText}" (${address}) | Confidence: ${score} | Block size: ${rows} row(s) x ${cols} col(s)`,
    emptyCell: '(empty)',
    overrideHint: 'If it matched the wrong thing, point it at the right cell by hand:',
    overrideSheetLabel: 'Sheet',
    overrideAddressLabel: 'Cell (e.g. C5)',
    overrideApplyBtn: 'Re-match',
    excludeLabel: "Don't export this field for this file (skip)",
    doneBtn: 'Done',
    overrideMissingFields: 'Please fill in both the sheet name and the cell address (e.g. C5).',
    overrideSheetNotFound: "Couldn't find that sheet — check the name.",
    overrideFailed: (msg) => `Re-match failed: ${msg}`,
    recipeErrorNotArray: 'This is not a Fast Excel setup file (the top level is not a list).',
    recipeErrorNoValidFields: "This setup file doesn't contain any valid field definitions.",
    recipeErrorInvalidJson: "This file isn't valid JSON.",
    langToggle: '中文',
    defaultFieldName: 'Revenue',
    defaultFieldKeywords: 'Revenue, Sales, Income',
  },
  zh: {
    tagline: '用關鍵字抓欄位,不用管每份檔案的格子長在哪',
    recipeHeading: '抓取設定',
    importBtn: '匯入',
    importBtnTitle: '讀取之前存的設定檔',
    exportBtn: '匯出',
    exportBtnTitle: '把目前設定存成檔案,可以分享或搬到別台電腦',
    addFieldBtn: '+ 新增欄位',
    fieldNameLabel: '輸出欄位名稱',
    fieldNamePlaceholder: '例如：營收',
    fieldKeywordsLabel: '關鍵字(逗號分隔,支援同義詞)',
    fieldKeywordsPlaceholder: '營收, 營業收入, revenue',
    fieldHint: '自動判斷資料在標籤的右邊還是下面,並抓到第一個空格、框線,或撞到別的欄位為止,不用指定方向或格數。',
    removeFieldBtn: '移除這個欄位',
    thresholdLabel: '比對嚴格度',
    selectFolderBtn: '選擇資料夾',
    noFolderSelected: '尚未選擇資料夾',
    recursiveLabel: '含子資料夾',
    scanBtn: '開始掃描',
    cancelBtn: '取消',
    exportReportBtn: '匯出報表',
    statusNeedField: '請先在左側設定至少一個欄位(名稱+關鍵字)。',
    statusNeedFolder: '請先選擇資料夾。',
    statusScanning: '掃描中…',
    statusCancelled: (n) => `已取消,已處理 ${n} 個檔案。`,
    statusScanned: (found, ok) => `找到 ${found} 個檔案,成功處理 ${ok} 個。`,
    statusScanFailedSuffix: (n) => ` ${n} 個檔案讀取失敗(詳見下方清單)。`,
    statusScanFailed: (msg) => `掃描失敗:${msg}`,
    statusExporting: '匯出中…',
    statusExported: (path) => `已匯出至 ${path}`,
    statusExportFailed: (msg) => `匯出失敗:${msg}`,
    statusRecipeSaved: (path) => `設定已存到 ${path}`,
    statusRecipeSaveFailed: (msg) => `存檔失敗:${msg}`,
    statusRecipeImported: '設定已匯入。',
    statusRecipeImportFailed: (msg) => `匯入失敗:${msg}`,
    tableSourceFile: '來源檔案',
    cellSkipped: '已略過',
    cellSkippedTitle: '已標記略過,點一下可以改',
    cellNotFound: '未找到',
    cellNotFoundTitle: '這個檔案裡找不到符合關鍵字的表頭,點一下可以手動指定',
    ambiguousNote: ' ⚠ 往右/往下兩個方向都有資料,方向是自動判斷的,建議點開確認',
    cellTitle: (sheet, headerText, address, score, rows, cols, ambiguous) =>
      `工作表:${sheet} | 命中表頭:「${headerText}」(${address}) | 相似度:${score} | 區塊大小:${rows} 列 x ${cols} 欄${ambiguous}`,
    previewTitle: '預覽',
    previewNotFound: '這個檔案裡找不到符合關鍵字的表頭。可以在下面手動指定儲存格重新抓取。',
    previewAmbiguous: ' ⚠ 往右跟往下兩個方向都抓到資料,這是自動判斷的結果,請確認是不是你要的區塊。',
    previewMeta: (sheet, headerText, address, score, rows, cols) =>
      `工作表:${sheet} ｜ 命中表頭:「${headerText}」(${address}) ｜ 相似度:${score} ｜ 區塊大小:${rows} 列 x ${cols} 欄`,
    emptyCell: '(空)',
    overrideHint: '如果抓錯了,手動指定正確的儲存格重新抓一次:',
    overrideSheetLabel: '工作表',
    overrideAddressLabel: '儲存格(例如 C5)',
    overrideApplyBtn: '重新抓取',
    excludeLabel: '這個檔案的這個欄位不匯出(略過)',
    doneBtn: '完成',
    overrideMissingFields: '請填工作表名稱跟儲存格位址(例如 C5)。',
    overrideSheetNotFound: '找不到這個工作表,請確認名稱正確。',
    overrideFailed: (msg) => `重新抓取失敗:${msg}`,
    recipeErrorNotArray: '這不是一份 Fast Excel 設定檔(最外層不是陣列)。',
    recipeErrorNoValidFields: '這份設定檔裡沒有找到任何有效的欄位設定。',
    recipeErrorInvalidJson: '這個檔案不是合法的 JSON 格式。',
    langToggle: 'English',
    defaultFieldName: '營收',
    defaultFieldKeywords: '營收, 營業收入, revenue',
  },
};

const LANG_KEY = 'fast-excel-lang';

function detectLanguage() {
  const nav = (navigator.language || 'en').toLowerCase();
  return nav.startsWith('zh') ? 'zh' : 'en';
}

function getLang() {
  try {
    const saved = localStorage.getItem(LANG_KEY);
    if (saved === 'en' || saved === 'zh') return saved;
  } catch (_) {
    // fall through to auto-detect
  }
  return detectLanguage();
}

function setLang(lang) {
  localStorage.setItem(LANG_KEY, lang);
}

function t(key, ...args) {
  const lang = getLang();
  const entry = (STRINGS[lang] && STRINGS[lang][key]) ?? STRINGS.en[key];
  if (entry === undefined) return key;
  return typeof entry === 'function' ? entry(...args) : entry;
}

// Applies every [data-i18n]/[data-i18n-title]/[data-i18n-placeholder]
// element's translation, and updates <html lang> for accessibility.
function applyStaticI18n() {
  document.documentElement.lang = getLang() === 'zh' ? 'zh-Hant' : 'en';
  document.querySelectorAll('[data-i18n]').forEach((el) => {
    el.textContent = t(el.dataset.i18n);
  });
  document.querySelectorAll('[data-i18n-title]').forEach((el) => {
    el.title = t(el.dataset.i18nTitle);
  });
  document.querySelectorAll('[data-i18n-placeholder]').forEach((el) => {
    el.placeholder = t(el.dataset.i18nPlaceholder);
  });
}

// This file is loaded as a plain <script> in the renderer (no bundler, no
// module system there), so `module` is undefined in the browser and this
// must stay a no-op there. Guarded like this, the exact same file can
// also be `require()`d from a plain Node test to check the two language
// tables stay in sync -- the actual risk with a hand-maintained dictionary
// like this one.
if (typeof module !== 'undefined') {
  module.exports = { STRINGS, t, getLang, setLang, detectLanguage };
}
