const STORAGE_KEY = 'fast-excel-recipe-v1';

const state = {
  fields: loadRecipe(),
  folderPath: null,
  threshold: 0.72,
  lastResults: null,
  lastRecipe: null,
  excluded: new Set(), // `${fileIndex}::${fieldName}` -- excluded from export
};

let currentScanId = null;
let offProgress = null;
let previewState = null; // { fileIndex, fieldName }

const fieldListEl = document.getElementById('field-list');
const addFieldBtn = document.getElementById('add-field-btn');
const loadRecipeBtn = document.getElementById('load-recipe-btn');
const saveRecipeBtn = document.getElementById('save-recipe-btn');
const thresholdInput = document.getElementById('threshold-input');
const thresholdValue = document.getElementById('threshold-value');
const selectFolderBtn = document.getElementById('select-folder-btn');
const folderPathEl = document.getElementById('folder-path');
const recursiveCheckbox = document.getElementById('recursive-checkbox');
const scanBtn = document.getElementById('scan-btn');
const cancelScanBtn = document.getElementById('cancel-scan-btn');
const exportBtn = document.getElementById('export-btn');
const progressWrap = document.getElementById('progress-wrap');
const progressFill = document.getElementById('progress-fill');
const progressText = document.getElementById('progress-text');
const statusLine = document.getElementById('status-line');
const errorListEl = document.getElementById('error-list');
const theadEl = document.getElementById('results-thead');
const tbodyEl = document.getElementById('results-tbody');

const previewModal = document.getElementById('preview-modal');
const previewTitle = document.getElementById('preview-title');
const previewMeta = document.getElementById('preview-meta');
const previewGrid = document.getElementById('preview-grid');
const overrideSheet = document.getElementById('override-sheet');
const overrideAddress = document.getElementById('override-address');
const overrideApplyBtn = document.getElementById('override-apply-btn');
const excludeCheckbox = document.getElementById('exclude-checkbox');
const previewCloseBtn = document.getElementById('preview-close-btn');
const previewDoneBtn = document.getElementById('preview-done-btn');

function loadRecipe() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (_) {
    // fall through to default
  }
  return [{ name: '營收', keywords: '營收, 營業收入, revenue' }];
}

function saveRecipe() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.fields));
}

function renderFieldList() {
  fieldListEl.innerHTML = '';
  state.fields.forEach((field, index) => {
    const card = document.createElement('div');
    card.className = 'field-card';
    card.innerHTML = `
      <div class="field-card-row">
        <div style="flex:1">
          <label>輸出欄位名稱</label>
          <input type="text" data-key="name" value="${escapeAttr(field.name)}" placeholder="例如：營收" />
        </div>
      </div>
      <div class="field-card-row">
        <div style="flex:1">
          <label>關鍵字(逗號分隔,支援同義詞)</label>
          <textarea data-key="keywords" placeholder="營收, 營業收入, revenue">${escapeHtml(field.keywords)}</textarea>
        </div>
      </div>
      <p class="field-card-hint">自動判斷資料在標籤的右邊還是下面,並抓到第一個空格、框線,或撞到別的欄位為止,不用指定方向或格數。</p>
      <div class="field-card-row">
        <button class="field-card-remove" data-action="remove">移除這個欄位</button>
      </div>
    `;

    card.querySelectorAll('[data-key]').forEach((input) => {
      input.addEventListener('input', () => {
        const key = input.dataset.key;
        state.fields[index][key] = input.value;
        saveRecipe();
      });
    });

    card.querySelector('[data-action="remove"]').addEventListener('click', () => {
      state.fields.splice(index, 1);
      saveRecipe();
      renderFieldList();
    });

    fieldListEl.appendChild(card);
  });
}

function escapeHtml(str) {
  return String(str || '').replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
}

function escapeAttr(str) {
  return escapeHtml(str).replace(/"/g, '&quot;');
}

addFieldBtn.addEventListener('click', () => {
  state.fields.push({ name: '', keywords: '' });
  saveRecipe();
  renderFieldList();
});

saveRecipeBtn.addEventListener('click', async () => {
  try {
    const savedPath = await window.fastExcel.saveRecipe(state.fields);
    if (savedPath) statusLine.textContent = `設定已存到 ${savedPath}`;
  } catch (err) {
    statusLine.textContent = `存檔失敗:${err.message}`;
  }
});

loadRecipeBtn.addEventListener('click', async () => {
  try {
    const fields = await window.fastExcel.loadRecipe();
    if (!fields) return;
    state.fields = fields;
    saveRecipe();
    renderFieldList();
    statusLine.textContent = '設定已匯入。';
  } catch (err) {
    statusLine.textContent = `匯入失敗:${err.message}`;
  }
});

thresholdInput.addEventListener('input', () => {
  state.threshold = Number(thresholdInput.value);
  thresholdValue.textContent = state.threshold.toFixed(2);
});

selectFolderBtn.addEventListener('click', async () => {
  const folder = await window.fastExcel.selectFolder();
  if (!folder) return;
  state.folderPath = folder;
  folderPathEl.textContent = folder;
  scanBtn.disabled = false;
});

function currentRecipe() {
  const fields = state.fields
    .filter((f) => f.name.trim() && f.keywords.trim())
    .map((f) => ({
      name: f.name.trim(),
      keywords: f.keywords.split(',').map((k) => k.trim()).filter(Boolean),
    }));
  return { fields };
}

scanBtn.addEventListener('click', async () => {
  const recipe = currentRecipe();
  if (recipe.fields.length === 0) {
    statusLine.textContent = '請先在左側設定至少一個欄位(名稱+關鍵字)。';
    return;
  }
  if (!state.folderPath) {
    statusLine.textContent = '請先選擇資料夾。';
    return;
  }

  scanBtn.disabled = true;
  exportBtn.disabled = true;
  cancelScanBtn.hidden = false;
  errorListEl.hidden = true;
  errorListEl.innerHTML = '';
  progressWrap.hidden = false;
  progressFill.style.width = '0%';
  progressText.textContent = '';
  statusLine.textContent = '掃描中…';

  const scanId = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  currentScanId = scanId;
  if (offProgress) offProgress();
  offProgress = window.fastExcel.onScanProgress((data) => {
    if (data.scanId !== scanId) return;
    const pct = data.total ? Math.round((data.done / data.total) * 100) : 0;
    progressFill.style.width = `${pct}%`;
    progressText.textContent = `${data.done} / ${data.total}(${data.file}）`;
  });

  try {
    const { results, errors, fileCount, cancelled } = await window.fastExcel.scanFolder(
      state.folderPath,
      recipe,
      recursiveCheckbox.checked,
      scanId,
      state.threshold
    );

    state.lastResults = results;
    state.lastRecipe = recipe;
    state.excluded = new Set();
    renderResultsTable(results, recipe);

    let msg = cancelled
      ? `已取消,已處理 ${results.length} 個檔案。`
      : `找到 ${fileCount} 個檔案,成功處理 ${results.length} 個。`;
    if (errors.length) {
      msg += ` ${errors.length} 個檔案讀取失敗(詳見下方清單)。`;
      errorListEl.hidden = false;
      errorListEl.innerHTML = errors
        .map((e) => `<div>${escapeHtml(e.file.split(/[\\/]/).pop())}${escapeHtml(':' + e.message)}</div>`)
        .join('');
    }
    statusLine.textContent = msg;
    exportBtn.disabled = results.length === 0;
  } catch (err) {
    statusLine.textContent = `掃描失敗:${err.message}`;
  } finally {
    scanBtn.disabled = false;
    cancelScanBtn.hidden = true;
    progressWrap.hidden = true;
    if (offProgress) {
      offProgress();
      offProgress = null;
    }
    currentScanId = null;
  }
});

cancelScanBtn.addEventListener('click', () => {
  if (currentScanId) window.fastExcel.cancelScan(currentScanId);
});

function formatBlock(block) {
  if (block.rows === 1 && block.cols === 1) return String(block.cells[0][0].value);
  return block.cells.map((row) => row.map((c) => c.value).join(', ')).join(' | ');
}

function renderResultsTable(results, recipe) {
  theadEl.innerHTML = '';
  tbodyEl.innerHTML = '';

  const headRow = document.createElement('tr');
  headRow.innerHTML =
    '<th>來源檔案</th>' + recipe.fields.map((f) => `<th>${escapeHtml(f.name)}</th>`).join('');
  theadEl.appendChild(headRow);

  results.forEach((result, fileIndex) => {
    const row = document.createElement('tr');
    const fileName = result.file.split(/[/\\]/).pop();
    let html = `<td class="cell-file">${escapeHtml(fileName)}</td>`;

    for (const field of recipe.fields) {
      const excludeKey = `${fileIndex}::${field.name}`;
      const fr = result.fields[field.name];

      if (state.excluded.has(excludeKey)) {
        html += `<td><span class="cell-unmatched cell-clickable" data-file-index="${fileIndex}" data-field="${escapeAttr(field.name)}" title="已標記略過,點一下可以改">已略過</span></td>`;
      } else if (fr && fr.matched) {
        const values = formatBlock(fr.block);
        const ambiguousClass = fr.ambiguous ? ' cell-ambiguous' : '';
        const ambiguousNote = fr.ambiguous ? ' ⚠ 往右/往下兩個方向都有資料,方向是自動判斷的,建議點開確認' : '';
        const title = `工作表:${fr.sheet} | 命中表頭:「${fr.headerText}」(${fr.headerAddress}) | 相似度:${fr.score} | 區塊大小:${fr.block.rows} 列 x ${fr.block.cols} 欄${ambiguousNote}`;
        html += `<td><span class="cell-matched cell-clickable${ambiguousClass}" title="${escapeAttr(title)}" data-file-index="${fileIndex}" data-field="${escapeAttr(field.name)}">${escapeHtml(values)}</span></td>`;
      } else {
        html += `<td><span class="cell-unmatched cell-clickable" title="這個檔案裡找不到符合關鍵字的表頭,點一下可以手動指定" data-file-index="${fileIndex}" data-field="${escapeAttr(field.name)}">未找到</span></td>`;
      }
    }

    row.innerHTML = html;
    tbodyEl.appendChild(row);
  });

  tbodyEl.querySelectorAll('[data-file-index]').forEach((el) => {
    el.addEventListener('click', () => {
      openPreview(Number(el.dataset.fileIndex), el.dataset.field);
    });
  });
}

function openPreview(fileIndex, fieldName) {
  previewState = { fileIndex, fieldName };
  const result = state.lastResults[fileIndex];
  const fr = result.fields[fieldName];
  const fileName = result.file.split(/[/\\]/).pop();

  previewTitle.textContent = `${fileName} — ${fieldName}`;
  const fallbackSheet = Object.values(result.fields).find((f) => f && f.sheet)?.sheet || '';
  overrideSheet.value = fr && fr.sheet ? fr.sheet : fallbackSheet;
  overrideAddress.value = fr && fr.matched ? fr.headerAddress : '';
  excludeCheckbox.checked = state.excluded.has(`${fileIndex}::${fieldName}`);

  renderPreviewContent(fr);
  previewModal.hidden = false;
}

function renderPreviewContent(fr) {
  if (!fr || !fr.matched) {
    previewMeta.innerHTML =
      '<span class="warn">這個檔案裡找不到符合關鍵字的表頭。可以在下面手動指定儲存格重新抓取。</span>';
    previewGrid.innerHTML = '';
    return;
  }
  const ambiguousNote = fr.ambiguous
    ? '<div class="warn">⚠ 往右跟往下兩個方向都抓到資料,這是自動判斷的結果,請確認是不是你要的區塊。</div>'
    : '';
  previewMeta.innerHTML = `
    工作表:${escapeHtml(fr.sheet)} ｜ 命中表頭:「${escapeHtml(fr.headerText)}」(${fr.headerAddress}) ｜
    相似度:${fr.score} ｜ 區塊大小:${fr.block.rows} 列 x ${fr.block.cols} 欄
    ${ambiguousNote}
  `;
  previewGrid.innerHTML = fr.block.cells
    .map(
      (row) =>
        '<tr>' +
        row
          .map((c) => {
            const isEmpty = c.value === '' || c.value === null || c.value === undefined;
            return `<td class="${isEmpty ? 'empty-cell' : ''}" title="${escapeAttr(c.address)}">${
              isEmpty ? '(空)' : escapeHtml(String(c.value))
            }</td>`;
          })
          .join('') +
        '</tr>'
    )
    .join('');
}

overrideApplyBtn.addEventListener('click', async () => {
  if (!previewState) return;
  const { fileIndex, fieldName } = previewState;
  const result = state.lastResults[fileIndex];
  const sheetName = overrideSheet.value.trim();
  const address = overrideAddress.value.trim();
  if (!sheetName || !address) {
    previewMeta.innerHTML = '<span class="warn">請填工作表名稱跟儲存格位址(例如 C5)。</span>';
    return;
  }

  overrideApplyBtn.disabled = true;
  try {
    const fr = await window.fastExcel.rescanField(
      result.file,
      sheetName,
      address,
      state.lastRecipe,
      fieldName
    );
    if (!fr) {
      previewMeta.innerHTML = '<span class="warn">找不到這個工作表,請確認名稱正確。</span>';
      return;
    }
    result.fields[fieldName] = fr;
    renderPreviewContent(fr);
    renderResultsTable(state.lastResults, state.lastRecipe);
  } catch (err) {
    previewMeta.innerHTML = `<span class="warn">重新抓取失敗:${escapeHtml(err.message)}</span>`;
  } finally {
    overrideApplyBtn.disabled = false;
  }
});

excludeCheckbox.addEventListener('change', () => {
  if (!previewState) return;
  const key = `${previewState.fileIndex}::${previewState.fieldName}`;
  if (excludeCheckbox.checked) state.excluded.add(key);
  else state.excluded.delete(key);
});

function closePreview() {
  previewModal.hidden = true;
  previewState = null;
  if (state.lastResults) renderResultsTable(state.lastResults, state.lastRecipe);
}

previewCloseBtn.addEventListener('click', closePreview);
previewDoneBtn.addEventListener('click', closePreview);

exportBtn.addEventListener('click', async () => {
  if (!state.lastResults) return;
  const outputPath = await window.fastExcel.selectSavePath();
  if (!outputPath) return;

  // Excluded file/field pairs are dropped just for this export -- clone
  // so the on-screen results (and the excluded marks) are untouched.
  const payload = state.lastResults.map((result, fileIndex) => {
    const fields = {};
    for (const [name, fr] of Object.entries(result.fields)) {
      fields[name] = state.excluded.has(`${fileIndex}::${name}`) ? { matched: false } : fr;
    }
    return { file: result.file, fields };
  });

  statusLine.textContent = '匯出中…';
  exportBtn.disabled = true;
  try {
    await window.fastExcel.exportConsolidated(payload, state.lastRecipe, outputPath);
    statusLine.textContent = `已匯出至 ${outputPath}`;
  } catch (err) {
    statusLine.textContent = `匯出失敗:${err.message}`;
  } finally {
    exportBtn.disabled = false;
  }
});

renderFieldList();
