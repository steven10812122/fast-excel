// Validates a recipe JSON file before it's handed to the renderer. The
// file comes from `dialog.showOpenDialog` -- any file on disk the user
// points at, not something this app wrote itself -- so a malformed or
// hand-edited file must be rejected cleanly instead of getting to
// `renderFieldList()` and breaking on a missing `.trim()` etc.
function validateRecipeFields(raw) {
  if (!Array.isArray(raw)) {
    return { valid: false, error: '這不是一份 Fast Excel 設定檔(最外層不是陣列)。' };
  }

  const fields = raw.filter(
    (item) =>
      item &&
      typeof item === 'object' &&
      typeof item.name === 'string' &&
      typeof item.keywords === 'string'
  );

  if (fields.length === 0) {
    return { valid: false, error: '這份設定檔裡沒有找到任何有效的欄位設定。' };
  }

  return { valid: true, fields };
}

module.exports = { validateRecipeFields };
