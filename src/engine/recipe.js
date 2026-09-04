// Validates a recipe JSON file before it's handed to the renderer. The
// file comes from `dialog.showOpenDialog` -- any file on disk the user
// points at, not something this app wrote itself -- so a malformed or
// hand-edited file must be rejected cleanly instead of getting to
// `renderFieldList()` and breaking on a missing `.trim()` etc.
//
// Returns an `errorCode` rather than a message: this runs in the main
// process, which has no notion of the renderer's active UI language, so
// the renderer maps the code to a localized string itself.
function validateRecipeFields(raw) {
  if (!Array.isArray(raw)) {
    return { valid: false, errorCode: 'not_array' };
  }

  const fields = raw.filter(
    (item) =>
      item &&
      typeof item === 'object' &&
      typeof item.name === 'string' &&
      typeof item.keywords === 'string'
  );

  if (fields.length === 0) {
    return { valid: false, errorCode: 'no_valid_fields' };
  }

  return { valid: true, fields };
}

module.exports = { validateRecipeFields };
