// Core matching engine: find a header cell by fuzzy keyword match, then pull
// a range of values relative to it. No Electron/DOM dependency so it can be
// unit tested with plain `node`.

const ExcelJS = require('exceljs');

function normalize(text) {
  if (text === null || text === undefined) return '';
  return String(text)
    .toLowerCase()
    .replace(/[\s　]+/g, '') // strip normal + full-width spaces
    .replace(/[，,。.：:；;（）()【】\[\]\-_/\\]/g, '');
}

// Levenshtein distance, used to score near-miss keyword matches
// (typos, extra characters, minor rewording).
function levenshtein(a, b) {
  if (a === b) return 0;
  const al = a.length;
  const bl = b.length;
  if (al === 0) return bl;
  if (bl === 0) return al;

  let prev = new Array(bl + 1);
  let curr = new Array(bl + 1);
  for (let j = 0; j <= bl; j++) prev[j] = j;

  for (let i = 1; i <= al; i++) {
    curr[0] = i;
    for (let j = 1; j <= bl; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(prev[j] + 1, curr[j - 1] + 1, prev[j - 1] + cost);
    }
    [prev, curr] = [curr, prev];
  }
  return prev[bl];
}

function similarity(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (a.includes(b) || b.includes(a)) {
    const shorter = Math.min(a.length, b.length);
    const longer = Math.max(a.length, b.length);
    return 0.85 + 0.15 * (shorter / longer);
  }
  const dist = levenshtein(a, b);
  const maxLen = Math.max(a.length, b.length);
  return 1 - dist / maxLen;
}

// Best score for a raw cell string against a set of keyword synonyms.
function scoreCell(cellText, keywords) {
  const normCell = normalize(cellText);
  if (!normCell) return 0;
  let best = 0;
  for (const kw of keywords) {
    const s = similarity(normCell, normalize(kw));
    if (s > best) best = s;
  }
  return best;
}

const DEFAULT_THRESHOLD = 0.72;

// Scan every used cell in a worksheet, return the single best match for
// `keywords` above `threshold`, or null if nothing qualifies.
function findHeaderCell(worksheet, keywords, threshold = DEFAULT_THRESHOLD) {
  let best = null;
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const text = cellText(cell);
      if (!text) return;
      const score = scoreCell(text, keywords);
      if (score >= threshold && (!best || score > best.score)) {
        best = {
          row: rowNumber,
          col: colNumber,
          address: cell.address,
          text,
          score,
        };
      }
    });
  });
  return best;
}

// A formula's `result` is either the computed value, or -- if the formula
// itself errors -- an object like `{ error: '#DIV/0!' }`. Left unhandled,
// that object leaks through as `[object Object]` in text, and doesn't
// count as blank, so it would silently be trusted as a real data cell.
// Surfacing the error string itself instead is both correct (it can't
// accidentally fuzzy-match a keyword or be mistaken for a real header
// line) and honest -- it shows up in the export as "#DIV/0!" so a broken
// source formula stays visible rather than disappearing.
function formulaResultToPrimitive(result) {
  if (result && typeof result === 'object' && 'error' in result) return result.error;
  return result;
}

function cellText(cell) {
  const v = cell.value;
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if (v.text) return String(v.text);
    if (v.result !== undefined) return String(formulaResultToPrimitive(v.result));
    return '';
  }
  return String(v);
}

function cellRawValue(cell) {
  const v = cell.value;
  if (v && typeof v === 'object') {
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if (v.text) return v.text;
    if (v.result !== undefined) return formulaResultToPrimitive(v.result);
    return '';
  }
  return v === null || v === undefined ? '' : v;
}

function isBlank(raw) {
  return raw === '' || raw === null || raw === undefined;
}

function colLetterToNumber(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function parseRange(rangeStr) {
  const [start, end] = rangeStr.split(':');
  const parseCell = (addr) => {
    const m = addr.match(/^([A-Z]+)(\d+)$/);
    return { row: Number(m[2]), col: colLetterToNumber(m[1]) };
  };
  const a = parseCell(start);
  const b = parseCell(end);
  return {
    top: Math.min(a.row, b.row),
    bottom: Math.max(a.row, b.row),
    left: Math.min(a.col, b.col),
    right: Math.max(a.col, b.col),
  };
}

// If the anchor cell is merged (e.g. a title merged across A1:D1, centered,
// so the label is written once instead of once per column), the merge's
// own span is a far more authoritative width/height signal than scanning
// for blanks: it's the report author explicitly stating how many columns
// or rows that label covers, and it stays correct even if a cell in the
// data block below happens to be blank for one file.
function findMergeSpan(worksheet, row, col) {
  for (const rangeStr of worksheet.model.merges || []) {
    const range = parseRange(rangeStr);
    if (row >= range.top && row <= range.bottom && col >= range.left && col <= range.right) {
      return range;
    }
  }
  return null;
}

// Deliberate divider lines (medium/thick/double) mark a table edge that a
// human drew on purpose. Thin/hair borders are excluded because they're
// routinely applied as generic grid styling across whole sheets, including
// past the real table -- using them would cause false stops everywhere.
const BOUNDARY_BORDER_STYLES = new Set(['medium', 'thick', 'double']);

function isBoundaryBorder(borderSide) {
  return !!borderSide && BOUNDARY_BORDER_STYLES.has(borderSide.style);
}

// Walk from (row, col) along `axis` ('col' to scan rightward across a row,
// 'row' to scan downward across a column), counting cells until: a blank
// cell is hit, a divider border (medium/thick/double) marks the edge, or
// `reference` says this position has fallen off the header/label line that
// vouches for it. Any one of the three ends the run -- together they let
// the block's size be fully self-determined, with no fixed count or manual
// cap, even when an unrelated block sits immediately adjacent with no
// blank gap and no border between them.
function measureRun(worksheet, row, col, axis, reference, otherAnchors, tolerateBlanks) {
  let n = 0;
  let prevCell = null;
  let blanksTolerated = 0;
  // The perpendicular/"guard" axis (buildLabelGuard, not buildReference)
  // always treats its own n=0 position as trivially valid -- for a
  // 'right' block that's the anchor's own row, for a 'down' block that's
  // the anchor's own column -- which is exactly what buildLabelGuard's
  // own `r === anchorRow ||` / `c === anchorCol ||` clause encodes. But
  // if the *other* axis happens to have tolerated a gap at that exact
  // (row, col), this cell can be genuinely blank -- and without this,
  // that single blank at n=0 would kill the whole block (both axes
  // start from the same point), even though the primary axis already
  // correctly determined real data exists in this row/column elsewhere.
  const isGuardAxis = !tolerateBlanks && Boolean(reference);
  for (;;) {
    const r = axis === 'row' ? row + n : row;
    const c = axis === 'col' ? col + n : col;
    const cell = worksheet.getCell(r, c);
    const blank = isBlank(cellRawValue(cell));

    if (reference && !reference(r, c)) break;

    if (blank && !(n === 0 && isGuardAxis)) {
      // A single missing value (this file's number for one period, say)
      // doesn't necessarily mean the table ended -- if a header/label line
      // is actively vouching that a value belongs at this exact position
      // (only true for the primary reference line, not the perpendicular
      // label guard, which merely means "no other field has claimed this
      // yet"), trust it and keep going, recording an empty cell here. Only
      // one such gap is bridged per block. It counts as genuinely
      // isolated -- as opposed to the table having actually ended one
      // cell earlier -- when either the next cell has real data, or
      // there's no next cell to have data in at all (the header/label
      // line itself stops right there too, so this was always going to
      // be the last column/row regardless). Two real gaps in a row, where
      // the line does keep going, is the only case that isn't trusted.
      if (!tolerateBlanks || blanksTolerated >= 1) break;
      const nr = axis === 'row' ? r + 1 : r;
      const nc = axis === 'col' ? c + 1 : c;
      const next = worksheet.getCell(nr, nc);
      const nextHasData = !isBlank(cellRawValue(next));
      const nextInRange = !reference || reference(nr, nc);
      if (!nextHasData && nextInRange) break;
      blanksTolerated++;
    }

    // Walked into a cell that's already claimed as a *different* field's
    // own anchor (e.g. extending "營收" rightward reaches the cell that
    // matched "虧損") -- that cell is another block's label, not more of
    // this one's data, no matter how tightly packed the sheet is.
    if (otherAnchors && otherAnchors.has(cell.address)) break;

    if (prevCell) {
      const trailing = axis === 'col' ? prevCell.border?.right : prevCell.border?.bottom;
      const leading = axis === 'col' ? cell.border?.left : cell.border?.top;
      if (isBoundaryBorder(trailing) || isBoundaryBorder(leading)) break;
    }

    n++;
    prevCell = cell;
  }
  return n;
}

// A row of data almost always has a header/label line naming each of its
// real columns -- either the row directly above it (period headers like
// "Q1, Q2, Q3...") or the column directly to its left (row labels running
// down the side). That line is a far more reliable boundary than the data
// itself, which can butt straight up against an unrelated block with no
// visual gap at all. Only trust it if it actually has something behind the
// first data cell -- otherwise there's no real header/label line here and
// this falls back to blank/border detection alone.
function buildReference(worksheet, startRow, startCol, axis) {
  if (axis === 'col') {
    const refRow = startRow - 1;
    if (refRow < 1 || isBlank(cellRawValue(worksheet.getCell(refRow, startCol)))) return null;
    return (_r, c) => !isBlank(cellRawValue(worksheet.getCell(refRow, c)));
  }
  const refCol = startCol - 1;
  if (refCol < 1 || isBlank(cellRawValue(worksheet.getCell(startRow, refCol)))) return null;
  return (r, _c) => !isBlank(cellRawValue(worksheet.getCell(r, refCol)));
}

// Guards the *perpendicular* axis of a block against swallowing the next
// field down/across. A 'right' block's own label sits in `anchorCol` --
// growing downward into a 2D matrix is fine as long as that label column
// stays blank on the way down (that's still the same table), but the
// moment a new row picks up its own label there, that's a different
// field's row, not a continuation of this one. `axis` says which
// perpendicular direction is being grown: 'row' guards height (for a
// 'right' block), 'col' guards width (for a 'down' block).
function buildLabelGuard(worksheet, anchorRow, anchorCol, axis) {
  if (axis === 'row') {
    return (r, _c) => r === anchorRow || isBlank(cellRawValue(worksheet.getCell(r, anchorCol)));
  }
  return (_r, c) => c === anchorCol || isBlank(cellRawValue(worksheet.getCell(anchorRow, c)));
}

// Starting adjacent to `anchor` on the given `origin` side, auto-detect a
// rectangular block of data: measure how far right the first row runs and
// how far down the first column runs (blank cell, divider border, or a
// lapsed header/label line ends each run), then pull that width x height
// rectangle. No fixed cell count or manual cap is needed -- the block's
// own boundary decides its size, and it can resolve to a single cell, a
// single row/column, or a full 2D block.
//
// The header/label reference line only applies to the axis that starts
// adjacent to the anchor (width for a 'right' block, height for a 'down'
// one) -- the perpendicular axis is how a 2D block (e.g. a matrix under a
// corner label) is allowed to grow, so it stays on blank/border alone.
//
// A merged anchor overrides that same axis with its own span (see
// findMergeSpan above) -- e.g. a title merged across A1:D1 above a 'down'
// block fixes the width at 4 regardless of what blank/border scanning of
// the data rows would have found, since the merge is a more deliberate
// signal than the data itself.
function extractBlockForOrigin(worksheet, anchor, origin, otherAnchors) {
  const mergeSpan = findMergeSpan(worksheet, anchor.row, anchor.col);

  const startRow = mergeSpan && origin === 'right' ? mergeSpan.top : origin === 'down' ? anchor.row + 1 : anchor.row;
  const startCol = mergeSpan && origin === 'down' ? mergeSpan.left : origin === 'right' ? anchor.col + 1 : anchor.col;

  let width;
  if (mergeSpan && origin === 'down') {
    width = mergeSpan.right - mergeSpan.left + 1;
  } else {
    const widthRef =
      origin === 'right'
        ? buildReference(worksheet, startRow, startCol, 'col')
        : buildLabelGuard(worksheet, anchor.row, anchor.col, 'col');
    width = measureRun(worksheet, startRow, startCol, 'col', widthRef, otherAnchors, origin === 'right' && Boolean(widthRef));
  }

  let height;
  if (mergeSpan && origin === 'right') {
    height = mergeSpan.bottom - mergeSpan.top + 1;
  } else {
    const heightRef =
      origin === 'down'
        ? buildReference(worksheet, startRow, startCol, 'row')
        : buildLabelGuard(worksheet, anchor.row, anchor.col, 'row');
    height = measureRun(worksheet, startRow, startCol, 'row', heightRef, otherAnchors, origin === 'down' && Boolean(heightRef));
  }

  if (width === 0 || height === 0) return { rows: 0, cols: 0, cells: [] };

  const cells = [];
  for (let r = 0; r < height; r++) {
    const rowCells = [];
    for (let c = 0; c < width; c++) {
      const cell = worksheet.getCell(startRow + r, startCol + c);
      rowCells.push({ address: cell.address, value: cellRawValue(cell) });
    }
    cells.push(rowCells);
  }
  return { rows: height, cols: width, cells };
}

// A single field can't be told in advance whether a given report lays its
// data out beside its label or below it -- real reports mix both, often
// file to file. So try both origins and keep whichever finds more actual
// cells; on a tie (including both empty) prefer 'down', since a label
// sitting above a column of values is the more common shape in practice.
//
// When *both* directions turned up real data, this was a genuine fork --
// even though one is bigger and almost certainly right, that's still a
// guess. `ambiguous` surfaces that so the UI can flag it for a human
// glance instead of silently trusting the guess.
function extractBlock(worksheet, anchor, otherAnchors) {
  const right = extractBlockForOrigin(worksheet, anchor, 'right', otherAnchors);
  const down = extractBlockForOrigin(worksheet, anchor, 'down', otherAnchors);
  const rightSize = right.rows * right.cols;
  const downSize = down.rows * down.cols;
  const chosen = downSize >= rightSize ? down : right;
  return { ...chosen, ambiguous: rightSize > 0 && downSize > 0 };
}

// Extraction anchored at a cell the user picked by hand, bypassing keyword
// search entirely -- the escape hatch for when auto-detection guessed
// wrong. Reuses the exact same boundary logic (blank/border/label
// line/merge/other-anchors) from that point on, so it's just as capable,
// only the starting cell is manual instead of found by fuzzy match.
function extractManual(worksheet, address, otherAnchors) {
  const m = String(address).trim().toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  const col = colLetterToNumber(m[1]);
  const row = Number(m[2]);
  const cell = worksheet.getCell(row, col);
  const anchor = { row, col, address: cell.address, text: cellText(cell), score: 1 };
  const block = extractBlock(worksheet, anchor, otherAnchors || new Set());
  return {
    matched: block.rows > 0 && block.cols > 0,
    sheet: worksheet.name,
    headerAddress: anchor.address,
    headerText: anchor.text,
    score: 1,
    manual: true,
    block,
  };
}

// Run one recipe (a list of fields) against one already-loaded workbook.
// Each field is searched independently across every worksheet, so a
// layout that shifts row/column position -- or even moves to a different
// tab -- from file to file still resolves correctly.
//
// Two passes: first locate every field's anchor, then extract each block.
// The second pass needs the full set of anchors up front so that, say,
// "營收"'s block extending rightward knows to stop the instant it reaches
// the cell that's already "虧損"'s own anchor, instead of swallowing it.
//
// Split from `processWorkbook` below so a workbook that was already
// loaded some other way (e.g. a legacy .xls file converted to .xlsx in
// memory first) can be scanned without a second disk read.
function processLoadedWorkbook(workbook, filePath, recipe, threshold = DEFAULT_THRESHOLD) {
  const anchors = {}; // field.name -> { match, sheet } | null
  for (const field of recipe.fields) {
    let bestMatch = null;
    let bestSheet = null;

    workbook.eachSheet((worksheet) => {
      const match = findHeaderCell(worksheet, field.keywords, threshold);
      if (match && (!bestMatch || match.score > bestMatch.score)) {
        bestMatch = match;
        bestSheet = worksheet;
      }
    });

    anchors[field.name] = bestMatch ? { match: bestMatch, sheet: bestSheet } : null;
  }

  // Anchor addresses grouped by sheet name, so a field only treats another
  // field's anchor as a boundary when it's actually on the same sheet.
  const anchorsBySheet = {};
  for (const [name, entry] of Object.entries(anchors)) {
    if (!entry) continue;
    (anchorsBySheet[entry.sheet.name] ||= []).push({ name, address: entry.match.address });
  }

  const fields = {};
  for (const field of recipe.fields) {
    const entry = anchors[field.name];
    if (!entry) {
      fields[field.name] = { matched: false };
      continue;
    }
    const { match: bestMatch, sheet: bestSheet } = entry;

    const otherAnchors = new Set(
      (anchorsBySheet[bestSheet.name] || [])
        .filter((a) => a.name !== field.name)
        .map((a) => a.address)
    );

    const block = extractBlock(bestSheet, bestMatch, otherAnchors);

    fields[field.name] = {
      matched: block.rows > 0 && block.cols > 0,
      sheet: bestSheet.name,
      headerAddress: bestMatch.address,
      headerText: bestMatch.text,
      score: Math.round(bestMatch.score * 100) / 100,
      ambiguous: block.ambiguous,
      block,
    };
  }

  return { file: filePath, fields };
}

async function processWorkbook(filePath, recipe, threshold = DEFAULT_THRESHOLD) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  return processLoadedWorkbook(workbook, filePath, recipe, threshold);
}

module.exports = {
  normalize,
  similarity,
  levenshtein,
  scoreCell,
  findHeaderCell,
  extractBlock,
  extractManual,
  processWorkbook,
  processLoadedWorkbook,
  DEFAULT_THRESHOLD,
};
