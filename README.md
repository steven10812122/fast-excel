# Fast Excel

**Stop copy-pasting the same cells out of a folder of messy Excel reports every month.** Fast Excel finds your data by column/row **keyword**, not by fixed cell position — so it doesn't matter that this month's revenue number moved from `B2` to `C5`, or that someone merged a title cell across four columns.

用關鍵字抓資料,不用管每份 Excel 檔案的格子長在哪。丟一個資料夾進去,設定「我要抓哪些欄位」,它會自動在每個檔案裡找到符合的表頭,抓出對應的區塊,合併成一份報表。

![Fast Excel screenshot (English)](build/screenshot-en.png)
![Fast Excel screenshot (中文)](build/screenshot.png)

The UI is bilingual (English / 繁體中文), auto-detected from your OS language with a manual toggle in the top-left corner — the matching engine itself is language-agnostic, so keywords work the same whether they're "Revenue" or "營收".

## Why

Every month, someone opens 20 Excel files, hunts for the same handful of cells in each one, copies, and pastes them into a summary sheet for the exec report. It's boring, slow, and the layout is never quite consistent from file to file — one month the header is in row 1, the next it's in row 2 because someone inserted a row.

Fast Excel automates that hunt. You describe *what* you're looking for (keywords, with synonyms), not *where* it is. It figures out the "where" per file.

## How the matching works

No manual cell ranges, no fixed "grab N cells" count. A field's range is auto-detected using five stacked signals, any one of which can end it:

1. **Blank cell** — the classic case, stop at the first empty cell.
2. **Deliberate border** — a medium/thick/double border marks a table edge even when the next cell isn't blank (e.g. an unrelated table sits right next to it).
3. **Header/label line** — the row above (or column to the left) naming each real column is trusted over the data row itself, and can even bridge a single isolated missing value (one month's number is blank, but neighbors aren't).
4. **Another field's own anchor** — if two labels sit right next to each other with zero gap ("營收" then immediately "虧損"), the first field's range stops exactly where the second one's label starts.
5. **A merged, centered title** — if the label is a merged cell (e.g. `A1:D1`), its own span is used as the width/height directly, since that's the report author explicitly stating "this covers 4 columns."

Direction (label-then-values-to-the-right vs. label-above-a-column-of-values) is auto-picked per anchor too — both are tried, and whichever finds more real data wins. When both directions found *something*, the result is flagged **ambiguous** in the UI so you can glance at it before trusting it.

None of this is a black box: click any result cell to see exactly which cell it matched, with what confidence, and the raw grid it pulled — and to manually override the anchor if it guessed wrong.

## Benchmark

`npm run benchmark` generates 24 files -- two years of monthly reports for a fictional company, systematically cycling through every messy layout this engine handles (position drift, merged titles, borders with unrelated data glued on, zero-gap adjacent fields, missing values, legacy .xls, mixed English/Chinese labels, a formula-error cell) -- then scans them with the real engine and checks every extracted value against the ground truth recorded at generation time.

```
Files:          24
Fields/file:    4
Time:           86 ms (3.6 ms/file)
Fields checked: 96
Correct:        96 (100.0%)
```

This is generated data with known-correct answers, not a claim about arbitrary real-world files -- see Known limitations below for where the engine genuinely can't tell blocks apart. The point of committing the generator (`benchmark/generate.js`) rather than just the numbers is that anyone can regenerate and re-verify this themselves.

## Features

- **Keyword + synonym matching**, fuzzy (handles typos/rewording) — not exact string match
- **Auto-detected 2D block extraction** — a single value, a row, a column, or a full matrix, whichever the sheet actually has
- **Preview & manual override** — see what matched before you trust it; re-anchor a wrong match by hand; exclude a bad match from export
- Reads both **.xlsx and legacy .xls**
- **Recursive folder scan**, with a progress bar and cancel
- **Recipe import/export** as a JSON file — share your field setup with a teammate or another machine
- Multi-value fields spread into real numbered columns in the export (not just joined text) when the shape allows it
- **Bilingual UI** (English / 中文), auto-detected with a manual toggle
- Runs entirely **locally** — nothing leaves your machine

## Known limitations

- If two tables sit right next to each other with **no blank, no border, and no header-line difference between them**, there's no signal left for a rule-based matcher (or a human glancing at bare numbers) to tell them apart automatically. Use the manual override in that case.
- Only one blank value is tolerated inside a block (to bridge one missing month's data); two blanks in a row are treated as the real end of the table.
- A field whose match is a genuine 2D matrix in some files exports as one joined-text cell rather than spread columns, since a matrix doesn't fit a "one row per file" table.

## Getting started

Requires [Node.js](https://nodejs.org) 18+.

```bash
npm install
npm start        # run the app
npm test          # run the test suite
```

### Build a distributable

```bash
npm run dist:win   # Windows installer (.exe, NSIS)
npm run dist:mac    # macOS app (.dmg)
```

Both are **unsigned** builds (no paid code-signing certificate). What that actually means per platform, verified rather than guessed:

- **Windows**: SmartScreen shows a warning on first run. Click "More info" → "Run anyway" to proceed.
- **macOS**: this is worse than a simple warning. Verified with `spctl -a -vv -t execute` against both built `.app` bundles: Gatekeeper **rejects them outright** (`rejected / no usable signature` on Intel, `code has no resources but signature indicates they must be present` on Apple Silicon — the arm64 linker attaches an ad-hoc signature that isn't a full, Gatekeeper-accepted one). After downloading through a browser, this will most likely show as **"Fast Excel" is damaged and can't be opened**, not the milder "unidentified developer" prompt that a right-click → Open bypasses. To actually run it:
  - Open Terminal and run `xattr -cr "/Applications/Fast Excel.app"` (strips the quarantine flag the browser added), **or**
  - Try opening it once (it will be blocked), then go to System Settings → Privacy & Security, scroll down, and click "Open Anyway" next to the blocked-app notice.

  The proper long-term fix is an Apple Developer ID certificate + notarization, which costs money and isn't part of this build.

## License

MIT — see [LICENSE](LICENSE).
