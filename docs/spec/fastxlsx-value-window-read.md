# FastXLSX Value Window Reads

## Background

`fastxlsx` already exposes two read modes for worksheet cells:

- eager `Sheet` getters such as `getCell()`, `getRow()`, `getCellEntries()`, and `rowCount`
- sparse window reads through `readWindow()`

That still leaves a gap for callers that only need cell values.

Today:

- eager getters build the full `SheetIndex`
- `readWindow()` avoids `SheetIndex`, but its first sheet-cache build still parses every cell snapshot and also collects layout metadata such as comments, merges, freeze panes, row styles, and column styles

For value-only consumers, that is more work than necessary.

## Goals

- Add an additive API for sparse worksheet value reads inside one requested window
- Avoid `buildSheetIndex()` for that API
- Avoid first-pass full-sheet parsing of style, display, comment, merge, and filter metadata for that API
- Preserve current sparse logical-cell semantics:
  - cells with a value are included
  - cells with a formula are included even when the cached value is `null`
  - blank placeholder `<c>` nodes without a value or formula are omitted
- Reuse a dedicated per-sheet cache across repeated value-window reads

## Non-goals

- A new full-sheet value iterator API
- CLI additions in this slice
- Returning styles, display strings, formulas, comments, merge ranges, row metadata, or column metadata
- Matching `readWindow()` used bounds for comment-only sheets
- Changing the existing `readWindow()` benchmark shape instead of extending it additively

## Public Surface

```ts
interface SheetValueWindowCell {
  address: string;
  columnNumber: number;
  rowNumber: number;
  value: CellValue;
}

interface SheetValueWindowSnapshot {
  sheetName: string;
  requestedRange: string;
  clampedRange: string | null;
  sheetRange: string | null;
  rowCount: number;
  columnCount: number;
  cells: SheetValueWindowCell[];
}

interface Workbook {
  readSheetValueWindow(sheetName: string, options: SheetWindowReadOptions): SheetValueWindowSnapshot;
}

interface Sheet {
  readValueWindow(options: SheetWindowReadOptions): SheetValueWindowSnapshot;
  iterValueWindowCells(options: SheetWindowReadOptions): IterableIterator<SheetValueWindowCell>;
}
```

Rules:

- `cells` is sparse and row-major
- `cells` includes logical formula cells even when `value` is `null`
- `sheetRange`, `rowCount`, and `columnCount` describe the logical value/formula bounds for the whole sheet
- comment-only cells do not extend `sheetRange`, `rowCount`, or `columnCount`
- if the requested window falls outside the current logical value/formula bounds, `clampedRange` is `null` and `cells` is empty

## Internal Model

Add a dedicated lightweight read cache alongside the existing window-read cache.

Per sheet, the value-window cache stores:

- `sheetXml`
- row boundary metadata needed to skip directly to relevant rows
- lightweight logical used-bounds metadata based only on value-or-formula presence

The cache does not store:

- parsed `CellSnapshot`s for every cell
- style alignments
- comments
- freeze panes
- auto filters
- merged ranges
- shared-formula text

Value parsing happens only for cells inside the effective clamped window.

## OOXML Mapping

This feature is read-only.

The implementation inspects only worksheet XML under `<sheetData>`:

- `<row r="...">`
- `<c r="..." t="..." ...>`
- `<f ...>` for formula presence
- `<v>` for cached values
- `<is>` for inline strings

Rules:

- used-bounds detection treats a cell as logical when it has either a formula tag or a value-bearing tag
- inline-string cells count as logical when they contain `<is>`
- value reads decode only cells inside the requested window

## Mutation Semantics

This feature does not add setters.

Returned snapshots are detached DTOs. Mutating them does not mutate the workbook.

## Structure Transform Semantics

This feature is read-only, but cached reads must track workbook mutations correctly.

Required behavior:

- reading a value window after `setCell()`, `setFormula()`, `deleteCell()`, or batch flush reflects the latest persisted worksheet XML
- reading a value window after `insertRow()`, `deleteRow()`, `insertColumn()`, or `deleteColumn()` reflects updated coordinates and value bounds
- workbook-level cache invalidation follows the same invalidation points as existing manifest and window reads

## Compatibility

This change is additive.

- existing eager getters and `readWindow()` semantics stay unchanged
- callers only opt into the lighter behavior by using the new value-window APIs

## Test Matrix

Add coverage in `test/lossless.test.ts` for:

- `Workbook.readSheetValueWindow()` returns sparse value cells inside a requested range
- `Sheet.readValueWindow()` and `iterValueWindowCells()` match the workbook-level result
- formula cells with cached values are included with those values
- formula cells without cached values are included with `value: null`
- blank placeholder cells are omitted
- comment-only sheets return empty value bounds
- value-window reads refresh after structural edits
- value-window reads refresh after cell writes

## Acceptance Criteria

- `Workbook.readSheetValueWindow()`, `Sheet.readValueWindow()`, and `Sheet.iterValueWindowCells()` are available on the public API surface
- value-window reads avoid `buildSheetIndex()`
- value-window results exclude style, display, merge, comment, and layout metadata
- repeated value-window reads on the same sheet reuse a lightweight cache until workbook mutation invalidates it
- tests cover formula, blank, comment-only, and structural-edit cases
- `scripts/benchmark.ts` reports additive `valueWindowResult` output alongside the existing `windowResult`
- viewport benchmark helpers request a fixed `50 x 20` window directly instead of consulting eager `sheet.rowCount` / `sheet.columnCount` first
