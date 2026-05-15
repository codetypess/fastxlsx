# FastXLSX Deep Window Read Performance

## Background

`fastxlsx` already exposes sparse worksheet window reads through:

- `Workbook.readSheetWindow()` / `Sheet.readWindow()`
- `Workbook.readSheetValueWindow()` / `Sheet.readValueWindow()`

Those APIs are a good fit for viewport-based editors, but the current deep-row path still degrades with the requested start row.

Today:

- full window reads walk `base.rowInfos` from the beginning twice:
  - once in `collectRowWindowMetadata()`
  - once in `readWindowCells()`
- value-window reads also walk `rowInfos` from the beginning until they reach `startRow`
- column metadata reads scan all parsed column definitions for every requested column

That means a request for rows near `20000` still pays work proportional to the number of earlier parsed row entries, even though the cache already stores sorted row metadata. In a scroll-driven editor, repeated deep-window reads turn that into visible white gaps while data catches up.

## Goals

- Remove the `O(startRow)` row-scan behavior from repeated window reads.
- Reuse one row-range lookup for the whole request instead of rescanning from the beginning in multiple phases.
- Keep full-window and value-window behavior unchanged for returned data.
- Improve sequential deep-window reads enough that adjacent requests do not grow linearly with row depth.
- Reduce the per-request column metadata overhead without changing column-style semantics.

## Non-goals

- Adding a new public sheet reader or cursor API in this slice.
- Changing `SheetWindowSnapshot` or `SheetValueWindowSnapshot`.
- Reworking eager `SheetIndex` accessors such as `getCell()` or `iterCellEntries()`.
- Adding cancellation, accumulation, or request coalescing behavior in downstream editor consumers.

## Public Surface

This change is internal-only.

- `Workbook.readSheetWindow()`, `Sheet.readWindow()`, `Workbook.readSheetValueWindow()`, and `Sheet.readValueWindow()` keep their current signatures and return shapes.
- Cache invalidation behavior stays unchanged.

## Internal Model

Extend the lightweight base sheet read cache with row-window lookup metadata derived from the already sorted `rowInfos`.

Recommended cache shape:

```ts
interface BaseSheetReadCache {
  rowInfos: SheetRowReadInfo[];
  rowNumbers: number[];
  sheetXml: string;
  usedBounds: UsedBounds | null;
}
```

Rules:

- `rowInfos` remain sorted by `rowNumber`
- `rowNumbers[index] === rowInfos[index].rowNumber`
- window reads resolve the first row index with `rowNumber >= startRow`
- window reads resolve the first row index with `rowNumber > endRow`
- the effective row slice is then `[startIndex, endExclusiveIndex)`

Complexity target:

- row-range lookup: `O(logN)`
- row iteration inside the window: `O(windowRows + windowCells)`
- no extra `O(startRow)` prefix walk after the cache is built

## Read Semantics

### Full window reads

`readWorkbookSheetWindow()` should:

1. resolve the requested row slice once
2. iterate only rows inside that slice
3. collect row metadata and sparse cells in the same row pass
4. resolve column metadata separately for the requested column range

Preserved semantics:

- sparse logical-cell filtering stays unchanged
- cell, row, and column alignment resolution stays unchanged
- shared-formula followers still resolve correctly when the anchor is outside the requested viewport
- comment-aware used bounds for full reads stay unchanged

### Value window reads

`readWorkbookSheetValueWindow()` should use the same indexed row-slice lookup so deep value-window reads also avoid prefix scans.

Preserved semantics:

- cells remain sparse and row-major
- logical formula cells remain included even when cached value is `null`
- blank placeholder `<c>` nodes remain excluded

### Column metadata

Column metadata collection should stop doing `for each requested column -> scan all definitions`.

Accepted strategies for this slice:

- build a per-request window state by scanning intersecting definitions once, or
- use a definition lookup structure that avoids full rescans for each column

Required semantic rule:

- when multiple `<col>` definitions overlap a requested column, later matching definitions still win exactly as today

## Mutation Semantics

This feature adds no new mutation APIs.

Existing invalidation behavior remains required:

- sheet mutations invalidate both base and full sheet window caches
- workbook-level invalidation drops window read state

Returned snapshots remain detached DTOs.

## Structure Transform Semantics

This slice is read-only, but reads after structural edits must still reflect updated worksheet XML.

Required preserved behavior after:

- `insertRow`
- `deleteRow`
- `insertColumn`
- `deleteColumn`
- cell writes and deletes

The next window read must use refreshed row lookup metadata rebuilt from the latest XML.

## Test Matrix

Add or update coverage for:

- full window reads returning the same sparse cells and row metadata after the row-pass refactor
- shared-formula resolution still working when the anchor is outside the requested window
- deep window reads returning correct cells and metadata for high row numbers
- repeated sequential window reads returning stable results across overlapping deep windows
- value-window reads still returning correct deep-row results

Add benchmark coverage for:

- one large sheet reading windows near rows `1`, `5000`, `10000`, and `20000`
- overlapping sequential window reads such as `8000-8060`, `8040-8100`, `8080-8140`

## Acceptance Criteria

- full window reads no longer scan `rowInfos` from the beginning multiple times per request
- full and value window reads both use indexed row-slice lookup instead of prefix walks
- deep-row window latency no longer grows linearly with `startRow` after cache build
- column metadata collection avoids scanning all column definitions for every requested column
- existing semantic tests continue to pass
- benchmark output can demonstrate stable deep-window behavior across far-apart row ranges
