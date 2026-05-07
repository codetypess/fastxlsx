# FastXLSX Load Performance

## Background

`fastxlsx` already supports workbook-level inspection and full worksheet access, but the current read path becomes expensive when consumers only need workbook shell metadata or a viewport-sized slice of a large sheet.

Today, callers typically do one of two things:

- open a workbook and inspect workbook-level metadata such as sheet names, sheet visibility, active sheet, and defined names
- open a workbook and then expand one or more worksheets through `Sheet` APIs such as `getCell()`, `getDisplayValue()`, `getFormula()`, `getStyleId()`, `getCellEntries()`, `rowCount`, or `columnCount`

The expensive path starts when a caller crosses into the current sheet index model. `Sheet` getters build a full in-memory index for the worksheet by scanning `<sheetData>` and materializing row and cell structures for the whole sheet. That is correct and stable, but it is much more work than consumers like viewport-based editors or diff tools need for first paint.

This work item defines additive read primitives for:

- lightweight workbook manifest reads
- bounded worksheet window reads
- cache reuse across repeated window reads on the same sheet

The goal is to make those read paths explicit without changing the semantics of the existing full-sheet APIs.

## Goals

- Add an opt-in workbook manifest API that returns a stable, detached workbook shell object without requiring callers to enumerate worksheet cells.
- Add an opt-in worksheet window read API that returns sparse cell data and relevant layout metadata only for a requested row and column range.
- Preserve current cell value, display value, formula, style-id, row-style, column-style, width, height, merge, freeze-pane, and auto-filter semantics for data inside the requested window.
- Reuse workbook-level caches and per-sheet read state across repeated window reads so revisiting the same sheet does not rebuild the same lookup state from scratch.
- Keep the existing eager `Workbook` and `Sheet` behavior available for callers that still want the full in-memory model.

## Non-goals

- Rewriting workbook save or serialization behavior.
- Changing the semantics of existing eager getters such as `sheet.getCell()` or `sheet.getCellEntries()`.
- Transparently speeding up every existing full-sheet caller without opt-in.
- Adding a CLI surface in the first slice.
- Designing consumer-specific paging protocols for editors, webviews, or other applications.

## Public Surface

This feature is additive.

```ts
interface WorkbookManifest {
  sheetCount: number;
  visibleSheetCount: number;
  activeSheetName: string | null;
  sheets: WorkbookSheetManifest[];
  definedNames: DefinedName[];
}

interface WorkbookSheetManifest {
  name: string;
  visibility: SheetVisibility;
}

interface SheetWindowReadOptions {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

interface SheetWindowCell extends CellEntry {
  displayValue: string | null;
}

interface SheetWindowRowEntry {
  alignment: CellStyleAlignment | null;
  hidden: boolean;
  height: number | null;
  rowNumber: number;
  styleId: number | null;
}

interface SheetWindowColumnEntry {
  alignment: CellStyleAlignment | null;
  columnLabel: string;
  columnNumber: number;
  hidden: boolean;
  styleId: number | null;
  width: number | null;
}

interface SheetWindowSnapshot {
  sheetName: string;
  requestedRange: string;
  clampedRange: string | null;
  sheetRange: string | null;
  rowCount: number;
  columnCount: number;
  cells: SheetWindowCell[];
  cellAlignments: Record<string, CellStyleAlignment>;
  rowAlignments: Record<string, CellStyleAlignment>;
  columnAlignments: Record<string, CellStyleAlignment>;
  rowStyleIds: Record<string, number>;
  rowHeights: Record<string, number>;
  hiddenRows: number[];
  columnStyleIds: Record<string, number>;
  columnWidths: Record<string, number>;
  hiddenColumns: string[];
  mergedRanges: string[];
  freezePane: FreezePane | null;
  autoFilter: AutoFilterDefinition | null;
}

interface Workbook {
  getManifest(): WorkbookManifest;
  readSheetWindow(sheetName: string, options: SheetWindowReadOptions): SheetWindowSnapshot;
}

interface Sheet {
  readWindow(options: SheetWindowReadOptions): SheetWindowSnapshot;
  *iterWindowCells(options: SheetWindowReadOptions): IterableIterator<SheetWindowCell>;
  *iterWindowRows(options: SheetWindowReadOptions): IterableIterator<SheetWindowRowEntry>;
  *iterWindowColumns(options: SheetWindowReadOptions): IterableIterator<SheetWindowColumnEntry>;
}
```

Public API rules:

- `WorkbookManifest` is a detached DTO. Mutating it does not mutate the workbook.
- `WorkbookSheetManifest` intentionally stays small. It is a workbook shell object, not a worksheet summary with counts.
- `SheetWindowSnapshot.cells` is sparse. Missing cells inside the requested range are omitted instead of emitted as null placeholders.
- `SheetWindowSnapshot` also includes sparse alignment maps for cells, rows, and columns.
- `SheetWindowCell.styleId` follows the existing raw cell style-id semantics from `CellEntry`. Effective row and column defaults remain separate through `rowStyleIds` and `columnStyleIds`.
- `rowStyleIds`, `rowHeights`, `columnStyleIds`, and `columnWidths` include only explicit metadata inside the clamped window.
- `hiddenRows` includes row numbers inside the clamped window that are explicitly hidden.
- `hiddenColumns` includes column labels inside the clamped window that are explicitly hidden.
- `mergedRanges` includes only merged ranges that intersect the clamped window.
- `freezePane` and `autoFilter` remain whole-sheet metadata because they are small and are commonly needed for layout or filtering decisions.
- `readSheetWindow()` throws the same sheet-not-found error shape as other `Workbook` sheet lookups.

## Internal Model

The implementation should add a dedicated read-cache layer instead of trying to force window behavior through the existing full `SheetIndex`.

Recommended internal model:

- `WorkbookManifest`
  - built from the existing workbook context, active-sheet metadata, sheet visibility, and defined names
- `WorkbookReadCache`
  - reuses the existing workbook-level shared strings and styles caches
  - owns per-sheet lazy read state keyed by worksheet path
- `SheetReadCache`
  - stores sheet XML text and lightweight parse products needed by repeated window reads
  - stores row boundary metadata so window reads can skip directly to relevant rows
  - stores parsed column definitions, merged ranges, freeze-pane state, and auto-filter definition
  - stores shared-formula master state so dependent formula cells can be resolved even when the master cell sits outside the requested window

The important design constraint is that `SheetReadCache` is not a second eager `SheetIndex`.

It should be allowed to scan the sheet once to build lightweight row offsets and shared-formula master metadata, but it should not materialize full cell snapshots for rows and cells that the caller did not request.

`SheetWindowSnapshot` should be constructed from:

- the clamped row and column bounds
- sparse logical cells that intersect those bounds
- explicit row and column metadata inside those bounds
- whole-sheet metadata that is small and already useful for consumers

## OOXML Mapping

This work item is read-only. It does not introduce new XML writers.

Workbook-level mapping:

- `_rels/.rels`
  - resolve the workbook part path
- `xl/workbook.xml` or equivalent workbook part
  - parse sheet order
  - parse active sheet index
  - parse defined names
  - parse per-sheet visibility
- `xl/_rels/workbook.xml.rels` or equivalent
  - resolve worksheet part paths
  - resolve `styles.xml`
  - resolve `sharedStrings.xml`

Worksheet-level mapping:

- `<dimension ref="...">`
  - may be used as a cheap bounds hint
  - must not be treated as authoritative if later row scanning proves it stale
- `<cols><col .../></cols>`
  - parse explicit width, hidden, and style metadata
- `<sheetViews>`
  - parse freeze-pane state
- `<autoFilter>`
  - parse structured filter definition and nested `sortState`
- `<mergeCells><mergeCell ref="..."/></mergeCells>`
  - collect only ranges intersecting the requested window in the public snapshot
- `<sheetData><row ...><c ...>...</c></row></sheetData>`
  - parse row numbers, row style ids, row heights, row hidden state
  - parse cell addresses, raw types, style ids, formulas, inline strings, shared strings, errors, and values
- `<f t="shared" si="...">`
  - preserve shared-formula semantics even when the master formula cell is outside the requested window

Because this feature is read-only, all untouched XML remains untouched by definition. No new serializer or XML patch helper should be introduced as part of this work item unless implementation later proves that cache invalidation needs a shared helper.

## Mutation Semantics

This work item introduces no new workbook mutation APIs. Its mutation semantics are cache and snapshot semantics.

- `getManifest()` and `readSheetWindow()` return detached snapshots of workbook state at read time.
- Mutating a returned manifest or window snapshot does not affect the workbook.
- Returned snapshots do not live-update after later workbook edits. Callers must re-read after mutation if they want fresh state.
- `readSheetWindow()` clamps the requested bounds to the current logical sheet bounds.
- If the requested window falls fully outside the current logical sheet bounds, `clampedRange` is `null`, `cells` is empty, and window-scoped row and column metadata collections are empty.
- `sheetRange`, `rowCount`, and `columnCount` describe the full current logical worksheet range, not just the requested window.

Cache invalidation rules:

- Mutating a worksheet invalidates that sheet's `SheetReadCache`.
- Mutating workbook sheet order, active sheet, or defined names invalidates manifest-derived cache state.
- Mutating `styles.xml` invalidates style-dependent read cache state.
- Mutating `sharedStrings.xml` invalidates shared-string-dependent read cache state.
- A cache must never return cell, formula, display, or metadata data from stale XML after a workbook mutation.

## Structure Transform Semantics

This feature is read-only, but structure transforms still matter because callers may read windows before and after workbook mutations.

After any of the following existing operations:

- `insertRow`
- `deleteRow`
- `insertColumn`
- `deleteColumn`
- sheet rename
- sheet reorder
- style mutation that changes style definitions

the next `getManifest()` or `readSheetWindow()` call must reflect the new workbook state and must not reuse stale row offsets, stale merged-range intersections, stale freeze-pane positions, stale auto-filter ranges, or stale shared-formula masters.

Specific expectations:

- row and column insert/delete operations shift returned cell addresses, row metadata, column metadata, and intersecting merged ranges exactly as the current eager APIs would observe them after the same mutation
- sheet rename updates `WorkbookManifest.sheets[].name` and the target used by `readSheetWindow(sheetName, ...)`
- sheet reorder preserves the same per-sheet content while changing manifest order
- old detached snapshots remain stale objects; they are not patched in place

## Compatibility

This change is additive.

- Existing workbook open flows remain supported:
  - `Workbook.open()`
  - `Workbook.fromUint8Array()`
  - `Workbook.fromArrayBuffer()`
- Existing eager sheet APIs remain supported and unchanged in semantics.
- Existing callers do not need to migrate unless they want the new manifest or windowed read behavior.
- The first slice does not add CLI commands or CLI JSON output.

Compatibility strategy:

- keep `Workbook` and `Sheet` as the main public surfaces
- add explicit opt-in methods rather than changing current getter return types
- reuse existing `DefinedName`, `FreezePane`, `AutoFilterDefinition`, `CellEntry`, and style APIs instead of inventing parallel type systems

## Test Matrix

### `test/lossless.test.ts`

Add or extend coverage for:

- `workbook.getManifest()` returns sheet order, active sheet, visibility, and defined names consistent with existing workbook getters
- `sheet.readWindow()` returns sparse cells with the same `value`, `displayValue`, `formula`, and raw cell `styleId` as existing eager getters for the same addresses
- row and column metadata inside the requested window matches existing sheet getters:
  - row style ids
  - row heights
  - hidden rows
  - column style ids
  - column widths
  - hidden columns
- `mergedRanges` includes only ranges intersecting the window
- freeze-pane and auto-filter metadata remain whole-sheet and match existing getters
- shared-formula cells resolve correctly when the shared-formula master is outside the requested window
- reading a window after `insertRow`, `deleteRow`, `insertColumn`, or `deleteColumn` reflects the updated structure and does not reuse stale cached state
- reading a window after shared-string or style mutations reflects the updated workbook state

### `test/real-files.test.ts`

Add checks for one or more real workbooks where:

- `getManifest()` matches current workbook-level behavior
- `readWindow()` matches eager cell and metadata reads for representative windows
- a sparse large sheet can be read through a small window without requiring the consumer test to enumerate the full sheet

### `test/interop-matrix.test.ts`

If the new APIs become part of the stable observable surface, add snapshot fields for:

- workbook manifest shell metadata
- representative window-read cell and metadata outputs

### `test/xml-fuzz.test.ts`

Add targeted coverage if the new reader introduces new low-level row or cell scanning helpers, especially for:

- self-closing rows or cells
- mixed attribute ordering
- shared formulas
- inline strings and shared strings

### Performance verification

Add a focused benchmark or fixture-driven verification path outside normal correctness assertions that compares:

- full-sheet eager enumeration
- `getManifest()`
- `readWindow()` for a small top-of-sheet range
- repeated `readWindow()` calls on the same sheet

The benchmark does not need to become a hard CI gate in the first slice, but it must be easy to rerun locally on a representative large workbook fixture.

## Implementation Status

Implemented in `fastxlsx`:

- `Workbook.getManifest()`, `Workbook.readSheetWindow()`, and `Sheet.readWindow()` are available on the public API surface.
- `Sheet.iterWindowCells()`, `Sheet.iterWindowRows()`, and `Sheet.iterWindowColumns()` expose sparse window iteration helpers.
- Window snapshots include cell, row, and column alignment maps in addition to raw style ids and layout metadata.
- `scripts/benchmark.ts` reports `manifestResult` and `windowResult` alongside the existing dense, sparse, and batch-write benchmark modes.
- Local verification passes with `npm test`, `npm run build`, and `node --import tsx scripts/benchmark.ts res/monster.xlsx 1`.

## Acceptance

- `fastxlsx` exposes additive `Workbook.getManifest()`, `Workbook.readSheetWindow()`, and `Sheet.readWindow()` APIs with stable TypeScript types.
- Existing eager workbook and sheet APIs remain source-compatible and pass current tests unchanged.
- Window reads preserve current eager semantics for included cells and included row and column metadata.
- Shared-formula cells inside a requested window resolve correctly even when their master formula sits outside that window.
- Cache invalidation keeps post-mutation reads correct after sheet edits, structure transforms, style changes, and shared-string changes.
- Focused correctness tests for the new APIs pass in the relevant suites.
- A local benchmark or repeatable fixture check demonstrates that manifest reads and viewport-sized window reads perform less work than full-sheet eager expansion for representative large sheets.
