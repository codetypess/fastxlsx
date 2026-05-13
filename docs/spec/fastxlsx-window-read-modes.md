# FastXLSX Window Read Modes

## Background

`fastxlsx` currently exposes two sparse worksheet read entrypoints:

- `readWindow()`
  - returns value cells plus layout metadata such as comments, merge ranges, freeze panes, row styles, and column styles
- `readValueWindow()`
  - returns only sparse value cells and logical value/formula bounds

Those APIs are intentionally separate because the first-read cost is very different. `readWindow()` builds a heavier per-sheet cache than `readValueWindow()`.

That leaves two issues:

- callers cannot use one unified entrypoint to request either read mode
- the current full and value cache builders duplicate the same worksheet row scan shape even though both need `sheetXml`, row boundaries, and logical value/formula used bounds
- after a value-mode read, upgrading the same sheet to full-mode still rescans all cells through full snapshot parsing just to recover shared-formula anchors

## Goals

- Add an additive mode-based read surface to `Sheet.readWindow()` and `Workbook.readSheetWindow()`
- Keep `Sheet.readValueWindow()` and `Workbook.readSheetValueWindow()` as compatibility aliases
- Extract a shared base per-sheet cache for row boundaries and logical value/formula bounds
- Preserve the current performance contract where value-mode reads do not hydrate full layout metadata
- Preserve that lightweight contract when a caller later upgrades from value-mode to full-mode on the same sheet

## Non-goals

- Replacing `readValueWindow()` with a breaking API removal
- Collapsing full and value reads into one always-heavy cache
- Changing `iterWindowCells()` or `iterValueWindowCells()` in this slice
- Changing CLI behavior in this slice
- Reworking the primary full-first `readWindow()` row scan in this slice

## Public Surface

This change is additive.

```ts
type SheetWindowReadMode = "full" | "value";

interface SheetWindowReadSettings {
  mode?: SheetWindowReadMode;
}

interface Workbook {
  readSheetWindow(sheetName: string, options: SheetWindowReadOptions): SheetWindowSnapshot;
  readSheetWindow(
    sheetName: string,
    options: SheetWindowReadOptions,
    settings: { mode: "value" },
  ): SheetValueWindowSnapshot;
}

interface Sheet {
  readWindow(options: SheetWindowReadOptions): SheetWindowSnapshot;
  readWindow(options: SheetWindowReadOptions, settings: { mode: "value" }): SheetValueWindowSnapshot;
}
```

Rules:

- omitted `mode` means `"full"`
- `{ mode: "full" }` is equivalent to the current `readWindow()` behavior
- `{ mode: "value" }` is equivalent to the current `readValueWindow()` behavior
- `readValueWindow()` and `readSheetValueWindow()` remain supported and delegate to `{ mode: "value" }`
- return types remain detached snapshots

## Internal Model

Add a shared base sheet cache that stores only the data both read modes need:

- `sheetXml`
- row boundary metadata
- row attribute source
- logical value/formula used bounds

Suggested shape:

```ts
interface BaseSheetReadCache {
  sheetXml: string;
  rowInfos: SheetRowReadInfo[];
  usedBounds: UsedBounds | null;
}

interface FullSheetReadCache {
  base: BaseSheetReadCache;
  sharedFormulaAnchors: Map<string, SharedFormulaAnchor>;
  comments: SheetComment[];
  columnDefinitions: ColumnWindowDefinition[];
  freezePane: FreezePane | null;
  mergedRanges: WindowRange[];
  autoFilter: AutoFilterDefinition | null;
  usedBounds: UsedBounds | null;
}
```

Read-state behavior:

- value-mode reads reuse `BaseSheetReadCache`
- full-mode reads reuse `FullSheetReadCache`
- full-mode cache references `BaseSheetReadCache` instead of rebuilding its own independent row-boundary store
- when full-mode is the first caller, the implementation may build base-row metadata in the same pass as shared-formula anchor collection to avoid an unnecessary double scan
- when value-mode is the first caller, upgrading to full-mode must recover shared-formula anchors with a formula-focused scan that does not rebuild full cell snapshots for non-window cells

## OOXML Mapping

Base cache reads only worksheet XML under `<sheetData>`:

- `<row r="...">`
- `<c r="..." t="..." ...>`
- `<f ...>` for logical formula presence
- `<v>` for cached values
- `<is>` for inline strings

Full cache continues to additionally read:

- worksheet comments
- merged ranges
- freeze panes
- auto filters
- column definitions
- shared-formula anchor definitions

Mixed-mode upgrade constraint:

- shared-formula anchor recovery may rescan worksheet cells
- that rescan must only extract the shared-formula fields needed for follower translation:
  - cell address
  - shared index (`si`)
  - anchor formula text
- it must not decode cached values, shared strings, inline strings, or style ids just to build the upgrade cache

Value-mode cold-read constraint:

- base-cache construction is allowed to scan all worksheet cells once
- that scan should avoid duplicate per-cell XML parsing where possible
- specifically, value-mode row analysis should not require one helper to parse `<c ...>` metadata and a second helper to rescan the same cell body just to decide whether the cell is logically used

Untouched XML semantics remain unchanged because this feature is read-only.

## Mutation Semantics

This feature adds no setters.

Cache invalidation continues to follow existing workbook read-cache invalidation rules:

- any worksheet mutation invalidates both base and full caches for that sheet
- workbook-level invalidation drops all cached window read state

Returned snapshots remain detached DTOs.

## Structure Transform Semantics

This slice does not change row or column transform behavior.

Required preserved behavior:

- full-mode reads keep current comment-aware used-bounds semantics
- value-mode reads keep current value/formula-only used-bounds semantics
- after `insertRow`, `deleteRow`, `insertColumn`, `deleteColumn`, `setCell`, `setFormula`, `deleteCell`, and batch flush, both read modes observe updated worksheet XML and refreshed bounds

## Compatibility

This change is additive.

- existing `readWindow()` callers keep full-mode behavior
- existing `readValueWindow()` callers keep current behavior
- new callers may use `readWindow(..., { mode: "value" })` for the same lightweight semantics

## Test Matrix

Add coverage in `test/lossless.test.ts` for:

- `Sheet.readWindow(..., { mode: "value" })` matches `Sheet.readValueWindow(...)`
- `Workbook.readSheetWindow(..., { mode: "value" })` matches `Workbook.readSheetValueWindow(...)`
- `Sheet.readWindow(..., { mode: "full" })` keeps current layout metadata behavior
- value-mode unified reads still ignore comment-only bounds while full-mode unified reads still include them
- repeated mode-mixed reads continue to refresh after writes and structural edits
- a `readValueWindow()` call followed by `readWindow()` still resolves shared-formula followers correctly

## Acceptance Criteria

- spec-compliant mode-based overloads exist on `Sheet.readWindow()` and `Workbook.readSheetWindow()`
- existing alias methods remain supported
- value-mode reads continue to avoid full layout metadata hydration
- base per-sheet cache is shared between the two read families
- mixed `value -> full` upgrades do not require rebuilding full cell snapshots for off-window cells solely to recover shared-formula anchors
- targeted tests pass without changing current full-mode or value-mode semantics
