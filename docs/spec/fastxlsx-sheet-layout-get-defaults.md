# Sheet Layout Get Defaults

## Background

`fastxlsx sheet layout get` currently parses omitted `--columns` and `--rows` filters as empty arrays. That makes `columnWidths` and `rowHeights` come back as empty objects even when the worksheet already has explicit layout metadata.

This is a poor fit for inspection workflows. Users expect `get` to return the current layout state by default, not an empty projection that requires extra flags to be useful.

There is also a read-path constraint to account for: worksheet window reads clamp to used cell bounds, so they are not a safe source for "all layout metadata" when a workbook has explicit widths or heights on otherwise empty columns or rows.

## Goals

- Make `sheet layout get` return all explicit column widths when `--columns` is omitted.
- Make `sheet layout get` return all explicit row heights when `--rows` is omitted.
- Preserve the existing filtered behavior when callers do pass `--columns` or `--rows`.
- Ensure default reads include explicit layout metadata on rows or columns that have no logical cell content.

## Non-goals

- Expanding `sheet layout get` to include hidden flags, styles, or alignment metadata.
- Changing `sheet layout set`.
- Returning unbounded row or column keys with `null` values.

## Public Surface

### Sheet APIs

Add additive read helpers on `Sheet`:

```ts
interface Sheet {
  getColumnWidths(): Record<string, number>;
  getRowHeights(): Record<string, number>;
}
```

Semantics:

- `getColumnWidths()` returns only explicit width entries keyed by column label.
- `getRowHeights()` returns only explicit height entries keyed by row number string.
- Rows or columns without explicit height or width are omitted.
- For overlapping column definitions, later definitions win, matching `getColumnWidth(column)`.

### CLI

`fastxlsx sheet layout get <file> --sheet <name>` changes as follows:

- If `--columns` is omitted, `columnWidths` is populated from `sheet.getColumnWidths()`.
- If `--rows` is omitted, `rowHeights` is populated from `sheet.getRowHeights()`.
- If `--columns` is provided, `columnWidths` remains a keyed projection over the requested columns and may include `null` values.
- If `--rows` is provided, `rowHeights` remains a keyed projection over the requested rows and may include `null` values.

This is additive for callers that already pass explicit filters.

## Internal Model

- Reuse the existing worksheet row index plus parsed `<col>` definitions.
- Build explicit row height maps by scanning indexed row attribute sources and collecting `ht`.
- Build explicit column width maps by expanding `<col min="..." max="..." width="...">` spans into per-column labels.
- If a later overlapping `<col>` definition removes width for a covered column, that column is removed from the aggregate map to match single-column lookup behavior.

## OOXML Mapping

- Column widths come from worksheet `<cols><col ... width="..."/></cols>` definitions.
- Row heights come from worksheet `<row ... ht="..." customHeight="1">` attributes.
- The feature is read-only. No XML write behavior changes.

## Mutation Semantics

No new mutation behavior is introduced. The new `Sheet` helpers and the CLI default mode are read-only projections over current worksheet XML state.

## Structure Transform Semantics

No new structure transform logic is required. Existing insert/delete row and column operations continue to rewrite worksheet XML. The new read helpers simply reflect the current post-transform state.

## Test Matrix

- `test/lossless.test.ts`
  - `getRowHeights()` returns explicit heights only.
  - `getColumnWidths()` returns explicit widths only.
  - Hidden-only or style-only row and column definitions are not emitted in width or height maps.
- `test/cli.test.ts`
  - `sheet layout get` without `--columns` and `--rows` returns all explicit widths and heights.
  - Default reads include layout metadata on otherwise empty rows or columns.
  - Filtered `--columns` and `--rows` reads remain unchanged.

## Acceptance

- `sheet layout get` no longer returns empty layout maps by default when explicit widths or heights exist.
- Default output includes explicit widths and heights outside the used cell range.
- Filtered output remains backward-compatible for callers that already pass `--columns` or `--rows`.
