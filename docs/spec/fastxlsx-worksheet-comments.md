# FastXLSX Worksheet Comments

## Background

`fastxlsx` already supports whole-sheet worksheet comment APIs:

- `sheet.getComments()`
- `sheet.getComment(address)`
- `sheet.setComment(address, text, options?)`
- `sheet.removeComment(address)`
- `sheet.clearComments()`

That is enough for simple read and write flows, but it was not enough for viewport-based consumers.

The missing pieces were:

- `readWindow()` did not expose worksheet comments
- comments on blank cells could not be inferred from sparse `window.cells`
- row and column structural edits did not rewrite comment addresses
- `sortRange()` explicitly rejected ranges containing worksheet comments

## Goals

- Expose worksheet comments through bounded window reads
- Keep comments separate from sparse cell payloads so blank-cell comments survive
- Define row and column structural edit behavior for worksheet comments
- Preserve the existing `SheetComment` shape

## Non-goals

- Threaded comments
- Rich-text comment bodies
- Comment geometry, styling, or author-format details beyond current text and author fields
- Implementing `sortRange()` comment moves in this slice

## Public Surface

```ts
interface SheetComment {
  address: string;
  author: string | null;
  text: string;
}

interface SheetWindowSnapshot {
  // existing fields...
  comments: SheetComment[];
}

interface Sheet {
  getCommentsInRange(range: string): SheetComment[];
}
```

Rules:

- `readWindow()` returns comments whose addresses fall inside the effective clamped window range
- `window.comments` is sorted by normalized cell address
- comments remain separate from `window.cells`
- comments on blank cells are returned even when no matching sparse cell exists in `window.cells`

## Structural Edit Semantics

Required behavior:

- `insertRow(rowNumber, count)`
  - comments at or below `rowNumber` shift down by `count`
- `insertColumn(columnNumber, count)`
  - comments at or to the right of the insertion boundary shift right by `count`
- `deleteRow(rowNumber, count)`
  - comments inside the deleted row band are removed
  - comments below the deleted band shift up by `count`
- `deleteColumn(columnNumber, count)`
  - comments inside the deleted column band are removed
  - comments to the right of the deleted band shift left by `count`

For all four operations:

- unaffected comments preserve `author` and `text`
- resulting addresses remain normalized
- comments XML, VML drawing, sheet relationships, and content types remain consistent after save

## `sortRange()` Stance

This slice keeps the current explicit rejection:

- if comments exist inside the sorted range, `sortRange()` throws

That behavior remains acceptable until row-coupled comment moves are designed and tested.

## Tests

The minimum regression set is:

- window reads return `comments` inside the clamped range
- blank-cell comments appear in `window.comments` without synthetic sparse cells
- comment-only sheets still expose correct used bounds for window reads
- `insertRow()` shifts comment addresses
- `insertColumn()` shifts comment addresses
- `deleteRow()` removes in-band comments and shifts trailing comments
- `deleteColumn()` removes in-band comments and shifts trailing comments
- author and text survive structural rewrites after save and reload
- `sortRange()` remains explicitly unsupported when comments are inside the range

## Acceptance Criteria

- `SheetWindowSnapshot` exposes worksheet comments for the effective window range
- bounded reads keep blank-cell comments
- returned comments keep the existing `SheetComment` shape and sorted address order
- row and column insert/delete operations keep worksheet comment package parts valid after reload
- existing whole-sheet comment APIs continue to work
