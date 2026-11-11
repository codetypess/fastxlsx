# Structural Save Performance

Status: Done
Date: 2026-11-10

## Background

Copying 20 missing rows into a 12,000-row skill sheet takes over a minute per single-row insertion with FastXLSX 1.7.1. A checked compatibility patch in vscode-xlsx-diff reduces a 20-row save to approximately 2.7 seconds. Profiling still shows repeated index construction and eager expansion of shared formula followers during writes that never inspect formula text. This work moves the general fixes into the library and removes that remaining eager work.

## Goals

- Assemble structural row rewrites in one pass over the original string.
- Reuse row/column indexes instead of allocating a second full cell address Map.
- Finalize batches through sparse row mutations, rebuilding the index only on a later read.
- Materialize shared follower formula text only when a caller reads it, retaining cached-value reads and formula output semantics.
- Improve the same real 20-row save beyond the extension's 2.7-second baseline.

## Non-goals

- New public API, CLI behavior, or package publishing.
- A second authoritative workbook model or a piece table/tree implementation.
- Changing formula evaluation, cached results, alignment, or merge semantics.

## Public Surface

Workbook/Sheet APIs, returned cell fields, coordinate validation, error handling, and CLI JSON remain compatible. Lazy formula text is memoized within the index; serializing or spreading a cell snapshot still materializes its formula string.

## Internal Model

Keep the row Map, sparse per-row column array, parsed snapshots, and original string offsets. Structural writers produce ordered string fragments. Sparse pending-cell overlays reuse unchanged rows for used-bounds calculation. Finalization writes normalized bytes and invalidates stale offsets instead of immediately parsing them again. Shared follower snapshots capture the immutable anchor formula and coordinate delta; the first formula read resolves and caches the string. Used-bounds classification recognizes formula cells without forcing text expansion.

## OOXML Mapping

Reuse existing cell/row, formula, range, comment, table, defined-name, and worksheet metadata transformations. Preserve original gaps between row spans. Unknown or untouched package parts must stay byte-identical. This change does not introduce new XML representations.

## Mutation Semantics

Nested batches and intermediate reads continue to flush pending writes. Value/style/formula mutation precedence remains unchanged. A later read after finalization rebuilds an authoritative index; old snapshots retain their own original anchor/delta semantics. Missing shared anchors retain current behavior.

## Structure Transform Semantics

Row and column insert/delete continue updating formulas, shared formula refs, merges, comments, hyperlinks, data validation, worksheet/table filters, defined names, and other sheet formulas using existing helpers. Deferred formula expansion must use the transformed XML's anchors and coordinates on the rebuilt index.

## Compatibility

No package version bump or remote publish is required for development. Build a local package for extension integration. Remove extension compatibility patches from native-library builds, retaining checked fallbacks for the published 1.7.1 dependency until a library release is available.

## Test Matrix

- Full library tests, TypeScript build, and package checks.
- Structural insert/delete rows and columns with sparse rows and metadata.
- Batch dimension growth/shrinkage, empty placeholders, deletion, repeated batches, and intermediate reads.
- Shared formulas with followers preceding anchors, cached values, memoized formula reads, serialization, frozen/sealed snapshots, edits, and structural moves.
- Real supplied workbooks: compare archive entries with the previous extension fix, validate roundtrip, and measure save/reload timings without modifying source fixtures.

## Acceptance

- Build, focused regressions, and applicable full-suite tests pass.
- Shared follower formula output matches existing eager expansion.
- Real copied-row output is archive-equivalent to the previous fix, and its save is faster than the 2.7-second baseline.
- An updated local VSIX bundles the locally built library, while the published-dependency fallback remains buildable.

## Validation Results

- `npm run build` passes. Full library suite: 215 passing, one existing Windows packaging-launch failure. The failing test and `pack:check` call `spawnSync("npm.cmd", ...)`, which cannot start on this Node/Windows environment; their error handling then dereferences undefined output. Neither launcher was changed by this work.
- Six new regressions pass: lazy cached-value/bounds reads, late/missing anchors and mixed references, freezing/sealing before formula materialization, retained snapshots across all four structural transforms, sparse batch/intermediate-read equivalence, and a 12,000-row rewrite preserving physical row gaps and untouched parts.
- Shell-based `npm pack` succeeds; the final archive contains all 118 expected files. Extracted JS, type, CLI help/version, and workbook validation entrypoints are verified. The local package is `.tmp/fastxlsx-1.7.1.tgz`; nothing was published.
- Extension integration uses `FASTXLSX_LOCAL_ROOT` to bundle the compiled native library without changing installed dependencies or dependency manifests. Full VS Code suite: 582 passing and the same 14 pre-existing Windows SVN/path failures. Final focused integration: 19 passing. The default published-dependency compatibility fallback also passes its 12 focused tests.
- On the supplied workbook, the real host copies source rows 608–627 into target rows 607–626 with 295 cell writes. Final write/save: 2,249 ms versus 2,668 ms before the native shared-formula improvement (about 16% faster). Target reread: 2,254 ms; diff rebuild: 695 ms. These sum to 5,198 ms, excluding initial comparison setup and DOM rendering.
- All 55 output archive entries are byte-identical to the first extension fix. Copied display values, original target values, other sheets, comments, and unchanged source files are verified. The packed CLI returns `ok: true` and no roundtrip differences.
- Updated extension artifact: `vscode-xlsx-diff/xlsx-diff-1.7.0-fastxlsx-save-hotfix.vsix`, bundling the optimized native code.

- Follow-up (2026-11-11): [Comment Read and ZIP Throughput](comment-read-and-zip-throughput.md) removes redundant comment-read indexing, uses native compression, and reduces full-window allocations. Actual extension save-through-refresh is 4,321–4,360 ms; all 55 extracted package parts match the previous output.
