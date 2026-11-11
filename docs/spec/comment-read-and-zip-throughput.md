# Comment Read and ZIP Throughput

Status: Done
Date: 2026-11-11

## Background

Profiling the extension's actual 20-row copy/save finds redundant full-sheet indexing during comment reads (about 0.3 seconds), JavaScript ZIP compression (about 0.5 seconds), and intermediate per-cell allocations during full-window reads.

## Goals and Non-Goals

Improve native save/refresh throughput while preserving uncompressed package parts, metadata, comments, formulas, batching semantics, and output types. Keep compression level 6. No new public API, manual ZIP framing, dependency changes, persistent caches, or new authoritative row model.

## Public Surface

Workbook/Sheet/Zip APIs and CLI behavior remain compatible. ZIP compressed bytes may differ; decompressed contents and archive entry order remain unchanged. `writeArchiveData` stays synchronous and returns Uint8Array; `writeArchive` still returns a Promise and propagates filesystem errors.

## Internal Design

- `getComments` calls the existing pending-mutation synchronization, then reads current worksheet XML directly. It does not require row/cell indexing simply to locate comment parts.
- Replace JavaScript DEFLATE with Node zlib `deflateRawSync` at level 6. Use fflate's existing Zip/ZipPassThrough framing and CRC, overriding the compression hook for one complete entry chunk. Preserve previous duplicate-path last-value/first-position semantics. Empty archives, Unicode names, binary data, and invalid names retain ZIP-library validation.
- Construct each window cell once instead of spreading an intermediate CellEntry. Return identical field values and preserve logical-cell filtering, sparse alignment, and shared formula behavior.

## Structural Semantics

No row, column, formula, or range transformations change. Comment reads after nested batches and pending mutations must remain synchronized; comment bounds and styled empty cells must continue affecting windows as before.

## Test Plan

Cover cold comment reads without building a cell index, no-comment worksheets, edits inside/outside batches and row moves, and complete window consistency. ZIP tests compare all extracted bytes with fflate, use an independent ZIP reader, and cover empty archives/entries, Unicode names, binary data, duplicate paths, generators, disk and in-memory APIs, and error propagation. Run native build and full tests, then the extension's actual generated result and CLI validation.

## Acceptance

No new failures beyond the known Windows npm.cmd packaging-launch issue. Complete extracted workbook parts equal the previous verified result. Real extension save/refresh improves over about 5.5 seconds without profiler overhead. Document actual timings and limitations before completion.

## Risks and Rollback

Keep fflate in charge of ZIP structure and validation. Do not treat compressed byte equality as a lossless requirement. Preserve the current mutation flush before comment reads. Revert the scoped native changes to roll back.

## Verification Results

- Build passes. Five new regressions pass, covering cold comments without cell indexing, full windows, nested batches and shifted comments, Unicode/binary/empty/duplicate ZIP entries, invalid names, disk/export APIs, and filesystem errors. The focused native suite has 11 passing tests.
- Full native suite: 220 passing, one unchanged Windows packaging-launch failure caused by spawnSync npm.cmd and its error formatter (`.tmp/native-throughput-tests.log`). No new failures.
- Current-source extension suite: 592 passing and the same 14 pre-existing Windows SVN/path failures. The published-dependency compatibility fallback passes 11 snapshot/structural tests. Extension type checks, lint, and production builds pass.
- Two real extension copies of source rows 608–627 into target rows 607–626 take 4,321/4,360 ms through refresh, versus the earlier 5,395–5,511 ms (about 20% faster). Save is 1,919/1,925 ms; refresh is 2,402/2,435 ms. Compression and disk output take 194/189 ms. CPU profiling, initial setup, and DOM paint are excluded. A subsequent full-suite reproduction takes 4,356 ms.
- All copied/original values, shifted comments, unchanged input files, and complete refreshed models match independent reads. Independent .NET ZIP extraction verifies all 55 parts are byte-identical to the earlier verified result. Native CLI validation reports ok with no differences. Output size is 1,989,611 bytes, versus 2,108,862 before.
- The extension artifact `xlsx-diff-1.7.0-throughput-hotfix.vsix` bundles the optimized native build. Runtime hashes and native zlib markers are verified. No commits or remote publication.
