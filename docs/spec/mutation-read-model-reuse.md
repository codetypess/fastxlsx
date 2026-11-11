# Mutation Read Model Reuse

Status: Done
Date: 2026-11-11

## Background

Structural writes already parse final cell state, and batch setters retain final snapshots in a sparse overlay. Finalization discards the index; full-window reads then parse the same cells from serialized XML. An extension fixture spends about 1.71 seconds applying edits and 1.70 seconds rebuilding snapshots.

## Goals

Share finalized mutation cell data with full-window reads, retaining native reference-transform semantics and archive bytes. Give consumers an authoritative token for caching complete sheet reads.

## Non-goals

No new formula evaluator, structural transform rules, XML writer, incremental diff algorithm or CLI command. This is an additive library/cache capability, not a complete replacement of the indexed XML mutation model.

## Public Surface

Add Workbook.getSheetReadVersion(sheetName): string. It finalizes pending writes and returns an opaque token valid only within that Workbook instance. Equal tokens guarantee equivalent sheet window results and sheet metadata; changed tokens may conservatively invalidate unchanged data. Unknown sheet names throw as getSheet does. Worksheet writes advance that sheet's token; other package writes/removals conservatively advance all tokens. Existing batch/set/save/window signatures stay unchanged.

## Internal Model

Publish compact ordered row metadata and physical cell snapshots from a completed mutation index plus its final pending-cell overlay. The read model has no XML offsets and does not retain LocatedCell/LocatedRow indexes. The full-window reader can derive logical bounds and return cells/layout from this data without parsing cell XML again. Raw XML fallback remains available for every invalidated or unsupported state. Shared-formula anchor changes use conservative fallback when a pending overwrite could change followers.

Prepared cell models are separate from derived layout caches. A raw worksheet write/removal discards its model; shared-string/relationship changes discard all prepared cells. Comments/layout/style metadata are read from final package state. Models are weakly owned by Workbook and replaced per sheet, without accumulating edit histories.

## OOXML Mapping

Existing writers retain ownership of worksheet cells, formula/reference transforms, dimensions, tables, comments, names and package parts. Read-model publication changes no XML. Explicit blank-cell alignments inside requested windows are preserved without changing logical used bounds; the old scanner and prepared path follow the same rule.

## Mutation Semantics

The final overlay wins per address; deletions remove physical cells. Missing rows are added in numeric order. Metadata comes from the finalized sheet XML and current relationships. Reads inside a batch must observe pending changes and cannot return a stale prepared model. Returned window structures remain detached from mutable cache data.

## Structure Transform Semantics

insertRow/deleteRow/insertColumn/deleteColumn continue to use existing native transforms for moved coordinates, local/cross-sheet formulas, names, merges, filters, comments and tables. Publish after these transforms. Do not replay coordinate transforms independently in the extension. A read-version token includes any write to another sheet caused by these transforms.

## Compatibility

Existing APIs and file output are preserved. The extension feature-detects the new version API and retains full reads for published fastxlsx 1.7.1. Its local blank-alignment compatibility patch is bypassed only for the reviewed native implementation. No dependency manifest changes, commits, installation or publishing.

## Test Matrix

New native tests compare cached reads with fresh archive reads for cell types, blanks/styles, rows/columns, repeated edits, intermediate reads, shared formulas, cross-sheet references, comments/metadata, raw writes and removals. Test version scope/invalidation and returned-value detachment. Use existing lossless/structure/interop/XML/real-file suites as regressions. Extension tests compare complete snapshots and worker results to independent reopened snapshots. Validate the original 20-row fixture, unchanged original hashes and all 55 decompressed package entries.

## Acceptance

Native build passes and full tests add no failures beyond the existing Windows npm.cmd packaging-launch failure. The extension builds and tests against the compiled native checkout. Record save/refresh timings, event-loop response and memory; do not claim total latency improvement until measured.

Shared-formula fallback is scoped to pending overwrites/deletions of existing shared cells. Other pending changes reuse resolved followers from the finalized native index; the real fixture inserts new rows beside over 90,000 shared-formula cells.

## Verification Results

Native build passes. All 19 added differential tests pass, including all structural directions with shared followers, anchor/follower fallback, intermediate reads, raw replacement/removal, shared-string invalidation, blank alignments and detached error/comment results. The full suite reports 239 passing and the same existing Windows npm.cmd packaging-launch failure, with no new failures. Both READMEs document the additive API.

Extension integration reports 163 focused tests passing; the full suite reports 632 passing and the same 14 baseline Windows path/SVN failures. Published 1.7.1 compatibility reports 30 passing and 9 capability-specific skips. The original fixture hashes are unchanged; all 55 decompressed output parts match the established output, and CLI validate reports no differences.

The production worker copies 20 rows/295 values: save acknowledgement 1,774 ms, refresh 2,278 ms, total 4,052 ms, versus 1,979 / 2,702 / 4,681 ms before. Snapshot creation is 1,386 ms versus 1,697 ms. Maximum sampled host heartbeat gap is 17 ms, and sampled peak RSS is 1,522,888,704 bytes (previous 1,493,405,696); no memory reduction or percentile guarantee is claimed. The fixture changes names/comments, conservatively invalidating all sheet tokens, so its primary benefit is prepared read data on the large shared-formula sheet.

The extension artifact xlsx-diff-1.7.0-mutation-read-hotfix.vsix contains the compiled native changes; all six runtime hashes match the final build. No dependency installation, commit or publishing was performed. Changed-sheet object construction, shared formula materialization, native structure/reference transforms and full diff alignment remain costs.
