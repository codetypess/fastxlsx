---
name: spec-driven-development
description: Use when implementing or changing a non-trivial fastxlsx feature. Require a written spec before code changes for API, CLI, metadata, XML, structural edit, or lossless-roundtrip-sensitive work.
---

# Spec-Driven Development

Use this skill for non-trivial feature work in this repository.

Before changing code, confirm there is an SDD document for the feature.

- Preferred feature spec location: `docs/spec/<feature-name>.md`
- Repository contract: `docs/spec-driven-development.md`
- Spec index: `docs/spec/README.md`

## Required Gate

Do not start implementation until the spec covers:

- background and current gap
- goals and non-goals
- public API or CLI surface, if affected
- structural edit semantics when rows, columns, ranges, or cells are involved
- test plan and acceptance criteria

## When This Is Required

Use the spec gate for changes that touch:

- public API or TypeScript types
- CLI behavior or JSON output
- workbook or worksheet metadata
- XML read or write helpers
- row or column structural transforms
- comments, formulas, validations, filters, merges, protection, styles, tables, or defined names
- any work where lossless roundtrip could regress

## Exceptions

You do not need a dedicated feature spec for:

- typo-only edits
- comment-only edits
- isolated refactors with no behavior change
- narrowly scoped test cleanup without feature changes

## Implementation Order

Default order:

1. Write or update the spec.
2. Get alignment on the spec.
3. Implement code.
4. Add or update tests.
5. Update public docs if the user-facing surface changed.
