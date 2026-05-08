# Agent Instructions

For `.xlsx` tasks, use [.agents/skills/fastxlsx/SKILL.md](.agents/skills/fastxlsx/SKILL.md).
For non-trivial feature development tasks in this repository, use [.agents/skills/spec-driven-development/SKILL.md](.agents/skills/spec-driven-development/SKILL.md).

Prefer the `fastxlsx` CLI over direct workbook XML edits. Inspect before writing, validate after writing, and prefer `--profile` when `table-profiles.json` exists.

Keep this file as a thin Codex routing hook. Put detailed Codex workflow updates in `.agents/skills/fastxlsx/`.
