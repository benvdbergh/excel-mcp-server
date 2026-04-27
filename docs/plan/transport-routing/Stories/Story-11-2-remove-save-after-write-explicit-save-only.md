---
kind: story
id: STORY-11-2
title: Remove save_after_write; persistence via save_workbook only
status: draft
parent: EPIC-11
depends_on: []
traces_to:
  - path: docs/architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md
  - path: docs/architecture/adr/0003-read-path-com-parity.md
  - path: src/excel_mcp/routing/routing_env.py
slice: vertical
invest_check:
  independent: true
  negotiable: false
  valuable: true
  estimable: true
  small: false
  testable: true
acceptance_criteria:
  - save_after_write is removed from all MCP tool schemas and handler surfaces; no remaining env-driven default that implicitly saves after every mutation (pattern aligned with ADR 0008 removal of EXCEL_MCP_SAVE_AFTER_WRITE / effective_save_after_write).
  - save_workbook remains the sole routine persistence operation for operator intent; behavior documented for COM dirty state vs disk.
  - Audit completes across server tool registration and routing helpers so no stray parameter or dead env hook remains without documentation or removal note in CHANGELOG or README.
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-11-2: Remove `save_after_write`; persistence via `save_workbook` only

## Description

Execute **ADR 0008** §2 persistence model: **strip `save_after_write`** from **all** mutating tools and from **environment defaults**, leaving **`save_workbook`** as the **only** explicit persistence control. Align **[ADR 0003](../../../architecture/adr/0003-read-path-com-parity.md)** explicit-save anchor with the removal of per-call save flags.

## User story

As an **agent author**, I want **explicit save moments** so that **flushes to disk** are visible in scripts and audits.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Prefer sequencing early in the epic so downstream lifecycle and read stories do not depend on deprecated parameters.
- Coordinate README and TOOLS.md updates with **Story-11-7** or land minimal deltas here.

## Dependencies (narrative)

Can run in parallel with **Story-11-1** after a short design sync; must complete before treating Epic-11 as release-ready.
