---
kind: story
id: STORY-6-2
title: Single-thread COM executor compatible with sync MCP handlers
status: done
parent: EPIC-6
depends_on:
  - STORY-6-1
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/target-architecture.md
slice: vertical
invest_check:
  independent: true
  negotiable: false
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - COM calls are submitted to a dedicated worker thread/queue; callers block for results (FR-6).
  - Threading model documented in README or architecture docs.
  - Tests validate serialization without Excel using mocks where possible.
created: "2026-04-24"
updated: "2026-04-27"
---

# Story-6-2: Single-thread COM executor compatible with sync MCP handlers

## Description

Implement the **COM execution model** from target architecture §7: a **single-thread** worker honoring COM apartment rules while FastMCP handlers remain synchronous.

## User story

As an **operator**, I want **stable COM behavior** under concurrent MCP requests so that **Excel automation does not corrupt host state**.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Avoid starting Excel in executor init (**FR-10**).
- Consider shutdown hooks for clean Excel detach when process exits (document limitations).

## Dependencies (narrative)

Depends on **STORY-6-1** for guarded COM imports.

## Delivered

- `src/excel_mcp/com_executor.py`: `ComThreadExecutor` (queue + worker, blocking `submit`, `shutdown`).
- `tests/test_com_executor.py`: serialization and lifecycle without Excel.
- README: COM execution threading subsection.
