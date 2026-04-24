---
kind: story
id: STORY-3-3
title: Workbook lifecycle consistency (close on reads) where low-risk
status: draft
parent: EPIC-3
depends_on:
  - STORY-3-2
traces_to:
  - path: docs/architecture/pre-fork-architecture.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: false
  small: false
  testable: true
acceptance_criteria:
  - Documented improvements to wb.close patterns for read helpers where safe; no user-visible regression in file mode.
  - Any deferred items listed as debt with rationale.
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-3-3: Workbook lifecycle consistency (close on reads) where low-risk

## Description

Pay down **workbook lifecycle** inconsistencies noted in pre-fork architecture (read paths that omit `close`) as optional hardening within the façade (**target architecture** migration).

## User story

As an **operator**, I want **predictable resource usage** when agents perform many read operations against large files.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Mark as **negotiable / optional** within the epic if schedule pressure; keep scope small per file.
- Coordinate with tests for `read_excel_range` and similar.

## Dependencies (narrative)

Depends on **STORY-3-2** so all relevant paths flow through the façade.
