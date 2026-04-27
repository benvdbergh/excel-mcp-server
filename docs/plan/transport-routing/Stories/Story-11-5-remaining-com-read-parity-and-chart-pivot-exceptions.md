---
kind: story
id: STORY-11-5
title: Remaining COM read parity and chart/pivot file-forced documentation
status: draft
parent: EPIC-11
depends_on:
  - STORY-11-3
traces_to:
  - path: docs/architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md
  - path: docs/architecture/adr/0004-chart-pivot-com-parity-scope.md
  - path: docs/architecture/com-read-class-tools-design.md
slice: vertical
invest_check:
  independent: false
  negotiable: true
  valuable: true
  estimable: true
  small: false
  testable: true
acceptance_criteria:
  - All remaining read-class ComWorkbookService methods required by the workbook operation contract are implemented or intentionally delegated with explicit errors; parity with FileWorkbookService for metadata, merged cells, validation, range validation, and formula syntax reads per inventory.
  - Chart/pivot and other V1_FILE_FORCED tools remain file-backed; TOOLS.md and README explain why COM-first does not apply per ADR 0004.
  - Regression tests cover read routing for file-forced kinds versus normal READ tools.
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-11-5: Remaining COM read parity and chart/pivot file-forced documentation

## Description

Complete **COM read parity** for **non-range-primary** read operations. Cross-check **`tool_inventory.py`** for **ToolKind.V1_FILE_FORCED** entries and ensure **ADR 0004** exceptions are **visible** in routing tests and operator docs so agents do not assume COM execution for chart/pivot-era tools.

## User story

As an **operator**, I want **clear documentation** for **which tools always use file-backed execution** so **routing surprises** do not appear in automation.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Maps to **Epic-10 Story-10-4** breadth under the Epic-11 architecture.
- Coordinate acceptance wording with **Story-11-1** so tests do not contradict file-forced rules.

## Dependencies (narrative)

Requires **Story-11-3**. Parallel with **Story-11-4** after wiring lands.
