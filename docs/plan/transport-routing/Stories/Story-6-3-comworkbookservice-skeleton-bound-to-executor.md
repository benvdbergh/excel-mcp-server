---
kind: story
id: STORY-6-3
title: ComWorkbookService skeleton bound to executor
status: draft
parent: EPIC-6
depends_on:
  - STORY-6-2
  - STORY-3-1
traces_to:
  - path: docs/architecture/target-architecture.md
  - path: docs/architecture/adr/0002-com-automation-stack.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - ComWorkbookService implements a minimal subset of the shared contract via COM executor (e.g. attach open workbook, no-op or single-cell write) with Windows manual smoke steps documented.
  - Router can delegate executed COM path to ComWorkbookService when com stack installed and workbook open.
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-6-3: ComWorkbookService skeleton bound to executor

## Description

Introduce **`ComWorkbookService`** implementing the **shared contract** for a **thin vertical slice**, proving end-to-end **COM execution** behind the router (**FR-5**).

## User story

As a **developer**, I want a **proven COM path** for at least one operation before filling in the full tool matrix.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Finalize **ADR 0002** choice in the same deliverable as code.
- Use **running** `Excel.Application` only; bind to open workbook by normalized path.

## Dependencies (narrative)

Depends on **STORY-6-2** (executor) and **STORY-3-1** (parity target contract on file side).
