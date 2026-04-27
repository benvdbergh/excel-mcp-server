---
kind: story
id: STORY-3-2
title: Consolidate server.py inline workbook loads through FileWorkbookService
status: done
parent: EPIC-3
depends_on:
  - STORY-3-1
traces_to:
  - path: docs/architecture/pre-fork-architecture.md
  - path: docs/architecture/target-architecture.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - No remaining inline load_workbook in server.py for routed tools except explicitly documented temporary exceptions with linked follow-up.
  - get_data_validation_info and similar paths use the façade or shared loader.
created: "2026-04-24"
updated: "2026-04-25"
---

# Story-3-2: Consolidate server.py inline workbook loads through FileWorkbookService

## Description

Eliminate **distributed file access** called out in pre-fork architecture by routing handler code through **`FileWorkbookService`** (target architecture §5, migration step 2).

## User story

As a **maintainer**, I want **one code path** for file workbook access so that **routing and logging** apply uniformly.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Prefer mechanical moves; defer risky lifecycle refactors to **STORY-3-3**.
- Update any tests that reached modules directly if signatures change.

## Dependencies (narrative)

Depends on **STORY-3-1** for the façade skeleton.
