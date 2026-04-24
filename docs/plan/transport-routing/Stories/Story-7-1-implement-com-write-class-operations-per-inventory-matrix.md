---
kind: story
id: STORY-7-1
title: Implement COM write-class operations per inventory matrix
status: draft
parent: EPIC-7
depends_on:
  - STORY-6-3
  - STORY-5-3
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/target-architecture.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: false
  testable: true
acceptance_criteria:
  - Each write-class tool in inventory either executes via ComWorkbookService when router selects COM or documents explicit exception per ADR 0004.
  - US-1 acceptance scenario achievable on Windows with Excel open for supported tools (manual or integration where feasible).
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-7-1: Implement COM write-class operations per inventory matrix

## Description

Expand **`ComWorkbookService`** to cover the **write-class** tools agreed in **Epic 1**, prioritizing grid and formula writes before specialized features (target architecture phased parity).

## User story

As an **operator**, I want **mutations to land in Excel** when my workbook is open so that **the grid updates live** without stale file state (**US-1**).

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Respect **`save_after_write`** default false (**FR-8**).
- Keep **read tools** file-backed (**ADR 0003**).

## Dependencies (narrative)

Depends on **STORY-6-3** (COM path) and **STORY-5-3** (handler wiring).
