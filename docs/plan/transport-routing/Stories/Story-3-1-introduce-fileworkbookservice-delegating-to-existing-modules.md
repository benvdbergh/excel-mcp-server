---
kind: story
id: STORY-3-1
title: Introduce FileWorkbookService delegating to existing modules
status: done
parent: EPIC-3
depends_on:
  - STORY-1-2
  - STORY-2-2
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
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
  - FileWorkbookService exists and implements the shared contract for a first slice of operations (expandable).
  - Unit tests cover the façade delegation without requiring Excel.
created: "2026-04-24"
updated: "2026-04-25"
---

# Story-3-1: Introduce FileWorkbookService delegating to existing modules

## Description

Create **`FileWorkbookService`** as a thin layer over existing `workbook`, `sheet`, `data`, etc., matching the contract from **STORY-1-2** (**FR-4**).

## User story

As a **maintainer**, I want **one file-side service** so that **routing** can swap backends without scattering `load_workbook` calls.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Keep behavior identical to pre-fork for the first migrated tools; use characterization tests where helpful.
- Path input must come from **resolve_target** + allowlist path (**STORY-2-2**).

## Dependencies (narrative)

Depends on **STORY-1-2** (contract) and **STORY-2-2** (policy-aligned path).
