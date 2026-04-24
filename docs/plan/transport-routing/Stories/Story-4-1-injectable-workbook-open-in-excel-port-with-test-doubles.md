---
kind: story
id: STORY-4-1
title: Injectable workbook_open_in_excel port with test doubles
status: draft
parent: EPIC-4
depends_on:
  - STORY-2-1
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
  - workbook_open_in_excel(resolved_path) is behind an interface injectable in tests; default Windows implementation may be stubbed until Epic 6.
  - Tests run on CI without Excel using fakes (NFR-6).
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-4-1: Injectable workbook_open_in_excel port with test doubles

## Description

Introduce **`workbook_open_in_excel`** behind a port that compares **normalized** full paths to Excel’s open workbooks (**FR-2**), with **injection** for deterministic tests.

## User story

As a **maintainer**, I want **mockable open detection** so that **routing tests** are reliable on Linux CI.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Do not start Excel in default implementation (**FR-10**).
- Non-Windows: module absent or returns false / unsupported per **FR-12** (clarify in router story).

## Dependencies (narrative)

Depends on **STORY-2-1** for normalized path comparison inputs.
