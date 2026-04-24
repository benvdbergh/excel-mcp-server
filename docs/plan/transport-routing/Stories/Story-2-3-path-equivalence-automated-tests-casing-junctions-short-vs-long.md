---
kind: story
id: STORY-2-3
title: Path equivalence automated tests (casing, junctions, short vs long)
status: draft
parent: EPIC-2
depends_on:
  - STORY-2-1
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - At least one automated test demonstrates normalization equivalence per PRD release AC3 and US-4.
  - Windows-only cases are skipped or marked in CI when not applicable (NFR-6).
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-2-3: Path equivalence automated tests (casing, junctions, short vs long)

## Description

Add **focused tests** proving that `resolve_target` treats documented equivalence classes consistently so routing decisions are stable (**US-4**, **NFR-1**).

## User story

As a **maintainer**, I want **regression tests for path edge cases** so that **routing does not drift** across Windows versions and OneDrive layouts.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Optional later iteration: file-id equality; document as out-of-band if not in v1.
- Link failing cases to risk register in PRD (OneDrive aliasing).

## Dependencies (narrative)

Depends on **STORY-2-1** for the normalization implementation under test.
