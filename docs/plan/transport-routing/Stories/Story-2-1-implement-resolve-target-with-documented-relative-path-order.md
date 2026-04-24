---
kind: story
id: STORY-2-1
title: Implement resolve_target with documented relative-path order
status: draft
parent: EPIC-2
depends_on:
  - STORY-1-2
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
  - resolve_target returns normalized absolute path; resolution order for relative paths is documented (FR-1).
  - Unit tests cover absolute input and at least one relative resolution case per product decision.
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-2-1: Implement resolve_target with documented relative-path order

## Description

Implement **`resolve_target(path)`** (or equivalent) as the single normalization entry used downstream by allowlist, file backend, and COM open detection (**FR-1**, target architecture §1).

## User story

As an **operator**, I want **consistent path identity** so that **auto routing** picks the correct backend when paths look different but mean the same file (**US-4**).

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Document interaction with existing `get_excel_path` for stdio vs SSE/HTTP; avoid double-normalization bugs.
- Coordinate with **ADR 0001** naming in public docs vs internal function names.

## Dependencies (narrative)

Depends on **STORY-1-2** so inputs and call sites for resolution are known.
