---
kind: story
id: STORY-4-2
title: RoutingBackend selection matrix with file-only execution path
status: done
parent: EPIC-4
depends_on:
  - STORY-3-1
  - STORY-4-1
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
  - resolve_workbook_backend returns chosen backend plus machine-readable reason (e.g. forced_file, full_name_match) per target architecture.
  - Mocked tests prove auto selects COM when workbook_open_in_excel is true and file when false (NFR-1, PRD AC1).
  - When transport is com and COM execution is not available, behavior matches ADR 0005 (documented error under strict; never silent file write for com+strict).
created: "2026-04-24"
updated: "2026-04-25"
---

# Story-4-2: RoutingBackend selection matrix with file-only execution path

## Description

Implement **`RoutingBackend.resolve_workbook_backend`** per **FR-3**, returning **backend choice + reason**. Until **`ComWorkbookService`** executes real COM writes, **file** remains the executed backend for mutating tools when policy selects file or when COM is unavailable; **`transport=com`** under strict must **not** silently fall back to file (**ADR 0005**).

## User story

As an **integrator**, I want **deterministic routing decisions** with explicit **reason codes** so that **logs and tests** prove the matrix before COM execution exists.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Align interim `com` behavior with **ADR 0005** (no silent file writes when strict and com requested) even if COM executor is not ready: prefer **clear error** over silent file for `transport=com` when strict.
- Wire **tool_kind** for future read parity (**ADR 0003**).

## Dependencies (narrative)

Depends on **STORY-3-1** (file backend) and **STORY-4-1** (open detection port).
