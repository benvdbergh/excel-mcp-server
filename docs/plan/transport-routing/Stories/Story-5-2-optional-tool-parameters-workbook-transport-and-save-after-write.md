---
kind: story
id: STORY-5-2
title: Optional tool parameters workbook_transport and save_after_write
status: draft
parent: EPIC-5
depends_on:
  - STORY-5-1
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/adr/0001-workbook-transport-vs-mcp-wire-transport.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - Tools expose optional workbook_transport and save_after_write per FR-7 with JSON schema updates and manifest alignment.
  - File path ignores or documents save_after_write semantics per PRD.
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-5-2: Optional tool parameters workbook_transport and save_after_write

## Description

Extend MCP tool schemas with **per-call overrides** for workbook transport and COM save policy (**FR-7**, **FR-8**), using **`workbook_transport`** naming (**ADR 0001**).

## User story

As a **power user**, I want **per-call overrides** so that **individual tool invocations** can force file or COM when debugging.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Update `manifest.json`, `TOOLS.md` (if present), and server registrations consistently.
- Default **save_after_write** aligns with **FR-8** / target architecture (COM default no save until requested).

## Dependencies (narrative)

Depends on **STORY-5-1** for effective default resolution from environment.
