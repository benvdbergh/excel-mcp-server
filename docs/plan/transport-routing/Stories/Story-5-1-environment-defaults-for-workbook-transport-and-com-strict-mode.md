---
kind: story
id: STORY-5-1
title: Environment defaults for workbook transport and COM strict mode
status: draft
parent: EPIC-5
depends_on:
  - STORY-4-2
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/adr/0001-workbook-transport-vs-mcp-wire-transport.md
  - path: docs/architecture/adr/0005-com-strict-and-fallback-controls.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - EXCEL_MCP_TRANSPORT defaults to auto and is documented as workbook transport, not MCP wire (ADR 0001).
  - EXCEL_MCP_COM_STRICT gates silent fallback per ADR 0005 with tests.
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-5-1: Environment defaults for workbook transport and COM strict mode

## Description

Read **environment configuration** into the routing context: default workbook transport mode and **COM strict** behavior (**US-3**, **US-5**, **ADR 0005**).

## User story

As a **power user**, I want **environment-level controls** so that **servers behave predictably** across sessions without per-call parameters.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Optional **EXCEL_MCP_COM_ALLOW_FILE_FALLBACK** (or chosen name) per ADR 0005—env-first for v1.
- Ensure naming does not collide with FastMCP `mcp.run(transport=...)`.

## Dependencies (narrative)

Depends on **STORY-4-2** for the routing matrix implementation.
