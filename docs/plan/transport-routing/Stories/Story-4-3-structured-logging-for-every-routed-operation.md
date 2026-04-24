---
kind: story
id: STORY-4-3
title: Structured logging for every routed operation
status: draft
parent: EPIC-4
depends_on:
  - STORY-4-2
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
  - Routed calls emit structured fields transport, reason, duration_ms, workbook_path per redaction policy (NFR-3, US-6).
  - Log format is documented for operators.
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-4-3: Structured logging for every routed operation

## Description

Add **observability** on the router boundary so operators can troubleshoot production routing decisions (**NFR-3**, **US-6**).

## User story

As an **integrator**, I want **structured logs** for each routed call so that I can **prove** which backend ran and why.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Integrate with existing logging approach from pre-fork (`excel-mcp.log` vs stdio); avoid breaking MCP stdio JSON-RPC.
- Use vocabulary from **ADR 0001** in field names.

## Dependencies (narrative)

Depends on **STORY-4-2** for the routing decision and reason codes.
