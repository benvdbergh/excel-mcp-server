---
kind: story
id: STORY-1-1
title: Authoritative MCP tool inventory (read vs write vs v1 exception)
status: draft
parent: EPIC-1
depends_on: []
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/excel-mcp-fork-com-vs-file-routing.md
  - path: docs/architecture/adr/0004-chart-pivot-com-parity-scope.md
slice: vertical
invest_check:
  independent: true
  negotiable: false
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - Authoritative structured inventory (code) lists each tool from the pre-fork inventory with classification read, write, or v1-exception per ADR 0004.
  - Classifications trace to PRD tool routing section and blueprint §6.
  - Automated acceptance tests live under a descriptive test module (see Verification); story docs reference that module by path, not by story id in the filename.
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-1-1: Authoritative MCP tool inventory (read vs write vs v1 exception)

## Description

Produce the **single source of truth** for which MCP tools participate in workbook transport routing and how (write-class vs read-class vs tool-forced file per **ADR 0004**).

## User story

As a **maintainer**, I want a **classified tool inventory** linked to the PRD and blueprint so that **routing and parity work** does not drop or mis-scope tools.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Start from `docs/architecture/pre-fork-architecture.md` tool count and `server.py` registrations.
- Align default read behavior with **ADR 0003** (file-backed reads in v1).

## Verification

- **Implementation:** [`src/excel_mcp/routing/tool_inventory.py`](../../../../src/excel_mcp/routing/tool_inventory.py)
- **Acceptance tests (descriptive name):** [`tests/test_authoritative_mcp_tool_inventory.py`](../../../../tests/test_authoritative_mcp_tool_inventory.py)

## Dependencies (narrative)

None.
