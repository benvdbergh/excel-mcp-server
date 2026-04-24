---
kind: epic
id: EPIC-1
title: Tool inventory and shared workbook operation contract
status: draft
depends_on: []
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/excel-mcp-fork-com-vs-file-routing.md
  - path: docs/architecture/pre-fork-architecture.md
slice: vertical
acceptance_criteria:
  - Every MCP tool in scope is classified read vs write (or exempt) with traceability to the blueprint tool table.
  - A single internal contract (protocol, ABC, or facade interface) lists operations both backends must implement for routed tools.
created: "2026-04-24"
updated: "2026-04-24"
---

# Epic-1: Tool inventory and shared workbook operation contract

## Implementation and verification

Tests use **descriptive file names**; stories link to those tests (not the reverse).

| Deliverable | Code | Acceptance tests |
|-------------|------|------------------|
| Classified MCP tool inventory | [`src/excel_mcp/routing/tool_inventory.py`](../../../../src/excel_mcp/routing/tool_inventory.py) | [`tests/test_authoritative_mcp_tool_inventory.py`](../../../../tests/test_authoritative_mcp_tool_inventory.py) |
| Shared workbook operation contract | [`src/excel_mcp/routing/workbook_operation_contract.py`](../../../../src/excel_mcp/routing/workbook_operation_contract.py) | [`tests/test_shared_workbook_operation.py`](../../../../tests/test_shared_workbook_operation.py) |

Package re-exports: [`src/excel_mcp/routing/__init__.py`](../../../../src/excel_mcp/routing/__init__.py).

## Description

Establish the **authoritative tool inventory** and the **narrow shared API** (`FR-4`) before extracting `FileWorkbookService` or adding COM. This epic reduces rework by fixing the method surface that `FileWorkbookService` and `ComWorkbookService` must share.

## Objectives

- Map each existing `@mcp.tool` handler to **read**, **write**, or **exception** (per ADR 0004) for routing policy.
- Document gaps where `server.py` bypasses shared workbook access (e.g. inline `load_workbook`).
- Freeze the **operation contract** for phased COM parity (target architecture §6).

## User stories (links)

- [Story-1-1](../Stories/Story-1-1-authoritative-mcp-tool-inventory-read-vs-write-vs-v1-exception.md)
- [Story-1-2](../Stories/Story-1-2-define-shared-workbook-operation-contract-file-and-com.md)

## Dependencies (narrative)

None. This epic is the planning and contract gate for **Gate 2 — Build ready** in the PRD.

## Related sources

- `docs/specs/PRD-excel-mcp-transport-routing.md` — FR-4, tool routing expectations, traceability note.
- `docs/architecture/target-architecture.md` — tool routing defaults, phased parity.
- `docs/architecture/adr/0004-chart-pivot-com-parity-scope.md` — v1 exceptions for chart/pivot.
