---
kind: epic
id: EPIC-5
title: Operator controls and MCP tool wiring
status: draft
depends_on:
  - EPIC-4
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/adr/0001-workbook-transport-vs-mcp-wire-transport.md
  - path: docs/architecture/adr/0005-com-strict-and-fallback-controls.md
slice: vertical
acceptance_criteria:
  - EXCEL_MCP_TRANSPORT and per-call workbook_transport override with values auto|file|com (US-3, FR-7, ADR 0001).
  - save_after_write exposed where specified; COM default persistence matches FR-8 / target architecture.
  - EXCEL_MCP_COM_STRICT behavior documented and tested per ADR 0005.
created: "2026-04-24"
updated: "2026-04-24"
---

# Epic-5: Operator controls and MCP tool wiring

## Description

Expose **environment defaults** and **optional tool parameters** so operators and integrators can force behavior and predict errors (`US-3`, `US-5`, `FR-7`, `FR-8`). Wire FastMCP handlers through **`RoutingBackend`** using the shared operation contract.

## Objectives

- Avoid confusion between **MCP wire** `transport` and **workbook** transport (`ADR 0001`).
- Integrate **strict COM** semantics (`ADR 0005`) with router tests.

## User stories (links)

- [Story-5-1](../Stories/Story-5-1-environment-defaults-for-workbook-transport-and-com-strict-mode.md)
- [Story-5-2](../Stories/Story-5-2-optional-tool-parameters-workbook-transport-and-save-after-write.md)
- [Story-5-3](../Stories/Story-5-3-wire-mcp-handlers-through-routingbackend-to-backends.md)

## Dependencies (narrative)

Depends on **EPIC-4** for routing, logging, and injectable detection hooks.

## Related sources

- `docs/architecture/adr/0005-com-strict-and-fallback-controls.md` — strict and optional fallback.
- `docs/specs/PRD-excel-mcp-transport-routing.md` — US-3, US-5, FR-7, FR-8.
