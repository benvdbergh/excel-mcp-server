---
kind: story
id: STORY-12-2
title: MCP — discovery tool handler, schema, and routing classification
status: draft
parent: EPIC-12
depends_on:
  - STORY-12-1
traces_to:
  - path: docs/architecture/adr/0009-open-workbook-discovery-tool.md
  - path: docs/architecture/adr/0001-workbook-transport-vs-mcp-wire-transport.md
slice: vertical
invest_check:
  independent: false
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - New MCP tool is registered (exact name aligned with operator docs—e.g. excel_list_open_workbooks) with JSON result string consistent with Story 12-1; optional parameters reserved for future detail levels are either absent or no-op with documented placeholders.
  - Tool does not require filepath; workbook_transport is N/A or documented as ignored; no accidental RoutingBackend filepath resolution for this tool.
  - tool_inventory.py and any routing registries classify the tool per design (SESSION / lifecycle-adjacent or documented exception) so metrics and docs stay coherent.
  - Errors from missing COM stack (non-Windows / no [com]) and from Excel not running match user-actionable messaging used elsewhere.
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-12-2: MCP — discovery tool handler, schema, and routing classification

## Description

Expose the Story **12-1** implementation as an **MCP tool**: FastMCP / `server.py` registration, **input schema** (minimal or empty), **output schema** (`result` string carrying JSON). Decide **explicitly** whether **`RoutingBackend`** participates: discovery is **not** a per-file read; the likely path is **handler → COM executor only** with no `resolve_target` unless a future story adds filtering.

## User story

As an **agent author**, I can call **one MCP tool** to list open workbooks and copy a **`FullName`** into **`get_workbook_metadata`** without Python one-liners.

## Technical notes

- **ADR 0001:** do not confuse **MCP wire transport** with **workbook transport**; this tool is still “host introspection,” not `workbook_transport=auto` for a path.
- If Cursor / local **MCP tool JSON** descriptors are generated, update generation or committed descriptors as the project does today (align with Story 12-3 if split).

## Dependencies (narrative)

Requires **STORY-12-1** for the JSON contract and COM entry point.
