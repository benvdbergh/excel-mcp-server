---
kind: story
id: STORY-13-1
title: List native Excel tables
status: done
parent: EPIC-13
depends_on: []
traces_to:
  - path: docs/specs/table-catalog-query-and-views.md
    anchor: "#list_tables"
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - list_tables returns every ListObject in the workbook with sheet, name, range, header range, data range, row count, and whether a filter is applied.
  - detail=minimal omits column names. detail=schema includes them. The payload contains no cell values.
  - File and COM implementations share the JSON shape. COM runs on the existing executor. A workbook with no tables returns an empty list.
  - Reading the catalog does not change FilterMode, sort order, or hidden columns.
  - TOOLS.md and the tool inventory register the tool as READ.
created: "2026-09-23"
updated: "2026-09-23"
---

# Story-13-1: List native Excel tables

## Description

Expose workbook `ListObject`s through `list_tables`, on both the file and COM backends. This is the catalog step before any query. On the sample register the result includes `ap_sw_components` on `software components` and nothing on `stackfleet concept`.

## User story

As an **agent**, I want **the names and columns of every Excel table** so **I can query a table without guessing a sheet range**.

## Technical notes

- Extend `WorkbookOperation` and both workbook services. Register the tool in `tool_inventory.py` and `server.py`.
- COM walks each worksheet's `ListObjects`. Filter-applied is `ListObject.AutoFilter.FilterMode`, not `Worksheet.AutoFilterMode`.
- File mode uses openpyxl worksheet tables.

## Execution log

### Start — 2026-09-23

**Outcome:** partial

**Progress** — Claimed before implementation. `list_tables` on file and COM backends, registered as READ, with no cell values and no filter/sort/hidden-column mutation.

**Evidence** — branch `feat/epic-13-table-catalog-and-row-query`

**Next steps** — Implement the catalog, tests, and TOOLS.md, then close.

**Blockers** — none

### Close — 2026-09-23

**Outcome:** completed

**Progress** — Implemented `list_tables` (contract + file/COM backends + MCP READ tool). Shared JSON catalog shape via `tables.py` helpers. Default `detail=schema`. COM reads `ListObject.AutoFilter.FilterMode` only; no ShowAllData/AutoFilter/Sort/hide mutations.

**Evidence** —
- Files: `src/excel_mcp/tables.py`, `src/excel_mcp/routing/workbook_operation_contract.py`, `src/excel_mcp/routing/file_workbook_service.py`, `src/excel_mcp/routing/com_workbook_service.py`, `src/excel_mcp/routing/tool_inventory.py`, `src/excel_mcp/server.py`, `TOOLS.md`, tests under `tests/test_file_workbook_service.py`, `tests/test_com_workbook_service.py`, `tests/test_shared_workbook_operation.py`, `tests/test_authoritative_mcp_tool_inventory.py`
- Tests: `python -m pytest tests/test_file_workbook_service.py tests/test_com_workbook_service.py tests/test_shared_workbook_operation.py tests/test_authoritative_mcp_tool_inventory.py -q` → **141 passed**

**Next steps** — none (story done; Epic-13 remains open for 13-2 / 13-3)

**Blockers** — none
