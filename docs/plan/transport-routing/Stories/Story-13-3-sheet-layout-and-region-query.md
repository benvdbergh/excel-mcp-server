---
kind: story
id: STORY-13-3
title: Map sheet layout and query a non-table region
status: done
parent: EPIC-13
depends_on:
  - STORY-13-2
traces_to:
  - path: docs/specs/table-catalog-query-and-views.md
    anchor: "#map_sheet_layout-and-query_region"
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - map_sheet_layout lists native tables and the used-range islands outside them, each with bounds and a header guess.
  - query_region runs the Story 13-2 column, where, limit, and offset contract against a region id or an explicit range.
  - The call does not create a ListObject and does not change FilterMode, sort, or hidden columns.
  - A fixture shaped like a header row that is not row 1 still yields that header row rather than a blank first row.
  - TOOLS.md describes both tools as READ.
created: "2026-09-23"
updated: "2026-09-23"
---

# Story-13-3: Map sheet layout and query a non-table region

## Description

Some sheets are not `ListObject`s. `map_sheet_layout` separates real tables from the other occupied blocks. `query_region` reuses the 13-2 predicate on one of those blocks. This covers `stackfleet concept` and any grid that only looks like a table.

## User story

As an **agent**, I want **to query a rectangular block that is not an Excel table** so **mixed sheets are usable without a manual range guess**.

## Technical notes

- Subtract ListObject ranges before searching for islands.
- Header guess: the first row of an island that is mostly text. Document the guess in the region payload so the agent can override it with an explicit range.

## Execution log

### Start — 2026-09-23

**Outcome:** partial

**Progress** — Claimed after STORY-13-2. `map_sheet_layout` and `query_region` reuse the shared predicate in `src/excel_mcp/tables.py` (`filter_table_rows`, `build_query_table_payload`). No ListObject creation and no filter/sort/hidden-column mutation.

**Evidence** — branch `feat/epic-13-table-catalog-and-row-query`. `query_table` is done.

**Next steps** — Implement layout mapping and region query, tests, and TOOLS.md, then close.

**Blockers** — none

### Close — 2026-09-23

**Outcome:** completed

**Progress** — `map_sheet_layout` and `query_region` are READ tools on file and COM. Shared island/header helpers and `query_region_rows` live in `tables.py`; `build_query_table_payload` takes a `target` so region queries reuse the same predicate/paging/viewability path as `query_table`. TOOLS.md documents both as READ.

**Evidence** — Files: `src/excel_mcp/tables.py`, `routing/workbook_operation_contract.py`, `file_workbook_service.py`, `com_workbook_service.py`, `tool_inventory.py`, `server.py`, `TOOLS.md`, tests under `tests/test_table_query.py`, `test_file_workbook_service.py`, `test_com_workbook_service.py`, `test_authoritative_mcp_tool_inventory.py`, `test_shared_workbook_operation.py`, `test_server_routing_integration.py`. Command: `python -m pytest tests/test_table_query.py tests/test_file_workbook_service.py tests/test_com_workbook_service.py tests/test_authoritative_mcp_tool_inventory.py tests/test_shared_workbook_operation.py tests/test_server_routing_integration.py -q` → **173 passed**.

**Next steps** — Orchestrator reviews EPIC-13 (do not close here).

**Blockers** — none
