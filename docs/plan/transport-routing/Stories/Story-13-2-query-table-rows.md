---
kind: story
id: STORY-13-2
title: Query table rows by column and filter
status: done
parent: EPIC-13
depends_on:
  - STORY-13-1
traces_to:
  - path: docs/specs/table-catalog-query-and-views.md
    anchor: "#query_table"
  - path: docs/architecture/adr/0010-mcp-tool-response-envelope.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - query_table accepts a ListObject name, a column list, AND-combined where clauses (eq, neq, contains, in, gt, gte, lt, lte, is_empty), plus limit and offset.
  - The response is headers plus row objects, with row_count, truncated, view_spec, and view_applicability. Limit and offset are marked not viewable. in on one column is viewable.
  - File and COM return the same rows for the same arguments. COM reads only the selected columns.
  - After the call, FilterMode, sort order, and hidden columns match the pre-call state.
  - include_routing_metadata wraps the payload in the ADR 0010 envelope. TOOLS.md documents the arguments and the width guidance.
created: "2026-09-23"
updated: "2026-09-23"
---

# Story-13-2: Query table rows by column and filter

## Description

Implement the shared predicate once, then call it from both backends. `query_table` is how an agent filters a wide table, such as the 120-column component register, down to the columns it needs. The returned `view_spec` is what Epic-14 can show later. This story does not show anything in Excel.

## User story

As an **agent**, I want **filtered table rows as records** so **I can answer from the data without downloading the whole grid**.

## Technical notes

- Resolve the table by name. Names are unique in the workbook (`ap_sw_components` vs a pasted `ap_sw_components3`).
- Unknown table, unknown column, and invalid clause return the existing tool error style.
- Tests can use an openpyxl workbook with a table plus a mocked COM column matrix. No Excel required on Linux CI.

## Execution log

### Start — 2026-09-23

**Outcome:** partial

**Progress** — Claimed after STORY-13-1. Implement shared row predicate and `query_table` on file and COM. Reads must not change FilterMode, sort, or hidden columns. `include_routing_metadata` uses the ADR 0010 envelope.

**Evidence** — branch `feat/epic-13-table-catalog-and-row-query`. Catalog lives in `src/excel_mcp/tables.py` and `list_tables`.

**Next steps** — Implement query, tests, and TOOLS.md, then close.

**Blockers** — none

### Close — 2026-09-23

**Outcome:** completed

**Progress** — Shared `query_table_rows` predicate in `tables.py`; wired through `WorkbookReadOperations`, File/COM services, tool inventory, and MCP `query_table` with ADR 0010 `include_routing_metadata`. COM reads only projection + filter ListColumns. TOOLS.md documents 32-column width threshold and default limit 100.

**Evidence** — `python -m pytest tests/test_table_query.py tests/test_file_workbook_service.py tests/test_com_workbook_service.py tests/test_authoritative_mcp_tool_inventory.py tests/test_shared_workbook_operation.py tests/test_server_routing_integration.py -q` → 162 passed. Files: `src/excel_mcp/tables.py`, `routing/{workbook_operation_contract,file_workbook_service,com_workbook_service,tool_inventory}.py`, `server.py`, `TOOLS.md`, tests above.

**Next steps** — none (STORY-13-3 / Epic-14 remain separate)

**Blockers** — none
