---
kind: story
id: STORY-14-1
title: Apply and clear an in-place ListObject view
status: done
parent: EPIC-14
depends_on:
  - STORY-13-2
traces_to:
  - path: docs/specs/table-catalog-query-and-views.md
    anchor: "#apply_table_view"
  - path: docs/architecture/adr/0002-com-automation-stack.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - apply_table_view in_place on a ListObject applies viewable where clauses through AutoFilter, an optional sort through ListObject.Sort, and column focus by hiding columns that were visible.
  - A spec with limit, offset, or a cross-column OR is rejected and leaves the sheet unchanged.
  - The response includes a restore_token. clear_table_view returns the prior filter and sort, and unhides only columns this call hid.
  - Clearing uses ListObject.AutoFilter.ShowAllData(). A table with FilterMode true and AutoFilterMode false is cleared successfully.
  - The tool is WRITE and COM-only. File transport returns a clear error. TOOLS.md states that the desktop view is shared with co-authors.
created: "2026-09-23"
updated: "2026-09-23"
---

# Story-14-1: Apply and clear an in-place ListObject view

## Description

Show a `view_spec` from `query_table` on the table the user already has open, then put the sheet back. Equality, contains, comparisons, single-column `in`, AND across fields, sort, and column focus are in scope. Limit and offset are not.

## User story

As an **agent**, I want **to show a filtered and sorted table in Excel after the user agrees** so **they see the same rows I queried**.

## Technical notes

- Run on the COM executor ([ADR 0002](../../../architecture/adr/0002-com-automation-stack.md)).
- Field index is relative to the table's header range, not to column A. For `ap_sw_components`, `optionProvider` is field 6 because the table starts at column B.
- Excel stores a two-value filter as `xlOr` with `Criteria1` and `Criteria2`.
- If the existing sort cannot be captured well enough to restore, refuse the sort portion rather than reordering permanently.

## Execution log

### Start — 2026-09-23

**Outcome:** partial

**Progress** — Claimed STORY-14-1. Implement COM-only `apply_table_view` (in-place ListObject) and `clear_table_view`: AutoFilter, optional sort, column focus, restore token. Reject limit, offset, and cross-column OR without changing the sheet. File transport must return a clear error.

**Evidence** — branch `feat/epic-14-excel-view-from-query-spec`. Depends on `view_spec` from `query_table` in `src/excel_mcp/tables.py`.

**Next steps** — Implement, test, refactor, then close.

**Blockers** — none

### Close — 2026-09-23

**Outcome:** completed

**Progress** — COM-only in-place ListObject `apply_table_view` and `clear_table_view`: AutoFilter (field index relative to the table), optional sort, column focus, and `restore_token`. Limit, offset, cross-column OR, and non-viewable operators are refused with no sheet mutation. Clear uses `ListObject.AutoFilter.ShowAllData()` and unhides only columns this call hid. Snapshot and plain-range targets return a clear error. File transport returns a COM-required error. TOOLS.md states the desktop view is shared.

**Evidence** — `python -m pytest tests/test_table_query.py tests/test_file_workbook_service.py tests/test_com_workbook_service.py tests/test_authoritative_mcp_tool_inventory.py tests/test_shared_workbook_operation.py tests/test_server_routing_integration.py -q` → 190 passed. Files: `src/excel_mcp/tables.py`, `routing/{workbook_operation_contract,file_workbook_service,com_workbook_service,tool_inventory}.py`, `server.py`, `TOOLS.md`, and the tests above.

**Next steps** — STORY-14-2 plain-range view, then STORY-14-3 snapshot.

**Blockers** — none
