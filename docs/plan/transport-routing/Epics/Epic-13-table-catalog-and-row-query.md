---
kind: epic
id: EPIC-13
title: Table catalog and row query
status: done
depends_on: []
traces_to:
  - path: docs/specs/table-catalog-query-and-views.md
  - path: docs/architecture/adr/0010-mcp-tool-response-envelope.md
  - path: docs/architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md
  - path: docs/architecture/target-architecture.md
slice: vertical
acceptance_criteria:
  - An agent can list native Excel tables in an open workbook and query one table by column, filter, and limit without changing FilterMode, sort order, or hidden columns.
  - File and COM backends return the same row objects for the same table query.
  - A sheet that is not a ListObject can be mapped into regions and queried with the same filter contract, without creating a table.
  - TOOLS.md documents the new read tools, and Linux CI covers the predicate and file-backend behavior with mocked or openpyxl fixtures.
created: "2026-09-23"
updated: "2026-09-23"
---

# Epic-13: Table catalog and row query

## Description

Give agents a database-style read of Excel tables. `list_tables` finds `ListObject`s. `query_table` projects columns and filters rows. `map_sheet_layout` and `query_region` cover sheets that are not tables, such as `stackfleet concept`.

These tools are READ. They must not change what the user sees. Presentation is [Epic-14](Epic-14-excel-view-from-query-spec.md).

Normative behavior: [table catalog, query, and views](../../../specs/table-catalog-query-and-views.md).

## Rough effort

**Total (epic):** approximately **8–14 developer-days**.

| Story | Rough sizing |
|-------|----------------|
| [13-1](../Stories/Story-13-1-list-native-excel-tables.md) | ~2–4 days |
| [13-2](../Stories/Story-13-2-query-table-rows.md) | ~4–6 days |
| [13-3](../Stories/Story-13-3-sheet-layout-and-region-query.md) | ~3–5 days |

## Risks

| Risk | Mitigation |
|------|------------|
| COM filter state is easy to disturb while reading | Tests assert `FilterMode`, sort, and hidden columns are unchanged after every read. Clear tables only through `ShowAllData` if a test must touch a filter. |
| Wide tables (120 columns) blow the MCP payload | `columns` is required for queries over a documented width, or default `limit` stays small and `truncated` is set. |
| File and COM row equality | One predicate function in shared Python. COM only supplies column arrays. |

## User stories

- [Story-13-1](../Stories/Story-13-1-list-native-excel-tables.md) — catalog ListObjects.
- [Story-13-2](../Stories/Story-13-2-query-table-rows.md) — project, filter, and page table rows.
- [Story-13-3](../Stories/Story-13-3-sheet-layout-and-region-query.md) — regions that are not tables.

## Recommended sequencing

1. **13-1** — names and schemas, no cell values.
2. **13-2** — the query agents will actually call.
3. **13-3** — plain ranges, reusing the 13-2 predicate.

## Execution log

### Start — 2026-09-23

**Outcome:** partial

**Progress** — Epic claimed on `feat/epic-13-table-catalog-and-row-query` from `origin/main`. Sequencing is 13-1, then 13-2, then 13-3. Story 13-1 is in progress.

**Evidence** — branch `feat/epic-13-table-catalog-and-row-query`

**Next steps** — Implement `list_tables` (STORY-13-1), then query and region stories.

**Blockers** — none

### Start — 2026-09-23

**Outcome:** partial

**Progress** — STORY-13-1 completed (`list_tables`). STORY-13-2 claimed for `query_table`.

**Evidence** — Story-13-1 Close log. Tests cited there: 141 passed on the catalog suites.

**Next steps** — Finish 13-2, then 13-3.

**Blockers** — none

### Start — 2026-09-23

**Outcome:** partial

**Progress** — STORY-13-2 completed (`query_table`). STORY-13-3 claimed for `map_sheet_layout` and `query_region`.

**Evidence** — Story-13-2 Close log. Shared predicate is in `src/excel_mcp/tables.py`.

**Next steps** — Finish 13-3, then review the epic.

**Blockers** — none

### Close — 2026-09-23

**Outcome:** completed

**Progress** — STORY-13-1, STORY-13-2, and STORY-13-3 are done. Review fixes: `view_spec.where` keeps non-viewable clauses so a later view cannot drop them and show extra rows; island bounds are split so a ListObject inside the used range is not part of a region.

**Evidence** — `python -m pytest -q` → 347 passed, 1 skipped. Review added `test_layout_regions_do_not_cover_table_cells` and updated `test_is_empty_clause_not_viewable`.

**Next steps** — Pull request for `feat/epic-13-table-catalog-and-row-query`. Epic-14 is not in this implementation.

**Blockers** — none
