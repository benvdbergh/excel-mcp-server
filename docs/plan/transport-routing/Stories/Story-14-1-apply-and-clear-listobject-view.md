---
kind: story
id: STORY-14-1
title: Apply and clear an in-place ListObject view
status: draft
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
