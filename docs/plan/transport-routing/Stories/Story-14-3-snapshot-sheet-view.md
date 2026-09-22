---
kind: story
id: STORY-14-3
title: Snapshot a query onto a new sheet
status: draft
parent: EPIC-14
depends_on:
  - STORY-14-1
traces_to:
  - path: docs/specs/table-catalog-query-and-views.md
    anchor: "#apply_table_view"
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - apply_table_view mode=snapshot creates one worksheet of values for the projected, filtered, and sorted rows, including a header row.
  - Limit and offset are honored on the snapshot because it is a copy of the query result, not an AutoFilter.
  - The source sheet's FilterMode, row order, and hidden columns are unchanged.
  - The implementation uses Worksheets.Add and a value write. It does not call Worksheet.Copy.
  - The new sheet name does not collide with an existing sheet. TOOLS.md describes snapshot as the option that leaves a shared source sheet alone.
created: "2026-09-23"
updated: "2026-09-23"
---

# Story-14-3: Snapshot a query onto a new sheet

## Description

When other people have the workbook open, an in-place filter is the wrong visual. Snapshot writes the query result to a new sheet and leaves the source untouched. Limit and offset are allowed here because the sheet is a copy of the rows the agent already selected.

## User story

As an **agent**, I want **a result sheet for a query** so **the user can see the rows without changing the shared table**.

## Technical notes

- COM `Worksheet.Copy(After=...)` was observed to open a new workbook (`Book3`) and leave the original active sheet in place. That API is out of bounds for this story.
- Pasting a `ListObject` range can create another table with a generated name. Snapshot should write values, not duplicate the source table, unless a later story asks for a live table.
- `clear_table_view` for a snapshot deletes that sheet only when the restore token says this tool created it.
