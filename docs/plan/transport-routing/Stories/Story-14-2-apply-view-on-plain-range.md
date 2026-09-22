---
kind: story
id: STORY-14-2
title: Apply a view on a plain range
status: draft
parent: EPIC-14
depends_on:
  - STORY-14-1
  - STORY-13-3
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
  - apply_table_view in_place on a region that is not a ListObject uses Range.AutoFilter on that range and does not call ListObjects.Add.
  - While the filter is applied, Worksheet.AutoFilterMode and FilterMode are true. clear_table_view restores the previous AutoFilterMode.
  - The source sheet gains no table. Hidden columns outside the columns this call hid stay hidden.
  - TOOLS.md distinguishes table filters (ShowAllData on the ListObject) from plain-range filters (AutoFilterMode).
created: "2026-09-23"
updated: "2026-09-23"
---

# Story-14-2: Apply a view on a plain range

## Description

`stackfleet concept` has no table. A view on that kind of block uses `Range.AutoFilter` and must not promote the range into a `ListObject`. Restore has to remember whether AutoFilter arrows were already on.

## User story

As an **agent**, I want **to filter a plain range in Excel** so **sheets that are not tables can still be shown to the user**.

## Technical notes

- Reuse the viewable-clause rules from Story 14-1. Sort on a plain range uses `Worksheet.Sort` with a header row, still captured in the restore token.
- Do not reuse `Worksheet.Copy` to stage the experiment. Tests that need a scratch sheet add one and delete it.
