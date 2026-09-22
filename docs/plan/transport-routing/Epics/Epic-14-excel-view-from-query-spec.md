---
kind: epic
id: EPIC-14
title: Excel view from a query spec
status: draft
depends_on:
  - EPIC-13
traces_to:
  - path: docs/specs/table-catalog-query-and-views.md
  - path: docs/architecture/adr/0002-com-automation-stack.md
  - path: docs/architecture/adr/0004-chart-pivot-com-parity-scope.md
slice: vertical
acceptance_criteria:
  - apply_table_view can show a viewable query on a ListObject or a plain range, and clear_table_view restores the previous filter, sort, and only the columns this tool hid.
  - The tool refuses limit, offset, and cross-column OR instead of showing a different row set than the agent queried.
  - Snapshot mode adds a sheet of values and leaves the source sheet's filter, order, and hidden columns unchanged.
  - In-place mode is documented as the shared desktop view. The product does not claim a personal sheet view.
created: "2026-09-23"
updated: "2026-09-23"
---

# Epic-14: Excel view from a query spec

## Description

After [Epic-13](Epic-13-table-catalog-and-row-query.md) returns a `view_spec`, a separate COM write can show that spec in Excel. Query tools stay read-only. `apply_table_view` and `clear_table_view` are the only tools that change filters, sort, or column visibility.

In-place changes are visible to everyone in the desktop workbook. Excel COM does not expose named personal sheet views (`Window.SheetViews` is display chrome only). Snapshot mode is the option when the file is shared.

## Rough effort

**Total (epic):** approximately **8–13 developer-days**.

| Story | Rough sizing |
|-------|----------------|
| [14-1](../Stories/Story-14-1-apply-and-clear-listobject-view.md) | ~4–6 days |
| [14-2](../Stories/Story-14-2-apply-view-on-plain-range.md) | ~2–3 days |
| [14-3](../Stories/Story-14-3-snapshot-sheet-view.md) | ~2–4 days |

## Risks

| Risk | Mitigation |
|------|------------|
| Clearing via `AutoFilterMode` leaves a table filtered | Use `ListObject.AutoFilter.ShowAllData()`. Tests cover a table whose `AutoFilterMode` is false while `FilterMode` is true. |
| Restore unhides columns the user had hidden | Persist the pre-existing hidden set and unhide only columns this call hid. |
| `Worksheet.Copy` opens another workbook | Snapshot uses `Worksheets.Add` and a value paste. |
| Sort cannot be undone | Capture sort fields in `restore_token` before applying. If the prior sort cannot be reapplied, refuse in-place sort and offer snapshot. |

## User stories

- [Story-14-1](../Stories/Story-14-1-apply-and-clear-listobject-view.md) — in-place table filter, sort, column focus, and restore.
- [Story-14-2](../Stories/Story-14-2-apply-view-on-plain-range.md) — the same view on a range that is not a table.
- [Story-14-3](../Stories/Story-14-3-snapshot-sheet-view.md) — a result sheet that does not touch the source.

## Recommended sequencing

1. **14-1** after **13-2**, so the view spec already exists.
2. **14-2** after **13-3** and **14-1**.
3. **14-3** after **14-1**. It can proceed in parallel with **14-2**.
