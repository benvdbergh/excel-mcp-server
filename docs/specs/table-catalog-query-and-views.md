# Table catalog, row query, and Excel views

**Status:** Draft for implementation planning (2026-09-23).

**Audience:** Agents using `excel-com-mcp`, and implementers extending `WorkbookOperation`.

This spec turns a live Excel workbook into two cooperating planes:

1. **Data plane.** Discover tables and query rows. Reads do not change filters, sort order, or hidden columns.
2. **Presentation plane.** Apply the same specification as something the user can see in Excel, then clear it. COM only. Separate tools from the data plane.

Cell tools (`read_data_from_excel`, `export_worksheet_table`) stay for addresses and bulk dumps. This spec does not replace them.

## Evidence

Confirmed on Excel 16 against the open workbook `AP-Component Register.xlsx` (AutoSave on):

| Fact | Consequence for the product |
|------|-----------------------------|
| `software components` is ListObject `ap_sw_components` at `B2:DQ59` (120 columns). `stackfleet concept` is a plain range with no ListObject. | Catalog native tables and detected regions separately. |
| On a ListObject, an active filter sets `ListObject.AutoFilter.FilterMode` and `Worksheet.FilterMode`. `Worksheet.AutoFilterMode` stays false. | Clear table filters with `ListObject.AutoFilter.ShowAllData()`. Do not use `AutoFilterMode` as the signal. |
| Equality, `*contains*`, numeric comparison, AND across fields, and OR inside one field all work. A two-value list is stored as `xlOr` with `Criteria1` and `Criteria2`. | Those clauses are viewable. |
| `xlTop10Items` is rank-by-value (ties expand the row count). It is not "first N rows". | Limit and offset stay on the data plane. The view compiler rejects them. |
| `ListObject.Sort` reorders rows. A sort on a copied sheet left the source order unchanged. | In-place sort is a visible mutation and must be restorable. |
| Hiding a column hides the entire worksheet column, including cells outside the table. It does not hide that column on another sheet. The sample sheet already had 19 hidden columns. | Column focus is per worksheet. Restore only columns this tool hid. |
| `Worksheet.Copy` opened a new workbook instead of inserting a sheet. Pasting the range in-workbook duplicated the table under a new unique name. | Snapshot sheets are created with `Worksheets.Add` and a value paste. Do not call `Worksheet.Copy`. |
| `Window.SheetViews` is the worksheet display (gridlines, headings), one per sheet, with no `Add`. `CustomViews` was empty. Named personal sheet views are an Excel JavaScript API (`namedSheetViews`), not this COM surface. | In-place filter and sort are the shared desktop view. Do not promise a private "see just mine" view. |

## Tools

| Tool | Kind | Backends | Mutates the grid the user sees |
|------|------|----------|--------------------------------|
| `list_tables` | READ | file and COM | No |
| `query_table` | READ | file and COM | No |
| `map_sheet_layout` | READ | file and COM | No |
| `query_region` | READ | file and COM | No |
| `apply_table_view` | WRITE | COM only | Yes |
| `clear_table_view` | WRITE | COM only | Yes, back to the captured prior state |

`query_table` and `query_region` return rows plus a `view_spec`. They never call AutoFilter, Sort, or hide columns.

### `list_tables`

Workbook scope. For each `ListObject`: sheet, name, range, header range, data range, column names, row count, whether a filter is currently applied. `detail=minimal` omits column names. `detail=schema` includes them. No cell values.

### `query_table`

Arguments: `table` (ListObject name), `columns`, `where`, `limit`, `offset`.

`where` is a list of clauses combined with AND. Clause operators: `eq`, `neq`, `contains`, `in`, `gt`, `gte`, `lt`, `lte`, `is_empty`. OR across different columns is not a viewable clause. OR inside one column is the `in` operator (or an explicit `or` on a single column).

Response: `headers`, rows as objects keyed by header, `row_count`, `truncated`, `view_spec`, and `view_applicability` (which clauses Excel can show). Use the ADR 0010 envelope when `include_routing_metadata` is true. No per-cell addresses.

COM reads only the selected `ListColumns`. File mode uses the worksheet table collection, then the same filter function, so both backends return the same rows.

### `map_sheet_layout` and `query_region`

`map_sheet_layout` returns native tables first, then non-empty islands in the used range that sit outside those tables, each with bounds and a header guess.

`query_region` uses the same `columns` / `where` / `limit` contract against a region id or an explicit range. It does not convert the range into a ListObject.

### `apply_table_view`

COM only. Inputs: a `view_spec` and `mode`.

- `in_place` on a ListObject: `AutoFilter` and `ListObject.Sort` for viewable clauses; optional column focus hides columns that were not hidden already.
- `in_place` on a plain region: `Range.AutoFilter` on that range. Do not call `ListObjects.Add`.
- `snapshot`: add a worksheet and paste the projected, filtered, sorted values. The source filter, order, and hidden columns stay as they were.

The tool refuses the call when the spec contains a clause Excel cannot show (limit, offset, cross-column OR). It does not apply a partial view that would disagree with the row count the agent already has.

The response includes a `restore_token` capturing the previous filter criteria, sort fields, and the columns hidden by this call.

### `clear_table_view`

COM only. Applies `restore_token`: `ShowAllData` on the table (or `ShowAllData` plus restoring prior `AutoFilterMode` on a plain range), restores the prior sort, and unhides only columns this tool hid.

## Non-goals

- A SQL parser.
- Using AutoFilter as the implementation of `query_table`.
- Aggregations and pivots (`create_pivot_table` stays file-forced under ADR 0004).
- Personal or named sheet views.
- Unhiding columns the user had already hidden.
