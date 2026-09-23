# Excel MCP Server Tools

This document provides detailed information about all available tools in the Excel MCP server. For **install, transports, env vars, and allowlists**, see the repository **README**. The **MCP client** also receives a short **server `instructions`** string (see `FastMCP` in `src/excel_mcp/server.py`) summarizing `filepath` rules and where to read more.

## `filepath`: disk path, SharePoint URL, and COM matching

Every workbook tool takes **`filepath`** (sometimes shown as `filename` in older docs). The server resolves it through **`get_excel_path`** (`src/excel_mcp/server.py`):

- **Absolute local path** — Normal file identity; `resolve_target` / `realpath` apply when the path allowlist is on (see README).
- **`https://…` cloud workbook locator (v1)** — Allowed under **stdio** when validation passes (`parse_cloud_workbook_locator` in `excel_mcp.path_resolution`). Used for **COM** automation so the string matches Excel **`Workbook.FullName`** (often SharePoint). **Do not** pass `https` when **`EXCEL_FILES_PATH`** is set unless you use only local relative paths under the jail (cloud URLs are rejected there).
- **Finding the right string for an open cloud file:** In Excel, **VBA Immediate** → `? ActiveWorkbook.FullName`. If the result is an `https://` URL, pass that as **`filepath`**, not only the synced folder path on disk—otherwise COM may not match and **`auto`** can try **`openpyxl`** and fail with **permission denied** while Excel has the file open.
- **Discovery (no guesswork):** Call **`excel_list_open_workbooks`** (Windows + COM, Excel running) to get a JSON list of open workbooks with exact **`full_name`** strings. Copy a **`full_name`** into **`get_workbook_metadata`**, **`read_data_from_excel`**, writes, and lifecycle tools—same flow as the VBA one-liner, but in the MCP contract ([ADR 0009](docs/architecture/adr/0009-open-workbook-discovery-tool.md)). Allowlist policy applies when you **use** a path or URL as **`filepath`**; discovery only **reports** what Excel has open.

Optional **`workbook_transport`** (`auto` \| `file` \| `com`) applies to **routed** workbook tools (see table below). **Session / host** tools **`excel_list_open_workbooks`** (ADR 0009), **`excel_open_workbook`**, **`excel_close_workbook`**, and **`evaluate_range`** are **COM-only** and do not use the routing matrix (ADR 0008 / ADR 0009). **`evaluate_range`** additionally rejects explicit **`workbook_transport=file`**. Authentication for M365 is **Excel/Office**, not the MCP server.

## Workbook routing parameters (all tools)

Optional keyword arguments (when the host exposes JSON Schema for tool inputs, they appear there too):

| Parameter | Type | Default | Meaning |
| --------- | ---- | ------- | ------- |
| `workbook_transport` | string, optional | from `EXCEL_MCP_TRANSPORT` env (`auto` if unset) | Workbook execution mode: `auto`, `file`, or `com` (case-insensitive). **Not** the MCP wire transport (stdio/SSE/HTTP). |

**Persistence:** call the dedicated **`save_workbook`** tool when you need changes flushed to disk (ADR 0003 / ADR 0008).

Full operator env and read-vs-disk guidance: see repository **README** (Story 7-5).

## Workbook Operations

### create_workbook

Creates a new Excel workbook (file and/or COM path per `workbook_transport`).

```python
create_workbook(
    filepath: str,
    workbook_transport: str | None = None,
    open_in_excel: bool = False,
) -> str
```

- `filepath`: Path where to create workbook
- `workbook_transport`: same as other tools (default from `EXCEL_MCP_TRANSPORT`)
- `open_in_excel`: if true, after a successful create the server calls **`Workbooks.Open`** in Excel when COM is available, so `auto` routing can match the open host (ADR 0008). If COM is unavailable, a note is appended and the file is still created.
- Returns: Success message (and optional follow-up from open-in-Excel)

### excel_open_workbook

Open an **existing** workbook in the Excel application (`Workbooks.Open`). Use when the file exists on disk or as a allowed `https` locator; subsequent **`workbook_transport=auto`** / **`com`** tools can use COM-first routing for that identity. **Windows + COM only.**

```python
excel_open_workbook(filepath: str) -> str
```

### excel_close_workbook

Close a workbook in the Excel host. Does not delete the file. **`save`**: if true, Excel saves to the workbook path before close.

```python
excel_close_workbook(filepath: str, save: bool = False) -> str
```

### excel_list_open_workbooks

Enumerates **`Application.Workbooks`** in the running Excel instance and returns **JSON** (string payload). Default **`detail=minimal`**: `{"workbooks": [{"full_name", "name", "is_active"}, ...]}`. **`detail=active_context`** adds top-level **`active_workbook`**, **`active_sheet`**, and **`selection`** (same minimal workbook list). **`full_name`** is the exact COM locator (absolute path or **`https://…`** SharePoint-style URL); use it as **`filepath`** on routed tools. Order matches Excel’s workbook collection indexes (deterministic).

**Workflow:** discovery → choose **`full_name`** → **`get_workbook_metadata`** / **`read_data_from_excel`** / writes / **`excel_close_workbook`** as needed. **`get_workbook_metadata`** still requires **`filepath`**; there is no overload that omits it ([ADR 0009](docs/architecture/adr/0009-open-workbook-discovery-tool.md)).

```python
excel_list_open_workbooks(detail: str | None = None) -> str
```

- **`detail`**: ``minimal`` (default) — workbook list only; ``active_context`` — also active workbook, sheet name, and current selection address (e.g. ``B2:D5``).
- **COM-only**, **no `filepath`**; **`workbook_transport`** does not apply.
- **Empty list:** Excel is running but no workbooks are open (normal).
- **Errors:** Same class of messages as other COM tools when **`excel-com-mcp[com]`** is missing or Excel is not running (“No running Excel application found”); invalid **`detail`** returns an explicit error.

### evaluate_range

Force Excel to **recalculate** formulas in a worksheet or cell range via COM **before** reads when the host may hold stale calculated values. **COM-only** side effect on the in-memory workbook; **does not flush to disk**. After recalc, use **`read_data_from_excel`** with **`workbook_transport=auto`** / **`com`** (workbook open in Excel). For on-disk snapshots, call **`save_workbook`** first, then read with **`workbook_transport=file`**.

**Workflow:** discovery → open workbook → optional **`evaluate_range`** → COM read (or **`save_workbook`** then file read).

```python
evaluate_range(
    filepath: str,
    sheet_name: str,
    start_cell: str | None = None,
    end_cell: str | None = None,
    workbook_transport: str | None = None,
) -> str
```

- **`start_cell` / `end_cell`:** omit both to recalc the entire sheet; provide **`start_cell`** only for one cell; both for a rectangular range.
- **`workbook_transport=file`:** rejected with an explicit error (recalc cannot affect openpyxl reads).
- **Errors:** COM unavailable, Excel not running, workbook not open, or invalid cell references — same patterns as other COM session tools.

### save_workbook

Persist the workbook to disk via the routed backend (openpyxl file path, or Excel COM when routed to COM). Call when you need an **on-disk** snapshot (external tools, `workbook_transport=file`, or after COM writes before file-backed reads). COM-first reads (`workbook_transport=auto`/`com`) use the **live Excel grid** and do not require `save_workbook` first (ADR 0008; ADR 0003 disk semantics).

```python
save_workbook(
    filepath: str,
    workbook_transport: str | None = None,
) -> str
```

- `filepath`: Path to the Excel file to save
- `workbook_transport`: same workbook routing parameter as all tools (see table above)
- Returns: Success message (e.g. confirmation including the file path)

### create_worksheet

Creates a new worksheet in an existing workbook.

```python
create_worksheet(filepath: str, sheet_name: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name for the new worksheet
- Returns: Success message

### get_workbook_metadata

Get metadata about workbook including sheets and ranges.

```python
get_workbook_metadata(filepath: str, include_ranges: bool = False) -> str
```

- `filepath`: Path to Excel file
- `include_ranges`: Whether to include range information
- Returns: String representation of workbook metadata

## Data Operations

### write_data_to_excel

Write data to Excel worksheet.

```python
write_data_to_excel(
    filepath: str,
    sheet_name: str,
    data: List[List],
    start_cell: str = "A1"
) -> str
```

- `filepath`: Workbook path or cloud locator (see **filepath** section at top of this file)
- `sheet_name`: Target worksheet name
- `data`: List of rows, each row a list of cell values (see server implementation)
- `start_cell`: Starting cell (default: "A1")
- Returns: Success message

### read_data_from_excel

Read data from Excel worksheet with per-cell validation metadata.

**Discovery workflow ([ADR 0009](docs/architecture/adr/0009-open-workbook-discovery-tool.md)):** `excel_list_open_workbooks` → copy `full_name` → pass as `filepath` below (same as VBA `? ActiveWorkbook.FullName`).

**COM-first reads ([ADR 0008](docs/architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md)):** With `workbook_transport=auto` or `com`, values come from the **live Excel session**, not necessarily the last saved file on disk. Use `workbook_transport=file` only when you need the on-disk snapshot.

```python
read_data_from_excel(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str = None,
    workbook_transport: str | None = None,
    include_routing_metadata: bool = False,
    value_mode: str = "value",
    metadata_mode: str = "full",
) -> str
```

- `filepath`: Workbook path or cloud locator (see **filepath** section at top; SharePoint-style `https://…` URLs per [ADR 0006](docs/architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md); use `excel_list_open_workbooks` → `full_name` per ADR 0009)
- `sheet_name`: Source worksheet name
- `start_cell`: Starting cell (default: "A1")
- `end_cell`: Optional ending cell (auto-expands when omitted)
- `workbook_transport`: same workbook routing parameter as all routed tools (see table above; ADR 0001)
- `include_routing_metadata`: When `true`, wrap successful JSON in the ADR 0010 envelope (see below). Default `false` for backward compatibility.
- `value_mode`: How cell values are returned — `"value"` (default, raw `Value2` / `cell.value`) or `"text"` (display text). Unknown values return an actionable error.
- `metadata_mode`: Per-cell metadata density — `"full"` (default, includes per-cell `validation` objects) or `"compact"` (omits per-cell validation to reduce payload size on large ranges). Root JSON echoes `"metadata_mode"`.
- Returns: JSON string with per-cell metadata. Root JSON includes `"value_mode"` and `"metadata_mode"`.

**`metadata_mode` trade-offs:**

| Mode | Payload | Validation detail | When to use |
|------|---------|-------------------|-------------|
| `"full"` (default) | Larger; every cell may include a `validation` object (`has_validation: false` when absent) | Full per-cell rules (types, lists, prompts) | Small ranges, data-quality checks, form auditing |
| `"compact"` | Smaller; cells are `address`, `value`, `row`, `column` only | None per cell | Large rectangular reads where values matter more than validation |

For very large tabular exports (e.g. >50 rows), prefer **`export_worksheet_table`** instead of `read_data_from_excel` with `metadata_mode=compact` — it returns headers + row arrays without per-cell addressing overhead. Use `get_data_validation_info` when you need worksheet-level validation rules without reading cell values.

**`value_mode` and backends:**

| `value_mode` | COM backend | File backend (openpyxl) |
|--------------|-------------|-------------------------|
| `"value"` (default) | `Range.Value2` | `cell.value` |
| `"text"` | `Range.Text` (Excel display string) | Best-effort formatted string from `number_format`; **weaker fidelity** than COM — locale, conditional formats, and rich display rules are not reproduced. Prefer COM when exact display text matters. |

**Routing metadata envelope (ADR 0010):** When `include_routing_metadata=true`, the tool returns:

```json
{
  "result": { },
  "_meta": {
    "workbook_transport": "auto",
    "workbook_backend": "com",
    "routing_reason": "full_name_match",
    "duration_ms": 12.345
  },
  "warnings": []
}
```

- `_meta.workbook_backend` reflects the **executed** backend (`file` or `com`). Use it to verify COM routed reads after `workbook_transport=auto`.
- Failures still return plain `"Error: …"` strings (no envelope), even when `include_routing_metadata=true`.
- Field names match NFR-3 / `excel-mcp.routing` log vocabulary ([ADR 0010](docs/architecture/adr/0010-mcp-tool-response-envelope.md)).

**File backend formula limits (`.xlsm`):** The openpyxl file backend does **not** evaluate Excel formulas. For macro-enabled workbooks (`.xlsm`), reads may return `null` for formula cells when cached results are absent. Prefer **COM routing** (`workbook_transport=auto` or `com`) when the workbook is open in Excel so Excel evaluates formulas. When `include_routing_metadata=true` and the file backend reads a range containing formulas in an `.xlsm`, the envelope includes:

```json
{
  "code": "file_backend_formula_not_evaluated",
  "message": "The file (openpyxl) backend does not evaluate Excel formulas; cached values may be missing (null). Prefer COM routing (workbook_transport=auto or com) when the workbook is open in Excel."
}
```

**Troubleshooting — sparse or missing COM read values:** On COM transport, bulk `Range.Value2` (or `Range.Text` when `value_mode="text"`) can return a dense matrix of `null` for grouped, filtered, or outline-heavy sheets while individual cells still hold data. The server probes up to 24 blank-looking bulk cells; when at least 3 direct cell reads are non-blank, it falls back to per-cell reads for the whole range. Ranges larger than 64 cells use stratified grid sampling (evenly spaced rows/columns) so late rows are not under-sampled; smaller ranges scan row-major from the top-left. If values still look wrong, confirm the workbook is open in Excel with COM routing (`workbook_transport=com` or `auto` with a matching open workbook).

### export_worksheet_table

Bulk-read a worksheet as a compact table (header row + data rows). Prefer this over repeated `read_data_from_excel` calls when exporting large rectangular regions (e.g. >50 rows).

```python
export_worksheet_table(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str = None,
    max_rows: int = 10000,
    workbook_transport: str | None = None,
) -> str
```

- `filepath`: Path to Excel file, or exact `https://` SharePoint-style URL matching Excel `Workbook.FullName` when using COM
- `sheet_name`: Source worksheet name
- `start_cell`: Starting cell (default `A1`); when `end_cell` is omitted, exports the sheet used range (file: openpyxl bounds; COM: `UsedRange`)
- `end_cell`: Optional ending cell for a explicit rectangular range
- `max_rows`: Cap on **data** rows returned (default `10000`); first row of the range is always treated as headers and is not counted toward this cap
- `workbook_transport`: Workbook execution mode (`auto`, `file`, `com`)
- Returns: JSON string with `sheet_name`, `range`, `headers`, `rows`, `row_count`, `truncated`, `max_rows`

**Response shape:**

```json
{
  "sheet_name": "Sheet1",
  "range": "A1:F100",
  "headers": ["Col1", "Col2"],
  "rows": [["a", 1], ["b", 2]],
  "row_count": 99,
  "truncated": false,
  "max_rows": 10000
}
```

- `headers`: values from the first row of the exported range
- `rows`: remaining rows as a compact matrix (no per-cell metadata)
- `row_count`: total data rows in the range (excluding the header row), before truncation
- `truncated`: `true` when `row_count` exceeds `max_rows` (response `rows` is capped)


## Formatting Operations

### format_range

Apply formatting to a range of cells.

```python
format_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: int = None,
    font_color: str = None,
    bg_color: str = None,
    border_style: str = None,
    border_color: str = None,
    number_format: str = None,
    alignment: str = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: Dict[str, Any] = None,
    conditional_format: Dict[str, Any] = None
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Optional ending cell of range
- Various formatting options (see parameters)
- Returns: Success message

### merge_cells

Merge a range of cells.

```python
merge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Ending cell of range
- Returns: Success message

### unmerge_cells

Unmerge a previously merged range of cells.

```python
unmerge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Ending cell of range
- Returns: Success message

### get_merged_cells

Get merged cells in a worksheet.

```python
get_merged_cells(filepath: str, sheet_name: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- Returns: String representation of merged cells


## Formula Operations

### apply_formula

Apply Excel formula to cell.

```python
apply_formula(filepath: str, sheet_name: str, cell: str, formula: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `cell`: Target cell reference
- `formula`: Excel formula to apply
- Returns: Success message

### validate_formula_syntax

Validate Excel formula syntax without applying it.

```python
validate_formula_syntax(filepath: str, sheet_name: str, cell: str, formula: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `cell`: Target cell reference
- `formula`: Excel formula to validate
- Returns: Validation result message

## Chart Operations

### create_chart

Create chart in worksheet.

```python
create_chart(
    filepath: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = ""
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `data_range`: Range containing chart data
- `chart_type`: Type of chart (line, bar, pie, scatter, area)
- `target_cell`: Cell where to place chart
- `title`: Optional chart title
- `x_axis`: Optional X-axis label
- `y_axis`: Optional Y-axis label
- Returns: Success message

## Pivot Table Operations

### create_pivot_table

Create pivot table in worksheet.

```python
create_pivot_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    target_cell: str,
    rows: List[str],
    values: List[str],
    columns: List[str] = None,
    agg_func: str = "mean"
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `data_range`: Range containing source data
- `target_cell`: Cell where to place pivot table
- `rows`: Fields for row labels
- `values`: Fields for values
- `columns`: Optional fields for column labels
- `agg_func`: Aggregation function (sum, count, average, max, min)
- Returns: Success message

## Table Operations

### list_tables

**Kind:** READ

List every native Excel table (`ListObject`) in a workbook. Returns catalog
metadata only (no cell values). Does not change filters, sort order, or hidden
columns.

```python
list_tables(
    filepath: str,
    detail: str = "schema",
    workbook_transport: Optional[str] = None,
) -> str
```

- `filepath`: Workbook path or cloud locator (see **filepath** section at top of this file)
- `detail`: `schema` (default) includes column names; `minimal` omits them
- `workbook_transport`: Optional transport override (`auto` | `file` | `com`)
- Returns: JSON string `{"tables": [...]}` with `sheet`, `name`, `range`,
  `header_range`, `data_range`, `row_count`, `filter_applied`, and optionally
  `columns`. Empty workbook → `{"tables": []}`.

### query_table

**Kind:** READ

Filter rows from a native Excel table (`ListObject`) by column projection and
AND-combined `where` clauses. Returns row objects plus a `view_spec` Epic-14 can
apply later. Does **not** change filters, sort order, or hidden columns (never
calls AutoFilter, Sort, ShowAllData, or hide/unhide). File mode does not save.

```python
query_table(
    filepath: str,
    table: str,
    columns: Optional[list[str]] = None,
    where: Optional[list[dict]] = None,
    limit: Optional[int] = None,
    offset: Optional[int] = None,
    omit_empty: bool = False,
    search: Optional[str] = None,
    workbook_transport: Optional[str] = None,
    include_routing_metadata: bool = False,
) -> str
```

- `filepath`: Workbook path or cloud locator (see **filepath** section at top of this file)
- `table`: ListObject name (unique in the workbook; resolve by name, not sheet+range)
- `columns`: Column names to return. **Required when the table has more than 32
  columns** (unless `omit_empty` is true, which allows omitting `columns` up to
  **128** columns); at or under 32, omitting `columns` returns all columns. Prefer an
  explicit short list on wide tables.
- `where`: List of `{column, op, value}` clauses combined with AND. Operators:
  `eq`, `neq`, `contains` (case-insensitive substring), `in` (list value; OR
  within one column), `gt`, `gte`, `lt`, `lte`, `is_empty` (no `value`). Empty
  `where` returns all data rows (subject to limit/offset). Header row is never
  a data row. Comparisons coerce to numbers when both sides parse as numbers.
- `limit`: Max rows in this page (default **100**). Always marked not viewable.
- `offset`: Matching rows to skip (default **0**). Always marked not viewable.
- `omit_empty`: When `true`, after filtering/paging drop projected columns that
  are empty (`None` or `""`) for every returned row (`0` and `false` are kept).
  Default `false` preserves prior payloads. With `omit_empty`, omitting `columns`
  is allowed on tables with up to **128** columns.
- `search`: Optional case-insensitive substring; a row matches when any loaded
  string cell contains it. Combined with `where` as AND. Default unset.
- `workbook_transport`: Optional transport override (`auto` | `file` | `com`)
- `include_routing_metadata`: When `true`, wrap the payload in the ADR 0010
  envelope (`result` / `_meta` / `warnings`); default `false` keeps legacy JSON.
- Returns: JSON with `headers`, `rows` (objects keyed by header), `row_count`
  (rows in this page), `truncated` (true when more matches exist beyond
  offset+limit), `view_spec`, and `view_applicability`. No per-cell addresses.
  `view_applicability.limit_viewable` and `offset_viewable` are always `false`.
  `in` on one column is viewable; `is_empty`, and a partial page from limit/offset,
  are not. `view_spec.where` keeps every requested clause, including ones Excel
  cannot show. When `view_applicability.viewable` is false, a later view must
  refuse the spec rather than apply the viewable subset (that would disagree
  with `row_count`).

### map_sheet_layout

**Kind:** READ

Map one sheet into native Excel tables plus non-table occupied islands in the
used range (cells inside ListObjects are excluded from islands). Each region
includes bounds and a header guess (first island row when mostly text). Does not
create ListObjects or change filters, sort, or hidden columns. File mode does
not save.

```python
map_sheet_layout(
    filepath: str,
    sheet_name: str,
    workbook_transport: Optional[str] = None,
) -> str
```

- `filepath`: Workbook path or cloud locator (see **filepath** section at top of this file)
- `sheet_name`: Worksheet to map
- `workbook_transport`: Optional transport override (`auto` | `file` | `com`)
- Returns: JSON `{"tables": [...], "regions": [...]}`. Tables reuse `list_tables`
  fields plus `kind: "table"`. Regions have `kind: "region"`, stable `id`
  (`"{sheet}!{range}"`), `range`, `header_range`, `header_row`, `columns`, and
  `header_guess`. Empty sheet → empty `regions` (tables may still be listed).

### query_region

**Kind:** READ

Filter rows from a non-table rectangular region using the same column / where /
limit / offset contract as `query_table`. Provide exactly one of `region_id`
(from `map_sheet_layout`) or an explicit A1 `range` (first row is the header).
Does **not** create a ListObject and does **not** change filters, sort, or hidden
columns. File mode does not save.

```python
query_region(
    filepath: str,
    sheet_name: str,
    region_id: Optional[str] = None,
    range: Optional[str] = None,
    columns: Optional[list[str]] = None,
    where: Optional[list[dict]] = None,
    limit: Optional[int] = None,
    offset: Optional[int] = None,
    omit_empty: bool = False,
    search: Optional[str] = None,
    workbook_transport: Optional[str] = None,
    include_routing_metadata: bool = False,
) -> str
```

- `filepath`: Workbook path or cloud locator (see **filepath** section at top of this file)
- `sheet_name`: Worksheet name
- `region_id`: Stable id such as `"Sheet1!B4:E20"`. Requires a successful header
  guess; otherwise pass an explicit `range` that includes a header row.
- `range`: Explicit A1 range; first row is treated as the header (even when it is
  not worksheet row 1)
- `columns` / `where` / `limit` / `offset` / `omit_empty` / `search`: Same
  semantics as `query_table` (default limit **100**, width threshold **32**,
  `omit_empty` up to **128**)
- `workbook_transport`: Optional transport override (`auto` | `file` | `com`)
- `include_routing_metadata`: When `true`, wrap the payload in the ADR 0010
  envelope; default `false` keeps legacy JSON
- Returns: Same shape as `query_table`. `view_spec.target` is
  `{"kind": "region", "sheet": "...", "range": "A1:D10"}` (no ListObject name).

### apply_table_view

**Kind:** WRITE · **COM only**

Show a `view_spec` (from `query_table` or `query_region`) in Excel.

- **`in_place`**: mutates the **shared desktop view** on the source sheet (filters,
  sort, and hidden columns that co-authors also see). This is **not** a personal
  or named sheet view.
- **`snapshot`**: leaves the shared source sheet alone. Creates one new worksheet
  of **values only** (header + projected, filtered, sorted rows) via
  `Worksheets.Add` and a cell-value write. Does **not** call `Worksheet.Copy` and
  does **not** create a ListObject on the new sheet. Use snapshot when others
  have the workbook open and an in-place filter would be the wrong visual.

Supports `mode="in_place"` on a ListObject (`target.kind == "table"`) and on a
plain region (`target.kind == "region"`), and `mode="snapshot"` for both targets.

```python
apply_table_view(
    filepath: str,
    view_spec: dict,
    mode: str = "in_place",
    view_applicability: Optional[dict] = None,
    workbook_transport: Optional[str] = None,
    include_routing_metadata: bool = False,
) -> str
```

- `filepath`: Workbook path or cloud locator (must be open in Excel for COM)
- `view_spec`: From `query_table` / `query_region`, optionally with `sort` set.
  Table shape: `{"target": {"kind": "table", "name": "..."}, "columns": [...],
  "where": [...], "sort": null | {"by": [{"column": "Qty", "order": "asc"|"desc"}]} }`.
  Region shape: `{"target": {"kind": "region", "sheet": "...", "range": "A1:D10"},
  ...}` (no ListObject name). Single-key sugar: `{"column": "Qty", "order": "asc"}`
  is also accepted for `sort`. Field indexes for AutoFilter are relative to the
  **table or region header**, not column A.
- ListObject in-place path: `ListObject` AutoFilter + optional `ListObject.Sort`.
- Plain-range in-place path: `Range.AutoFilter` on that range (does **not** call
  `ListObjects.Add`). Optional sort uses `Worksheet.Sort` with a header row.
  While applied, `Worksheet.AutoFilterMode` and `FilterMode` are true.
- `mode`:
  - `in_place` (default): shared desktop filter/sort/focus on the source sheet.
  - `snapshot`: new sheet of query result values; source FilterMode, row order,
    and hidden columns stay unchanged. Snapshot sheet names are deterministic and
    collision-safe: `mcp_view`, then `mcp_view_2`, `mcp_view_3`, … (case-insensitive
    match against existing sheet names). Honors `limit` / `offset` on the
    `view_spec` when present (copy of the selected page). When both are absent,
    writes the **full** filtered/sorted set (does **not** apply the `query_table`
    default page size of 100).
- `view_applicability`: Optional payload from `query_table` / `query_region`.
  For **in_place**, when `viewable` is `false` (limit/offset truncation,
  `is_empty`, etc.), the call is refused and the sheet is not changed. Prefer
  refusing the whole call over a partial view. For **snapshot**,
  limit/offset-related `viewable: false` does **not** block the call (the sheet
  is a copy of that page). Cross-column OR and non-viewable operators
  (`is_empty`) are still refused in both modes.
- Refused (sheet unchanged):
  - **in_place**: `limit` / `offset` on the spec, cross-column OR, non-viewable
    operators, `view_applicability.viewable is false`, file transport.
  - **snapshot**: cross-column OR, non-viewable operators, file transport.
  In-place sort runs only when a previous sort was captured and can be put
  back. If there is no prior sort, or it cannot be captured, the **sort
  portion** is skipped (filters and column focus may still apply) and a warning
  tells the caller to use snapshot mode. Clearing sort fields does not undo a
  reorder, so the tool does not apply a sort it cannot restore.
  A snapshot of a paged query must carry that `limit` and `offset` on
  `view_spec`. If `view_applicability` says the query was limited or offset and
  those fields are missing, snapshot is refused so the sheet cannot show a
  different row set.
- Column focus (in_place only): hides worksheet columns for fields not listed in
  `view_spec.columns` that were previously visible. Hiding a column hides the
  whole sheet column.
- `workbook_transport`: File backend returns a clear COM-required error and does
  not mutate the xlsx.
- Returns: JSON with `restore_token`. In-place tokens carry kind, prior filters,
  prior sort when captured, columns this call hid (ranges also prior
  `AutoFilterMode` and the A1 range). Snapshot tokens are
  `{"v": 1, "kind": "snapshot", "sheet": "<name>", "created_by_tool": true}`.

### clear_table_view

**Kind:** WRITE · **COM only**

Restore the sheet using a `restore_token` from `apply_table_view`.

```python
clear_table_view(
    filepath: str,
    restore_token: dict,
    workbook_transport: Optional[str] = None,
    include_routing_metadata: bool = False,
) -> str
```

- **ListObject tokens** (`kind: "listobject"`): clear with
  `ListObject.AutoFilter.ShowAllData()` (works when `FilterMode` is true even if
  `Worksheet.AutoFilterMode` is false). Does **not** clear via
  `Worksheet.AutoFilterMode`.
- **Plain-range tokens** (`kind: "range"`): clear via worksheet AutoFilter /
  restoring prior `AutoFilterMode` (whether filter arrows were already on). Do
  **not** clear a plain-range filter by only calling
  `ListObject.AutoFilter.ShowAllData()`.
- **Snapshot tokens** (`kind: "snapshot"`): deletes **only** the sheet named in
  the token, and only when `created_by_tool` is `true` (this tool created that
  sheet). Does not delete arbitrary sheets.
- Re-applies prior filter criteria from the token (in-place kinds), restores prior
  sort when this apply had changed sort, and unhides **only** columns listed in
  `restore_token.columns_hidden` (columns the user already had hidden stay
  hidden).
- File transport returns a clear COM-required error without mutating the xlsx.

### create_table

Creates a native Excel table from a specified range of data.

```python
create_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    table_name: str = None,
    table_style: str = "TableStyleMedium9"
) -> str
```

- `filepath`: Path to the Excel file.
- `sheet_name`: Name of the worksheet.
- `data_range`: The cell range for the table (e.g., "A1:D5").
- `table_name`: Optional unique name for the table.
- `table_style`: Optional visual style for the table.
- Returns: Success message.

## Worksheet Operations

### copy_worksheet

Copy worksheet within workbook.

```python
copy_worksheet(filepath: str, source_sheet: str, target_sheet: str) -> str
```

- `filepath`: Path to Excel file
- `source_sheet`: Name of sheet to copy
- `target_sheet`: Name for new sheet
- Returns: Success message

### delete_worksheet

Delete worksheet from workbook.

```python
delete_worksheet(filepath: str, sheet_name: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name of sheet to delete
- Returns: Success message

### rename_worksheet

Rename worksheet in workbook.

```python
rename_worksheet(filepath: str, old_name: str, new_name: str) -> str
```

- `filepath`: Path to Excel file
- `old_name`: Current sheet name
- `new_name`: New sheet name
- Returns: Success message

## Range Operations

### copy_range

Copy a range of cells to another location.

```python
copy_range(
    filepath: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str = None
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Source worksheet name
- `source_start`: Starting cell of source range
- `source_end`: Ending cell of source range
- `target_start`: Starting cell for paste
- `target_sheet`: Optional target worksheet name
- Returns: Success message

### delete_range

Delete a range of cells and shift remaining cells.

```python
delete_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up"
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Ending cell of range
- `shift_direction`: Direction to shift cells ("up" or "left")
- Returns: Success message

### validate_excel_range

Validate if a range exists and is properly formatted.

```python
validate_excel_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_cell`: Starting cell of range
- `end_cell`: Optional ending cell of range
- Returns: Validation result message

### get_data_validation_info

Get data validation rules and metadata for a worksheet.

```python
get_data_validation_info(filepath: str, sheet_name: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- Returns: JSON string containing all data validation rules with metadata including:
  - Validation type (list, whole, decimal, date, time, textLength)
  - Operator (between, notBetween, equal, greaterThan, lessThan, etc.)
  - Allowed values for list validations (resolved from ranges)
  - Formula constraints for numeric/date validations
  - Cell ranges where validation applies
  - Prompt and error messages

**Note**: The `read_data_from_excel` tool automatically includes validation metadata for individual cells when available.

## Row and Column Operations

### insert_rows

Insert one or more rows starting at the specified row.

```python
insert_rows(
    filepath: str,
    sheet_name: str,
    start_row: int,
    count: int = 1
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_row`: Row number where to start inserting (1-based)
- `count`: Number of rows to insert (default: 1)
- Returns: Success message

### insert_columns

Insert one or more columns starting at the specified column.

```python
insert_columns(
    filepath: str,
    sheet_name: str,
    start_col: int,
    count: int = 1
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_col`: Column number where to start inserting (1-based)
- `count`: Number of columns to insert (default: 1)
- Returns: Success message

### delete_sheet_rows

Delete one or more rows starting at the specified row.

```python
delete_sheet_rows(
    filepath: str,
    sheet_name: str,
    start_row: int,
    count: int = 1
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_row`: Row number where to start deleting (1-based)
- `count`: Number of rows to delete (default: 1)
- Returns: Success message

### delete_sheet_columns

Delete one or more columns starting at the specified column.

```python
delete_sheet_columns(
    filepath: str,
    sheet_name: str,
    start_col: int,
    count: int = 1
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `start_col`: Column number where to start deleting (1-based)
- `count`: Number of columns to delete (default: 1)
- Returns: Success message
