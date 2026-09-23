"""COM ListObject catalog, row query, create-table, and layout helpers."""

from __future__ import annotations

import json
import uuid
from typing import Any, Dict, List, Optional, Tuple

from excel_mcp.exceptions import DataError
from excel_mcp.query import (
    DEFAULT_LIST_TABLES_DETAIL,
    DETAIL_SCHEMA,
    build_sheet_layout_payload,
    build_table_catalog_entry,
    build_table_catalog_payload,
    layout_regions_from_grid,
    normalize_excel_a1_range,
    normalize_where_clauses,
    parse_region_id,
    query_region_rows,
    query_table_rows,
    resolve_query_columns,
    resolve_region_query_source,
    validate_list_tables_detail,
)
from excel_mcp.com.read import _normalize_excel_matrix
from excel_mcp.com.session import _coerce_com_count, get_open_workbook_com

# Excel COM constants (avoid importing win32com at module load)
_XL_SRC_RANGE = 1
_XL_LIST_HAS_HEADERS_GUESS = 0

def _com_list_object_filter_applied(lo: Any) -> bool:
    """Read ``ListObject.AutoFilter.FilterMode`` only (never Worksheet.AutoFilterMode)."""
    try:
        af = lo.AutoFilter
    except Exception:
        return False
    if af is None:
        return False
    try:
        return bool(af.FilterMode)
    except Exception:
        return False

def _com_list_object_catalog_entry(
    lo: Any, sheet_name: str, detail_mode: str
) -> Optional[dict[str, Any]]:
    """Build one catalog entry from a COM ListObject (read-only property access)."""
    try:
        name = str(lo.Name)
        full_range = str(lo.Range.Address)
    except Exception:
        return None

    try:
        header_range = str(lo.HeaderRowRange.Address)
    except Exception:
        header_range = full_range

    data_range: Optional[str] = None
    row_count = 0
    try:
        body = lo.DataBodyRange
    except Exception:
        body = None
    if body is not None:
        try:
            data_range = str(body.Address)
            row_count = _coerce_com_count(getattr(body.Rows, "Count", 0))
        except Exception:
            try:
                row_count = _coerce_com_count(getattr(lo.ListRows, "Count", 0))
            except Exception:
                row_count = 0

    columns: Optional[list[str]] = None
    if detail_mode == DETAIL_SCHEMA:
        columns = []
        try:
            n_cols = _coerce_com_count(getattr(lo.ListColumns, "Count", 0))
        except Exception:
            n_cols = 0
        for ci in range(1, n_cols + 1):
            try:
                columns.append(str(lo.ListColumns.Item(ci).Name))
            except Exception:
                columns.append("")

    return build_table_catalog_entry(
        sheet=sheet_name,
        name=name,
        range_=full_range,
        header_range=header_range,
        data_range=data_range,
        row_count=row_count,
        filter_applied=_com_list_object_filter_applied(lo),
        columns=columns,
    )

def _com_list_object_column_names(lo: Any) -> list[str]:
    columns: list[str] = []
    try:
        n_cols = _coerce_com_count(getattr(lo.ListColumns, "Count", 0))
    except Exception:
        n_cols = 0
    for ci in range(1, n_cols + 1):
        try:
            columns.append(str(lo.ListColumns.Item(ci).Name))
        except Exception:
            columns.append("")
    return columns

def _com_find_list_object(wb_com: Any, table_name: str) -> Tuple[Optional[Any], Optional[str]]:
    """Return ``(list_object, error_string)``."""
    target = str(table_name).strip()
    try:
        n_sheets = _coerce_com_count(getattr(wb_com.Worksheets, "Count", 0))
    except Exception:
        n_sheets = 0
    for si in range(1, n_sheets + 1):
        try:
            ws = wb_com.Worksheets.Item(si)
        except Exception:
            continue
        try:
            n_lo = _coerce_com_count(getattr(ws.ListObjects, "Count", 0))
        except Exception:
            n_lo = 0
        for ti in range(1, n_lo + 1):
            try:
                lo = ws.ListObjects.Item(ti)
                name = str(lo.Name)
            except Exception:
                continue
            if name == target:
                return lo, None
    return None, f"Error: Table '{table_name}' not found"

def _com_list_object_field_map(lo: Any) -> dict[str, int]:
    """Header name → 1-based AutoFilter field index (relative to table, not column A)."""
    names = _com_list_object_column_names(lo)
    return {name: i + 1 for i, name in enumerate(names) if name}

def _com_list_column_sheet_index(lo: Any, field_index: int) -> int:
    """Absolute worksheet column index for a 1-based table field."""
    col = lo.ListColumns.Item(field_index)
    return int(col.Range.Column)

def _com_list_column_data_values(lo: Any, column_name: str) -> list[Any]:
    """Read one ListColumn ``DataBodyRange`` only (never the full table grid)."""
    try:
        col = lo.ListColumns.Item(column_name)
    except Exception as e:
        raise DataError(f"Unknown column: {column_name!r}") from e
    try:
        body = col.DataBodyRange
    except Exception:
        body = None
    if body is None:
        return []
    try:
        matrix = _normalize_excel_matrix(body.Value2)
    except Exception:
        return []
    return [row[0] if row else None for row in matrix]

def _com_list_object_row_dicts(
    lo: Any, needed_columns: Sequence[str]
) -> list[dict[str, Any]]:
    """Read ListObject data rows as dicts (DataBodyRange per column; no AutoFilter)."""
    col_arrays: dict[str, list[Any]] = {}
    for col_name in needed_columns:
        col_arrays[col_name] = _com_list_column_data_values(lo, col_name)
    row_count = max((len(v) for v in col_arrays.values()), default=0)
    rows: list[dict[str, Any]] = []
    for ri in range(row_count):
        row_map: dict[str, Any] = {}
        for col_name in needed_columns:
            values = col_arrays[col_name]
            row_map[col_name] = values[ri] if ri < len(values) else None
        rows.append(row_map)
    return rows

def _com_query_table_rows_from_list_object(
    lo: Any,
    table_name: str,
    columns: Optional[List[str]],
    where: Optional[List[Dict[str, Any]]],
    limit: Optional[int],
    offset: Optional[int],
    omit_empty: bool = False,
    search: Optional[str] = None,
) -> dict[str, Any]:
    table_columns = _com_list_object_column_names(lo)
    headers = resolve_query_columns(
        table_columns, columns, omit_empty=omit_empty
    )
    clauses = normalize_where_clauses(where, known_columns=table_columns)
    needed = list(dict.fromkeys(list(headers) + [c["column"] for c in clauses]))
    rows = _com_list_object_row_dicts(lo, needed)
    return query_table_rows(
        table_name=table_name,
        table_columns=table_columns,
        rows=rows,
        columns=headers,
        where=clauses,
        limit=limit,
        offset=offset,
        omit_empty=omit_empty,
        search=search,
    )

def _com_workbook_sheet_names(wb_com: Any) -> list[str]:
    """Worksheet names currently in the workbook (1-based COM collection)."""
    names: list[str] = []
    try:
        n = int(wb_com.Worksheets.Count)
    except Exception:
        return names
    for i in range(1, n + 1):
        try:
            names.append(str(wb_com.Worksheets.Item(i).Name))
        except Exception:
            continue
    return names

def _com_sheet_table_entries(
    ws: Any, sheet_name: str, detail: str = DETAIL_SCHEMA
) -> list[dict[str, Any]]:
    """ListObject catalog rows for one sheet, each with ``kind: table``."""
    mode = validate_list_tables_detail(detail)
    tables: list[dict[str, Any]] = []
    try:
        n_lo = _coerce_com_count(getattr(ws.ListObjects, "Count", 0))
    except Exception:
        n_lo = 0
    for ti in range(1, n_lo + 1):
        try:
            lo = ws.ListObjects.Item(ti)
        except Exception:
            continue
        entry = _com_list_object_catalog_entry(lo, sheet_name, mode)
        if entry is not None:
            entry = dict(entry)
            entry["kind"] = "table"
            tables.append(entry)
    return tables

def _com_read_a1_matrix(ws: Any, range_a1: str) -> list[list[Any]]:
    """Read one A1 range via ``Range.Value2`` (no ListObjects.Add / AutoFilter)."""
    rng = ws.Range(normalize_excel_a1_range(range_a1))
    return _normalize_excel_matrix(rng.Value2)

def _com_map_sheet_layout(ws: Any, sheet_name: str) -> dict[str, Any]:
    tables = _com_sheet_table_entries(ws, sheet_name)
    excluded = [str(t["range"]) for t in tables]
    try:
        used = ws.UsedRange
        origin_row = int(used.Row)
        origin_col = int(used.Column)
        matrix = _normalize_excel_matrix(used.Value2)
    except Exception:
        return build_sheet_layout_payload(tables, [])
    regions = layout_regions_from_grid(
        sheet=sheet_name,
        values=matrix,
        origin_row=origin_row,
        origin_col=origin_col,
        excluded_ranges=excluded,
    )
    return build_sheet_layout_payload(tables, regions)

def list_tables_com(filepath: str, detail: str) -> str:
    try:
        mode = validate_list_tables_detail(detail)
    except DataError as e:
        return f"Error: {str(e)}"

    wb_com, err = get_open_workbook_com(filepath)
    if err:
        return err

    tables: list[dict[str, Any]] = []
    try:
        n_sheets = _coerce_com_count(getattr(wb_com.Worksheets, "Count", 0))
    except Exception:
        n_sheets = 0

    for si in range(1, n_sheets + 1):
        try:
            ws = wb_com.Worksheets.Item(si)
            sheet_name = str(ws.Name)
        except Exception:
            continue
        try:
            n_lo = _coerce_com_count(getattr(ws.ListObjects, "Count", 0))
        except Exception:
            n_lo = 0
        for ti in range(1, n_lo + 1):
            try:
                lo = ws.ListObjects.Item(ti)
            except Exception:
                continue
            entry = _com_list_object_catalog_entry(lo, sheet_name, mode)
            if entry is not None:
                tables.append(entry)

    return json.dumps(build_table_catalog_payload(tables), indent=2, default=str)

def query_table_com(
    filepath: str,
    table: str,
    columns: Optional[List[str]],
    where: Optional[List[Dict[str, Any]]],
    limit: Optional[int],
    offset: Optional[int],
    omit_empty: bool = False,
    search: Optional[str] = None,
) -> str:
    wb_com, err = get_open_workbook_com(filepath)
    if err:
        return err
    lo, find_err = _com_find_list_object(wb_com, table)
    if find_err:
        return find_err
    assert lo is not None
    try:
        payload = _com_query_table_rows_from_list_object(
            lo,
            str(lo.Name),
            columns,
            where,
            limit,
            offset,
            omit_empty=omit_empty,
            search=search,
        )
        return json.dumps(payload, indent=2, default=str)
    except DataError as e:
        return f"Error: {str(e)}"

def map_sheet_layout_com(filepath: str, sheet_name: str) -> str:
    wb_com, err = get_open_workbook_com(filepath)
    if err:
        return err
    try:
        ws = wb_com.Worksheets(sheet_name)
    except Exception:
        return f"Error: Sheet '{sheet_name}' not found"
    try:
        payload = _com_map_sheet_layout(ws, sheet_name)
        return json.dumps(payload, indent=2, default=str)
    except DataError as e:
        return f"Error: {str(e)}"

def query_region_com(
    filepath: str,
    sheet_name: str,
    region_id: Optional[str],
    range: Optional[str],
    columns: Optional[List[str]],
    where: Optional[List[Dict[str, Any]]],
    limit: Optional[int],
    offset: Optional[int],
    omit_empty: bool = False,
    search: Optional[str] = None,
) -> str:
    has_id = region_id is not None and str(region_id).strip() != ""
    has_range = range is not None and str(range).strip() != ""
    if has_id == has_range:
        return "Error: Provide exactly one of region_id or range"

    wb_com, err = get_open_workbook_com(filepath)
    if err:
        return err
    try:
        ws = wb_com.Worksheets(sheet_name)
    except Exception:
        return f"Error: Sheet '{sheet_name}' not found"
    try:
        if has_id:
            range_a1 = parse_region_id(str(region_id), sheet_name=sheet_name)
            layout = _com_map_sheet_layout(ws, sheet_name)
            known_ids = {str(r["id"]) for r in layout["regions"]}
            if str(region_id).strip() not in known_ids:
                return f"Error: Unknown region_id: {region_id!r}"
            require_guess = True
        else:
            range_a1 = normalize_excel_a1_range(str(range))
            require_guess = False
        matrix = _com_read_a1_matrix(ws, range_a1)
        region_columns, rows, rng = resolve_region_query_source(
            matrix=matrix,
            range_a1=range_a1,
            require_header_guess=require_guess,
        )
        payload = query_region_rows(
            sheet=sheet_name,
            range_a1=rng,
            region_columns=region_columns,
            rows=rows,
            columns=columns,
            where=where,
            limit=limit,
            offset=offset,
            omit_empty=omit_empty,
            search=search,
        )
        return json.dumps(payload, indent=2, default=str)
    except DataError as e:
        return f"Error: {str(e)}"

def create_excel_table_com(
    filepath: str,
    sheet_name: str,
    data_range: str,
    table_name: Optional[str],
    table_style: str,
) -> str:
    wb_com, err = get_open_workbook_com(filepath, for_write=True)
    if err:
        return err
    try:
        ws = wb_com.Worksheets(sheet_name)
    except Exception:
        return f"Error: Sheet '{sheet_name}' not found"

    tname = table_name or f"Table_{uuid.uuid4().hex[:8]}"

    try:
        src = ws.Range(data_range)
        lo = ws.ListObjects.Add(
            _XL_SRC_RANGE,
            src,
            None,
            _XL_LIST_HAS_HEADERS_GUESS,
        )
        lo.Name = tname
        lo.TableStyle = table_style
    except Exception as exc:
        return f"Error: {exc}"

    return f"Successfully created table '{tname}' in sheet '{sheet_name}'."

