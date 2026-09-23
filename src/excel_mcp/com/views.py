"""COM apply/clear table/region/snapshot views and AutoFilter/sort helpers."""

from __future__ import annotations

import json
from typing import Any, Mapping, Optional, Sequence, Tuple

from openpyxl.utils.cell import range_boundaries

from excel_mcp.exceptions import DataError
from excel_mcp.query import (
    RESTORE_KIND_LISTOBJECT,
    RESTORE_KIND_RANGE,
    RESTORE_KIND_SNAPSHOT,
    VIEW_MODE_SNAPSHOT,
    XL_SORT_ASCENDING,
    XL_SORT_DESCENDING,
    allocate_unique_sheet_name,
    build_restore_token,
    build_snapshot_value_matrix,
    compile_autofilter_field_steps,
    filter_table_rows,
    normalize_excel_a1_range,
    parse_restore_token,
    resolve_query_columns,
    sort_table_rows,
    validate_view_spec_for_apply,
)
from excel_mcp.com.listobject import (
    _com_find_list_object,
    _com_list_column_sheet_index,
    _com_list_object_column_names,
    _com_list_object_field_map,
    _com_list_object_row_dicts,
    _com_read_a1_matrix,
    _com_workbook_sheet_names,
)
from excel_mcp.com.read import _normalize_excel_matrix
from excel_mcp.com.session import _coerce_com_count, get_open_workbook_com

# Excel COM constants (avoid importing win32com at module load)
_XL_SORT_ON_VALUES = 0
_XL_YES = 1

def _com_capture_listobject_filters(lo: Any) -> list[dict[str, Any]]:
    """Snapshot ListObject AutoFilter criteria (best-effort, JSON-safe)."""
    out: list[dict[str, Any]] = []
    try:
        af = lo.AutoFilter
    except Exception:
        return out
    if af is None:
        return out
    try:
        filters = af.Filters
        n = _coerce_com_count(getattr(filters, "Count", 0))
    except Exception:
        return out
    for i in range(1, n + 1):
        try:
            f = filters.Item(i)
            on = bool(getattr(f, "On", False))
        except Exception:
            continue
        entry: dict[str, Any] = {"field": i, "on": on}
        if on:
            try:
                entry["criteria1"] = getattr(f, "Criteria1", None)
            except Exception:
                entry["criteria1"] = None
            try:
                entry["criteria2"] = getattr(f, "Criteria2", None)
            except Exception:
                entry["criteria2"] = None
            try:
                entry["operator"] = int(getattr(f, "Operator", 0) or 0)
            except Exception:
                entry["operator"] = 0
        out.append(entry)
    return out

def _com_capture_listobject_sort(
    lo: Any, column_to_field: Mapping[str, int]
) -> Tuple[Optional[list[dict[str, str]]], bool]:
    """Return ``(prior_sort_entries_or_empty, capturable)``.

    ``capturable`` is False when an existing sort uses custom order / non-value
    SortOn / unmappable keys — callers must refuse applying a new sort.
    """
    field_to_column = {v: k for k, v in column_to_field.items()}
    try:
        sort_obj = lo.Sort
        fields = sort_obj.SortFields
        n = _coerce_com_count(getattr(fields, "Count", 0))
    except Exception:
        return [], True
    if n == 0:
        return [], True
    entries: list[dict[str, str]] = []
    for i in range(1, n + 1):
        try:
            sf = fields.Item(i)
        except Exception:
            return None, False
        try:
            custom = getattr(sf, "CustomOrder", None)
            if custom not in (None, "", 0, False):
                return None, False
        except Exception:
            pass
        try:
            sort_on = getattr(sf, "SortOn", _XL_SORT_ON_VALUES)
            if sort_on not in (None, _XL_SORT_ON_VALUES, 0):
                return None, False
        except Exception:
            pass
        try:
            key = sf.Key
            sheet_col = int(key.Column)
        except Exception:
            return None, False
        matched_col: Optional[str] = None
        for field_idx, col_name in field_to_column.items():
            try:
                if _com_list_column_sheet_index(lo, field_idx) == sheet_col:
                    matched_col = col_name
                    break
            except Exception:
                continue
        if matched_col is None:
            return None, False
        try:
            order_raw = int(getattr(sf, "Order", XL_SORT_ASCENDING))
        except Exception:
            order_raw = XL_SORT_ASCENDING
        order = "desc" if order_raw == XL_SORT_DESCENDING else "asc"
        entries.append({"column": matched_col, "order": order})
    return entries, True

def _com_apply_autofilter_steps(rng: Any, steps: Sequence[Mapping[str, Any]]) -> None:
    """Apply compiled AutoFilter steps on a Range (field index is range-relative)."""
    for step in steps:
        field = int(step["field"])
        criteria1 = step.get("criteria1")
        criteria2 = step.get("criteria2")
        operator = step.get("operator")
        kwargs: dict[str, Any] = {"Field": field, "Criteria1": criteria1}
        if operator is not None:
            kwargs["Operator"] = operator
        if criteria2 is not None:
            kwargs["Criteria2"] = criteria2
        rng.AutoFilter(**kwargs)

def _com_restore_autofilter_steps(
    rng: Any, prior_filters: Sequence[Mapping[str, Any]]
) -> None:
    """Re-apply captured filter criteria after ShowAllData."""
    for entry in prior_filters:
        if not entry.get("on"):
            continue
        field = int(entry["field"])
        kwargs: dict[str, Any] = {
            "Field": field,
            "Criteria1": entry.get("criteria1"),
        }
        op = entry.get("operator")
        if op:
            kwargs["Operator"] = op
        if entry.get("criteria2") is not None:
            kwargs["Criteria2"] = entry.get("criteria2")
        rng.AutoFilter(**kwargs)

def _com_apply_listobject_sort(
    lo: Any, sort_spec: Mapping[str, Any], column_to_field: Mapping[str, int]
) -> None:
    sort_obj = lo.Sort
    sort_obj.SortFields.Clear()
    for entry in sort_spec["by"]:
        field = column_to_field[entry["column"]]
        key = lo.ListColumns.Item(field).Range
        order = (
            XL_SORT_DESCENDING if entry["order"] == "desc" else XL_SORT_ASCENDING
        )
        sort_obj.SortFields.Add(
            Key=key,
            SortOn=_XL_SORT_ON_VALUES,
            Order=order,
        )
    sort_obj.Header = _XL_YES
    sort_obj.Apply()

def _com_restore_listobject_sort(
    lo: Any,
    prior_sort: Optional[Sequence[Mapping[str, Any]]],
    column_to_field: Mapping[str, int],
) -> None:
    sort_obj = lo.Sort
    sort_obj.SortFields.Clear()
    if not prior_sort:
        return
    for entry in prior_sort:
        col = str(entry["column"])
        if col not in column_to_field:
            continue
        field = column_to_field[col]
        key = lo.ListColumns.Item(field).Range
        order = (
            XL_SORT_DESCENDING
            if str(entry.get("order", "asc")).lower() == "desc"
            else XL_SORT_ASCENDING
        )
        sort_obj.SortFields.Add(
            Key=key,
            SortOn=_XL_SORT_ON_VALUES,
            Order=order,
        )
    sort_obj.Header = _XL_YES
    sort_obj.Apply()

def _com_hide_unfocused_columns(
    lo: Any,
    focus_columns: Sequence[str],
    table_columns: Sequence[str],
) -> list[int]:
    """Hide currently-visible table columns not in ``focus_columns``; return sheet cols hid."""
    if not focus_columns:
        return []
    focus = set(focus_columns)
    hidden: list[int] = []
    ws = lo.Parent
    for i, name in enumerate(table_columns, start=1):
        if name in focus:
            continue
        sheet_col = _com_list_column_sheet_index(lo, i)
        entire = ws.Columns(sheet_col).EntireColumn
        try:
            already = bool(entire.Hidden)
        except Exception:
            already = False
        if already:
            continue
        entire.Hidden = True
        hidden.append(sheet_col)
    return hidden

def _com_unhide_columns(ws: Any, columns: Sequence[int]) -> None:
    for sheet_col in columns:
        try:
            ws.Columns(int(sheet_col)).EntireColumn.Hidden = False
        except Exception:
            continue

def _com_show_all_listobject_data(lo: Any) -> None:
    """Clear table filters via ``ListObject.AutoFilter.ShowAllData`` only."""
    af = lo.AutoFilter
    if af is None:
        return
    try:
        filter_mode = bool(getattr(af, "FilterMode", False))
    except Exception:
        filter_mode = False
    if filter_mode:
        af.ShowAllData()

def _com_worksheet_autofilter_mode(ws: Any) -> bool:
    try:
        return bool(getattr(ws, "AutoFilterMode", False))
    except Exception:
        return False

def _com_capture_worksheet_filters(ws: Any) -> list[dict[str, Any]]:
    """Snapshot worksheet AutoFilter criteria (best-effort, JSON-safe)."""
    out: list[dict[str, Any]] = []
    if not _com_worksheet_autofilter_mode(ws):
        return out
    try:
        af = ws.AutoFilter
    except Exception:
        return out
    if af is None:
        return out
    try:
        filters = af.Filters
        n = _coerce_com_count(getattr(filters, "Count", 0))
    except Exception:
        return out
    for i in range(1, n + 1):
        try:
            f = filters.Item(i)
            on = bool(getattr(f, "On", False))
        except Exception:
            continue
        entry: dict[str, Any] = {"field": i, "on": on}
        if on:
            try:
                entry["criteria1"] = getattr(f, "Criteria1", None)
            except Exception:
                entry["criteria1"] = None
            try:
                entry["criteria2"] = getattr(f, "Criteria2", None)
            except Exception:
                entry["criteria2"] = None
            try:
                entry["operator"] = int(getattr(f, "Operator", 0) or 0)
            except Exception:
                entry["operator"] = 0
        out.append(entry)
    return out

def _com_show_all_worksheet_data(ws: Any) -> None:
    """Clear active worksheet AutoFilter criteria when FilterMode is true."""
    try:
        filter_mode = bool(getattr(ws, "FilterMode", False))
    except Exception:
        filter_mode = False
    if not filter_mode:
        return
    try:
        af = ws.AutoFilter
    except Exception:
        af = None
    if af is not None:
        try:
            af.ShowAllData()
            return
        except Exception:
            pass
    try:
        ws.ShowAllData()
    except Exception:
        pass

def _com_capture_worksheet_sort(
    ws: Any, column_to_field: Mapping[str, int], min_col: int
) -> Tuple[Optional[list[dict[str, str]]], bool]:
    """Return ``(prior_sort_entries_or_empty, capturable)`` for ``Worksheet.Sort``."""
    field_to_column = {v: k for k, v in column_to_field.items()}
    try:
        sort_obj = ws.Sort
        fields = sort_obj.SortFields
        n = _coerce_com_count(getattr(fields, "Count", 0))
    except Exception:
        return [], True
    if n == 0:
        return [], True
    entries: list[dict[str, str]] = []
    for i in range(1, n + 1):
        try:
            sf = fields.Item(i)
        except Exception:
            return None, False
        try:
            custom = getattr(sf, "CustomOrder", None)
            if custom not in (None, "", 0, False):
                return None, False
        except Exception:
            pass
        try:
            sort_on = getattr(sf, "SortOn", _XL_SORT_ON_VALUES)
            if sort_on not in (None, _XL_SORT_ON_VALUES, 0):
                return None, False
        except Exception:
            pass
        try:
            key = sf.Key
            sheet_col = int(key.Column)
        except Exception:
            return None, False
        field_idx = sheet_col - min_col + 1
        matched_col = field_to_column.get(field_idx)
        if matched_col is None:
            return None, False
        try:
            order_raw = int(getattr(sf, "Order", XL_SORT_ASCENDING))
        except Exception:
            order_raw = XL_SORT_ASCENDING
        order = "desc" if order_raw == XL_SORT_DESCENDING else "asc"
        entries.append({"column": matched_col, "order": order})
    return entries, True

def _com_apply_worksheet_sort(
    ws: Any,
    rng: Any,
    sort_spec: Mapping[str, Any],
    column_to_field: Mapping[str, int],
) -> None:
    sort_obj = ws.Sort
    sort_obj.SortFields.Clear()
    for entry in sort_spec["by"]:
        field = column_to_field[entry["column"]]
        key = rng.Columns(field)
        order = (
            XL_SORT_DESCENDING if entry["order"] == "desc" else XL_SORT_ASCENDING
        )
        sort_obj.SortFields.Add(
            Key=key,
            SortOn=_XL_SORT_ON_VALUES,
            Order=order,
        )
    sort_obj.SetRange(rng)
    sort_obj.Header = _XL_YES
    sort_obj.Apply()

def _com_restore_worksheet_sort(
    ws: Any,
    rng: Any,
    prior_sort: Optional[Sequence[Mapping[str, Any]]],
    column_to_field: Mapping[str, int],
) -> None:
    sort_obj = ws.Sort
    sort_obj.SortFields.Clear()
    if not prior_sort:
        sort_obj.SetRange(rng)
        sort_obj.Header = _XL_YES
        sort_obj.Apply()
        return
    for entry in prior_sort:
        col = str(entry["column"])
        if col not in column_to_field:
            continue
        field = column_to_field[col]
        key = rng.Columns(field)
        order = (
            XL_SORT_DESCENDING
            if str(entry.get("order", "asc")).lower() == "desc"
            else XL_SORT_ASCENDING
        )
        sort_obj.SortFields.Add(
            Key=key,
            SortOn=_XL_SORT_ON_VALUES,
            Order=order,
        )
    sort_obj.SetRange(rng)
    sort_obj.Header = _XL_YES
    sort_obj.Apply()

def _com_hide_unfocused_range_columns(
    ws: Any,
    focus_columns: Sequence[str],
    region_columns: Sequence[str],
    min_col: int,
) -> list[int]:
    """Hide currently-visible region columns not in ``focus_columns``."""
    if not focus_columns:
        return []
    focus = set(focus_columns)
    hidden: list[int] = []
    for i, name in enumerate(region_columns, start=1):
        if name in focus:
            continue
        sheet_col = min_col + i - 1
        entire = ws.Columns(sheet_col).EntireColumn
        try:
            already = bool(entire.Hidden)
        except Exception:
            already = False
        if already:
            continue
        entire.Hidden = True
        hidden.append(sheet_col)
    return hidden

def _com_region_header_columns(ws: Any, range_a1: str) -> tuple[list[str], Any, int]:
    """Return ``(headers, range_com, min_col)`` for a plain region (first row = headers)."""
    rng = ws.Range(normalize_excel_a1_range(range_a1))
    min_col, _min_row, _max_col, _max_row = range_boundaries(
        normalize_excel_a1_range(range_a1)
    )
    matrix = _normalize_excel_matrix(rng.Value2)
    if not matrix:
        raise DataError(f"Region range {range_a1!r} is empty")
    header_row = list(matrix[0])
    columns = [
        str(v).strip() if v is not None and str(v).strip() != "" else ""
        for v in header_row
    ]
    if not any(columns):
        raise DataError(f"Region range {range_a1!r} has no header row")
    return columns, rng, int(min_col)

_SORT_CAPTURE_WARNING = (
    "In-place sort was not applied because the previous row order cannot be "
    "restored (no capturable prior sort). Use snapshot mode for a reorder "
    "that leaves the source order unchanged."
)

def _decide_in_place_sort(
    sort_spec: Optional[Mapping[str, Any]],
    prior_sort: Optional[Sequence[Mapping[str, Any]]],
    sort_capturable: bool,
) -> tuple[bool, Optional[str]]:
    """Apply an in-place sort only when the previous sort can be reapplied.

    Excel sort reorders rows and clearing SortFields does not undo that.
    An empty prior sort is not restorable, so the sort portion is refused.
    """
    if sort_spec is None:
        return False, None
    if not sort_capturable or not prior_sort:
        return False, _SORT_CAPTURE_WARNING
    return True, None

def _snapshot_needed_columns(
    headers: Sequence[str],
    where: Sequence[Mapping[str, Any]],
    sort: Optional[Mapping[str, Any]],
) -> list[str]:
    sort_cols = (
        [str(e["column"]) for e in (sort or {}).get("by", [])] if sort else []
    )
    return list(
        dict.fromkeys(list(headers) + [c["column"] for c in where] + sort_cols)
    )

def _com_delete_sheet_quiet(wb_com: Any, ws: Any) -> None:
    """Best-effort delete with DisplayAlerts suppressed."""
    try:
        app = wb_com.Application
        prev = app.DisplayAlerts
        app.DisplayAlerts = False
        try:
            ws.Delete()
        finally:
            app.DisplayAlerts = prev
    except Exception:
        pass

def _com_load_snapshot_table_source(
    wb_com: Any,
    view_spec: Mapping[str, Any],
    view_applicability: Optional[Mapping[str, Any]],
) -> tuple[Optional[dict[str, Any]], Optional[str]]:
    """Return ``(ctx, error)`` for a ListObject snapshot source."""
    target = view_spec["target"]
    name = target.get("name")
    if not isinstance(name, str) or not name.strip():
        raise DataError("view_spec.target.name is required for a table view")
    lo, find_err = _com_find_list_object(wb_com, name.strip())
    if find_err:
        return None, find_err
    assert lo is not None
    source_columns = _com_list_object_column_names(lo)
    normalized = validate_view_spec_for_apply(
        view_spec,
        mode=VIEW_MODE_SNAPSHOT,
        view_applicability=view_applicability,
        table_columns=source_columns,
    )
    headers = resolve_query_columns(source_columns, normalized["columns"] or None)
    needed = _snapshot_needed_columns(
        headers, normalized["where"], normalized["sort"]
    )
    return {
        "normalized": normalized,
        "headers": headers,
        "rows": _com_list_object_row_dicts(lo, needed),
        "label": str(lo.Name),
    }, None

def _com_load_snapshot_region_source(
    wb_com: Any,
    view_spec: Mapping[str, Any],
    view_applicability: Optional[Mapping[str, Any]],
) -> tuple[Optional[dict[str, Any]], Optional[str]]:
    """Return ``(ctx, error)`` for a plain-range snapshot source."""
    target = view_spec["target"]
    sheet_raw = target.get("sheet")
    range_raw = target.get("range")
    if not isinstance(sheet_raw, str) or not sheet_raw.strip():
        raise DataError("view_spec.target.sheet is required for a region view")
    if not isinstance(range_raw, str) or not range_raw.strip():
        raise DataError("view_spec.target.range is required for a region view")
    sheet_name = sheet_raw.strip()
    range_a1 = normalize_excel_a1_range(range_raw)
    try:
        source_ws = wb_com.Worksheets(sheet_name)
    except Exception:
        return None, f"Error: Sheet '{sheet_name}' not found"
    matrix = _com_read_a1_matrix(source_ws, range_a1)
    source_columns, source_rows, _rng = resolve_region_query_source(
        matrix=matrix,
        range_a1=range_a1,
        require_header_guess=False,
    )
    normalized = validate_view_spec_for_apply(
        view_spec,
        mode=VIEW_MODE_SNAPSHOT,
        view_applicability=view_applicability,
        table_columns=source_columns,
    )
    headers = resolve_query_columns(source_columns, normalized["columns"] or None)
    return {
        "normalized": normalized,
        "headers": headers,
        "rows": source_rows,
        "label": f"{sheet_name}!{range_a1}",
    }, None

def _com_write_value_matrix(ws: Any, matrix: Sequence[Sequence[Any]]) -> None:
    """Write a rectangular value grid starting at A1 (no ListObject / Copy)."""
    if not matrix:
        return
    nrows = len(matrix)
    ncols = max((len(row) for row in matrix), default=0)
    if ncols == 0:
        return
    grid: list[list[Any]] = []
    for row in matrix:
        padded = list(row) + [None] * (ncols - len(row))
        grid.append(padded)
    start = ws.Range("A1")
    rng = start.Resize(nrows, ncols)
    if nrows == 1 and ncols == 1:
        rng.Value = grid[0][0]
    else:
        rng.Value = tuple(tuple(r) for r in grid)

def apply_table_view_com(
    filepath: str,
    view_spec: dict[str, Any],
    mode: str,
    view_applicability: Optional[dict[str, Any]],
) -> str:
    wb_com, err = get_open_workbook_com(
        filepath, for_write=True
    )
    if err:
        return err
    try:
        if not isinstance(view_spec, dict):
            raise DataError("view_spec must be an object")
        mode_norm = (mode or "").strip().lower()
        if mode_norm == VIEW_MODE_SNAPSHOT:
            return apply_snapshot_view_com(
                wb_com, view_spec, view_applicability
            )

        target = view_spec.get("target")
        target_kind = (
            str(target.get("kind", "")).strip().lower()
            if isinstance(target, Mapping)
            else ""
        )

        if target_kind == "table":
            return apply_listobject_view_com(
                wb_com, view_spec, mode, view_applicability
            )
        if target_kind == "region":
            return apply_range_view_com(
                wb_com, view_spec, mode, view_applicability
            )

        # Let validate_view_spec_for_apply raise the consistent error.
        validate_view_spec_for_apply(
            view_spec,
            mode=mode,
            view_applicability=view_applicability,
        )
        return "Error: Unsupported view_spec.target.kind"
    except DataError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Error: {e}"

def apply_snapshot_view_com(
    wb_com: Any,
    view_spec: dict[str, Any],
    view_applicability: Optional[dict[str, Any]],
) -> str:
    """Write projected/filtered/sorted values onto a new sheet; leave source alone."""
    target = view_spec.get("target")
    if not isinstance(target, Mapping):
        raise DataError("view_spec.target is required")
    kind = str(target.get("kind", "")).strip().lower()
    if kind == "table":
        ctx, err = _com_load_snapshot_table_source(
            wb_com, view_spec, view_applicability
        )
    elif kind == "region":
        ctx, err = _com_load_snapshot_region_source(
            wb_com, view_spec, view_applicability
        )
    else:
        raise DataError(
            f"view_spec.target.kind {kind!r} is not supported for "
            "apply_table_view; expected 'table' or 'region'. "
            "Sheet was not changed."
        )
    if err:
        return err
    assert ctx is not None

    normalized = ctx["normalized"]
    matching = filter_table_rows(ctx["rows"], normalized["where"])
    matching = sort_table_rows(matching, normalized["sort"])
    value_matrix = build_snapshot_value_matrix(
        headers=ctx["headers"],
        matching_rows=matching,
        limit=normalized.get("limit"),
        offset=int(normalized.get("offset") or 0),
    )

    new_name = allocate_unique_sheet_name(_com_workbook_sheet_names(wb_com))
    new_ws = None
    try:
        new_ws = wb_com.Worksheets.Add()
        new_ws.Name = new_name
        _com_write_value_matrix(new_ws, value_matrix)
    except Exception as exc:
        if new_ws is not None:
            _com_delete_sheet_quiet(wb_com, new_ws)
        return f"Error: {exc}"

    token = build_restore_token(
        kind=RESTORE_KIND_SNAPSHOT,
        sheet=new_name,
        created_by_tool=True,
    )
    return json.dumps(
        {
            "message": "Snapshot view applied",
            "mode": VIEW_MODE_SNAPSHOT,
            "sheet": new_name,
            "source": ctx["label"],
            "row_count": max(0, len(value_matrix) - 1),
            "restore_token": token,
        },
        indent=2,
        default=str,
    )

def apply_listobject_view_com(
    wb_com: Any,
    view_spec: dict[str, Any],
    mode: str,
    view_applicability: Optional[dict[str, Any]],
) -> str:
    target = view_spec.get("target")
    table_columns: Optional[list[str]] = None
    lo = None
    if isinstance(target, Mapping):
        name = target.get("name")
        if isinstance(name, str) and name.strip():
            lo, find_err = _com_find_list_object(wb_com, name)
            if find_err:
                return find_err
            assert lo is not None
            table_columns = _com_list_object_column_names(lo)

    normalized = validate_view_spec_for_apply(
        view_spec,
        mode=mode,
        view_applicability=view_applicability,
        table_columns=table_columns,
    )
    if lo is None or table_columns is None:
        return f"Error: Table '{normalized['target_name']}' not found"

    column_to_field = _com_list_object_field_map(lo)
    steps = compile_autofilter_field_steps(
        normalized["where"], column_to_field=column_to_field
    )

    prior_filters = _com_capture_listobject_filters(lo)
    prior_sort, sort_capturable = _com_capture_listobject_sort(
        lo, column_to_field
    )
    sort_spec = normalized["sort"]
    sort_applied, sort_warning = _decide_in_place_sort(
        sort_spec, prior_sort, sort_capturable
    )

    # Mutate only after validation + capture.
    _com_show_all_listobject_data(lo)
    if steps:
        _com_apply_autofilter_steps(lo.Range, steps)
    if sort_applied and sort_spec is not None:
        _com_apply_listobject_sort(lo, sort_spec, column_to_field)

    columns_hidden = _com_hide_unfocused_columns(
        lo, normalized["columns"], table_columns
    )
    try:
        sheet_name = str(lo.Parent.Name)
    except Exception:
        sheet_name = ""

    token = build_restore_token(
        kind=RESTORE_KIND_LISTOBJECT,
        table_name=str(lo.Name),
        sheet=sheet_name,
        prior_filters=prior_filters,
        prior_sort=prior_sort if sort_capturable else None,
        sort_applied=sort_applied,
        columns_hidden=columns_hidden,
    )
    payload: dict[str, Any] = {
        "message": "Table view applied",
        "mode": "in_place",
        "table": str(lo.Name),
        "sheet": sheet_name,
        "restore_token": token,
    }
    if sort_warning:
        payload["warnings"] = [sort_warning]
    return json.dumps(payload, indent=2, default=str)

def apply_range_view_com(
    wb_com: Any,
    view_spec: dict[str, Any],
    mode: str,
    view_applicability: Optional[dict[str, Any]],
) -> str:
    target = view_spec.get("target")
    if not isinstance(target, Mapping):
        raise DataError("view_spec.target is required")
    sheet_raw = target.get("sheet")
    range_raw = target.get("range")
    if not isinstance(sheet_raw, str) or not sheet_raw.strip():
        raise DataError("view_spec.target.sheet is required for a region view")
    if not isinstance(range_raw, str) or not range_raw.strip():
        raise DataError("view_spec.target.range is required for a region view")
    sheet_name = sheet_raw.strip()
    range_a1 = normalize_excel_a1_range(range_raw)
    try:
        ws = wb_com.Worksheets(sheet_name)
    except Exception:
        return f"Error: Sheet '{sheet_name}' not found"

    region_columns, rng, min_col = _com_region_header_columns(ws, range_a1)
    normalized = validate_view_spec_for_apply(
        view_spec,
        mode=mode,
        view_applicability=view_applicability,
        table_columns=region_columns,
    )
    column_to_field = {
        name: i + 1 for i, name in enumerate(region_columns) if name
    }
    steps = compile_autofilter_field_steps(
        normalized["where"], column_to_field=column_to_field
    )

    prior_autofilter_mode = _com_worksheet_autofilter_mode(ws)
    prior_filters = _com_capture_worksheet_filters(ws)
    prior_sort, sort_capturable = _com_capture_worksheet_sort(
        ws, column_to_field, min_col
    )
    sort_spec = normalized["sort"]
    sort_applied, sort_warning = _decide_in_place_sort(
        sort_spec, prior_sort, sort_capturable
    )

    # Mutate only after validation + capture. Never ListObjects.Add / Worksheet.Copy.
    _com_show_all_worksheet_data(ws)
    if steps:
        _com_apply_autofilter_steps(rng, steps)
    elif not prior_autofilter_mode:
        # Enable AutoFilter arrows on the range even with no criteria.
        rng.AutoFilter()
    if sort_applied and sort_spec is not None:
        _com_apply_worksheet_sort(ws, rng, sort_spec, column_to_field)

    columns_hidden = _com_hide_unfocused_range_columns(
        ws, normalized["columns"], region_columns, min_col
    )

    token = build_restore_token(
        kind=RESTORE_KIND_RANGE,
        sheet=sheet_name,
        range_a1=range_a1,
        prior_autofilter_mode=prior_autofilter_mode,
        prior_filters=prior_filters,
        prior_sort=prior_sort if sort_capturable else None,
        sort_applied=sort_applied,
        columns_hidden=columns_hidden,
    )
    payload: dict[str, Any] = {
        "message": "Region view applied",
        "mode": "in_place",
        "sheet": sheet_name,
        "range": range_a1,
        "restore_token": token,
    }
    if sort_warning:
        payload["warnings"] = [sort_warning]
    return json.dumps(payload, indent=2, default=str)

def clear_table_view_com(
    filepath: str, restore_token: dict[str, Any]
) -> str:
    wb_com, err = get_open_workbook_com(
        filepath, for_write=True
    )
    if err:
        return err
    try:
        token = parse_restore_token(restore_token)
        if token["kind"] == RESTORE_KIND_LISTOBJECT:
            return clear_listobject_view_com(wb_com, token)
        if token["kind"] == RESTORE_KIND_RANGE:
            return clear_range_view_com(wb_com, token)
        return clear_snapshot_view_com(wb_com, token)
    except DataError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        return f"Error: {e}"

def clear_snapshot_view_com(wb_com: Any, token: Mapping[str, Any]) -> str:
    """Delete the sheet this tool created for a snapshot restore_token only."""
    sheet_name = token["sheet"]
    try:
        ws = wb_com.Worksheets(sheet_name)
    except Exception:
        return f"Error: Sheet '{sheet_name}' not found"

    app = wb_com.Application
    prev = app.DisplayAlerts
    app.DisplayAlerts = False
    try:
        ws.Delete()
    except Exception as exc:
        return f"Error: {exc}"
    finally:
        app.DisplayAlerts = prev

    return json.dumps(
        {"message": "Snapshot view cleared", "sheet": sheet_name},
        indent=2,
        default=str,
    )

def clear_listobject_view_com(wb_com: Any, token: Mapping[str, Any]) -> str:
    lo, find_err = _com_find_list_object(wb_com, token["table"])
    if find_err:
        return find_err
    assert lo is not None
    try:
        ws = lo.Parent
    except Exception:
        return f"Error: Cannot resolve parent sheet for table '{token['table']}'"

    # Never clear via Worksheet.AutoFilterMode for ListObject tokens.
    _com_show_all_listobject_data(lo)
    if token["prior_filters"]:
        _com_restore_autofilter_steps(lo.Range, token["prior_filters"])

    if token["sort_applied"]:
        column_to_field = _com_list_object_field_map(lo)
        _com_restore_listobject_sort(
            lo, token["prior_sort"], column_to_field
        )

    _com_unhide_columns(ws, token["columns_hidden"])

    payload = {
        "message": "Table view cleared",
        "table": token["table"],
        "sheet": token["sheet"],
    }
    return json.dumps(payload, indent=2, default=str)

def clear_range_view_com(wb_com: Any, token: Mapping[str, Any]) -> str:
    sheet_name = token["sheet"]
    range_a1 = token["range"]
    try:
        ws = wb_com.Worksheets(sheet_name)
    except Exception:
        return f"Error: Sheet '{sheet_name}' not found"
    rng = ws.Range(normalize_excel_a1_range(range_a1))

    # Plain-range clear uses worksheet AutoFilterMode / ShowAllData — not
    # ListObject.AutoFilter.ShowAllData().
    _com_show_all_worksheet_data(ws)
    if token["prior_autofilter_mode"]:
        if token["prior_filters"]:
            _com_restore_autofilter_steps(rng, token["prior_filters"])
        elif not _com_worksheet_autofilter_mode(ws):
            rng.AutoFilter()
    else:
        try:
            ws.AutoFilterMode = False
        except Exception:
            pass

    if token["sort_applied"]:
        region_columns, rng, _min_col = _com_region_header_columns(ws, range_a1)
        column_to_field = {
            name: i + 1 for i, name in enumerate(region_columns) if name
        }
        _com_restore_worksheet_sort(
            ws, rng, token["prior_sort"], column_to_field
        )

    _com_unhide_columns(ws, token["columns_hidden"])

    payload = {
        "message": "Region view cleared",
        "sheet": sheet_name,
        "range": range_a1,
    }
    return json.dumps(payload, indent=2, default=str)

