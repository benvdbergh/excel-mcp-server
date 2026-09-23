"""View-spec validation, AutoFilter compilation, restore tokens, snapshots."""

from typing import Any, Mapping, Optional, Sequence

from excel_mcp.exceptions import DataError

from .layout import normalize_excel_a1_range
from .rows import (
    _VIEWABLE_OPS,
    _as_text,
    normalize_where_clauses,
)

# apply_table_view / clear_table_view (Epic 14 / STORY-14-1)
VIEW_MODE_IN_PLACE = "in_place"
VIEW_MODE_SNAPSHOT = "snapshot"
_VALID_VIEW_MODES = frozenset({VIEW_MODE_IN_PLACE, VIEW_MODE_SNAPSHOT})
_VALID_SORT_ORDERS = frozenset({"asc", "desc"})
RESTORE_TOKEN_VERSION = 1
# Excel XlAutoFilterOperator (numeric; avoid win32com import in shared module)
XL_FILTER_OR = 2
XL_FILTER_VALUES = 7
XL_SORT_ASCENDING = 1
XL_SORT_DESCENDING = 2

RESTORE_KIND_LISTOBJECT = "listobject"
RESTORE_KIND_RANGE = "range"
RESTORE_KIND_SNAPSHOT = "snapshot"
_VALID_RESTORE_KINDS = frozenset(
    {RESTORE_KIND_LISTOBJECT, RESTORE_KIND_RANGE, RESTORE_KIND_SNAPSHOT}
)

SNAPSHOT_SHEET_NAME_BASE = "mcp_view"


def normalize_view_sort(
    sort: Any,
    *,
    known_columns: Sequence[str],
) -> Optional[dict[str, Any]]:
    """Normalize ``view_spec.sort`` to ``{\"by\": [{\"column\", \"order\"}]}`` or ``None``.

    Accepted shapes:
    - ``None``
    - ``{\"by\": [{\"column\": \"Qty\", \"order\": \"asc\"|\"desc\"}, ...]}``
    - ``{\"column\": \"Qty\", \"order\": \"asc\"|\"desc\"}`` (single-key sugar)
    """
    if sort is None:
        return None
    if not isinstance(sort, Mapping):
        raise DataError("view_spec.sort must be an object or null")
    known = set(known_columns)
    if "by" in sort:
        raw_by = sort["by"]
        if not isinstance(raw_by, (list, tuple)) or not raw_by:
            raise DataError("view_spec.sort.by must be a non-empty list")
        entries = list(raw_by)
    elif "column" in sort:
        entries = [sort]
    else:
        raise DataError(
            "view_spec.sort must include 'by' or 'column'; "
            "expected {\"by\": [{\"column\", \"order\"}]}"
        )
    out: list[dict[str, str]] = []
    for i, entry in enumerate(entries):
        if not isinstance(entry, Mapping):
            raise DataError(f"Invalid sort entry at index {i}: expected object")
        column = entry.get("column")
        order_raw = entry.get("order", "asc")
        if not isinstance(column, str) or not column:
            raise DataError(f"Invalid sort entry at index {i}: column is required")
        if column not in known:
            raise DataError(f"Unknown sort column: {column!r}")
        if not isinstance(order_raw, str) or not order_raw:
            raise DataError(f"Invalid sort entry at index {i}: order is required")
        order = order_raw.strip().lower()
        if order not in _VALID_SORT_ORDERS:
            raise DataError(
                f"Invalid sort order: {order_raw!r}; expected 'asc' or 'desc'"
            )
        out.append({"column": column, "order": order})
    return {"by": out}


def _reject_cross_column_or(where: Sequence[Mapping[str, Any]]) -> None:
    """Refuse nested OR groups (use single-column ``in`` instead)."""
    for i, raw in enumerate(where):
        if not isinstance(raw, Mapping):
            continue
        op = str(raw.get("op", "")).strip().lower()
        if op != "or":
            continue
        nested = raw.get("clauses") or raw.get("where")
        cols = {
            str(c.get("column"))
            for c in (nested if isinstance(nested, (list, tuple)) else ())
            if isinstance(c, Mapping) and c.get("column")
        }
        if len(cols) > 1:
            raise DataError(
                "Cross-column OR is not viewable; use single-column 'in' or AND clauses"
            )
        raise DataError(
            "Nested OR groups are not supported; use operator 'in' on one column"
        )


def normalize_snapshot_pagination(
    limit: Any, offset: Any
) -> tuple[Optional[int], int]:
    """Parse optional limit/offset for snapshot (no default page size of 100).

    ``limit is None`` means write every remaining matching row after ``offset``.
    """
    if offset is None:
        off = 0
    else:
        try:
            off = int(offset)
        except (TypeError, ValueError) as e:
            raise DataError(f"Invalid offset: {offset!r}") from e
        if off < 0:
            raise DataError(f"Invalid offset: {offset!r}; must be >= 0")
    if limit is None:
        return None, off
    try:
        lim = int(limit)
    except (TypeError, ValueError) as e:
        raise DataError(f"Invalid limit: {limit!r}") from e
    if lim < 0:
        raise DataError(f"Invalid limit: {limit!r}; must be >= 0")
    return lim, off


def build_snapshot_value_matrix(
    *,
    headers: Sequence[str],
    matching_rows: Sequence[Mapping[str, Any]],
    limit: Optional[int],
    offset: int,
) -> list[list[Any]]:
    """Header row plus projected values for a snapshot sheet (values only)."""
    if offset < 0:
        raise DataError(f"Invalid offset: {offset!r}; must be >= 0")
    if limit is None:
        page = list(matching_rows[offset:])
    else:
        if limit < 0:
            raise DataError(f"Invalid limit: {limit!r}; must be >= 0")
        page = list(matching_rows[offset : offset + limit])
    matrix: list[list[Any]] = [list(headers)]
    for row in page:
        matrix.append([row.get(h) for h in headers])
    return matrix


def allocate_unique_sheet_name(
    existing_names: Sequence[str],
    *,
    base: str = SNAPSHOT_SHEET_NAME_BASE,
) -> str:
    """First free name: ``base``, then ``base_2``, ``base_3``, ... (case-insensitive)."""
    taken = {str(n).strip().lower() for n in existing_names if str(n).strip()}
    candidate = base
    if candidate.lower() not in taken:
        return candidate
    n = 2
    while True:
        candidate = f"{base}_{n}"
        if candidate.lower() not in taken:
            return candidate
        n += 1


def validate_view_spec_for_apply(
    view_spec: Mapping[str, Any],
    *,
    mode: str,
    view_applicability: Optional[Mapping[str, Any]] = None,
    table_columns: Optional[Sequence[str]] = None,
) -> dict[str, Any]:
    """Validate ``view_spec`` + ``mode`` for ``apply_table_view``.

    Supports ``target.kind`` ``table`` (ListObject) and ``region`` (plain range).
    ``in_place`` refuses limit/offset and ``view_applicability.viewable is false``.
    ``snapshot`` honors limit/offset (no default page size) and does not treat
    limit/offset ``view_applicability`` as a blocker. Both modes refuse
    cross-column OR and non-viewable operators (e.g. ``is_empty``).
    """
    if not isinstance(view_spec, Mapping):
        raise DataError("view_spec must be an object")
    mode_norm = (mode or "").strip().lower()
    if mode_norm not in _VALID_VIEW_MODES:
        raise DataError(
            f"Invalid mode: {mode!r}; expected '{VIEW_MODE_IN_PLACE}' or "
            f"'{VIEW_MODE_SNAPSHOT}'"
        )

    snapshot_limit: Optional[int] = None
    snapshot_offset = 0
    if mode_norm == VIEW_MODE_IN_PLACE:
        if "limit" in view_spec and view_spec.get("limit") is not None:
            raise DataError(
                "view_spec must not include limit; Excel AutoFilter cannot show a row page"
            )
        if "offset" in view_spec and view_spec.get("offset"):
            raise DataError(
                "view_spec must not include offset; Excel AutoFilter cannot skip rows"
            )
        if view_applicability is not None:
            if not isinstance(view_applicability, Mapping):
                raise DataError("view_applicability must be an object when provided")
            if view_applicability.get("viewable") is False:
                reason = view_applicability.get("reason") or (
                    "view_applicability.viewable is false"
                )
                raise DataError(
                    f"Cannot apply view: {reason}. Sheet was not changed."
                )
    else:
        # Snapshot: limit/offset are part of the copied result set.
        if view_applicability is not None and not isinstance(
            view_applicability, Mapping
        ):
            raise DataError("view_applicability must be an object when provided")
        snapshot_limit, snapshot_offset = normalize_snapshot_pagination(
            view_spec.get("limit") if "limit" in view_spec else None,
            view_spec.get("offset") if "offset" in view_spec else None,
        )
        if isinstance(view_applicability, Mapping) and (
            view_applicability.get("viewable") is False
        ):
            reason = str(view_applicability.get("reason") or "").lower()
            page_reason = any(
                token in reason for token in ("limit", "offset", "truncat")
            )
            spec_has_page = (
                "limit" in view_spec and view_spec.get("limit") is not None
            ) or bool(view_spec.get("offset"))
            if page_reason and not spec_has_page:
                raise DataError(
                    "Cannot snapshot this query page: view_applicability says "
                    "rows were limited or offset, but view_spec has no limit or "
                    "offset. Add them so the sheet matches the queried rows. "
                    "Sheet was not changed."
                )

    target = view_spec.get("target")
    if not isinstance(target, Mapping):
        raise DataError("view_spec.target is required")
    kind = str(target.get("kind", "")).strip().lower()
    target_name: Optional[str] = None
    target_sheet: Optional[str] = None
    target_range: Optional[str] = None
    if kind == "table":
        name = target.get("name")
        if not isinstance(name, str) or not name.strip():
            raise DataError("view_spec.target.name is required for a table view")
        target_name = name.strip()
    elif kind == "region":
        sheet = target.get("sheet")
        range_raw = target.get("range")
        if not isinstance(sheet, str) or not sheet.strip():
            raise DataError("view_spec.target.sheet is required for a region view")
        if not isinstance(range_raw, str) or not range_raw.strip():
            raise DataError("view_spec.target.range is required for a region view")
        target_sheet = sheet.strip()
        target_range = normalize_excel_a1_range(range_raw)
    else:
        raise DataError(
            f"view_spec.target.kind {kind!r} is not supported for apply_table_view; "
            "expected 'table' or 'region'. Sheet was not changed."
        )

    columns_raw = view_spec.get("columns")
    if columns_raw is None:
        columns: list[str] = []
    elif isinstance(columns_raw, (list, tuple)):
        columns = [str(c) for c in columns_raw]
    else:
        raise DataError("view_spec.columns must be a list of column names")

    where_raw = view_spec.get("where")
    known = list(table_columns) if table_columns is not None else columns
    if table_columns is not None and columns:
        unknown = [c for c in columns if c not in set(table_columns)]
        if unknown:
            raise DataError(f"Unknown column: {unknown[0]!r}")
    # Detect nested OR before normalize_where_clauses rejects unknown ops.
    _reject_cross_column_or(where_raw if isinstance(where_raw, (list, tuple)) else [])
    clauses = normalize_where_clauses(
        where_raw if where_raw is not None else [],
        known_columns=known if known else columns,
    )

    for clause in clauses:
        if clause["op"] not in _VIEWABLE_OPS:
            raise DataError(
                f"Operator {clause['op']!r} is not viewable; sheet was not changed"
            )

    sort = normalize_view_sort(view_spec.get("sort"), known_columns=known if known else columns)
    out: dict[str, Any] = {
        "target_kind": kind,
        "columns": columns,
        "where": clauses,
        "sort": sort,
        "mode": mode_norm,
    }
    if kind == "table":
        out["target_name"] = target_name
    else:
        out["target_sheet"] = target_sheet
        out["target_range"] = target_range
    if mode_norm == VIEW_MODE_SNAPSHOT:
        out["limit"] = snapshot_limit
        out["offset"] = snapshot_offset
    return out


def _excel_filter_literal(value: Any) -> str:
    """Escape Excel AutoFilter wildcards so a contains pattern stays literal."""
    text = _as_text(value)
    return text.replace("~", "~~").replace("*", "~*").replace("?", "~?")


def criteria_text_for_clause(clause: Mapping[str, Any]) -> Any:
    """Map one viewable where clause to AutoFilter Criteria1 / operator payload."""
    op = clause["op"]
    value = clause.get("value")
    if op == "eq":
        return {"criteria1": value}
    if op == "neq":
        return {"criteria1": f"<>{_as_text(value)}"}
    if op == "contains":
        return {"criteria1": f"*{_excel_filter_literal(value)}*"}
    if op == "gt":
        return {"criteria1": f">{_as_text(value)}"}
    if op == "gte":
        return {"criteria1": f">={_as_text(value)}"}
    if op == "lt":
        return {"criteria1": f"<{_as_text(value)}"}
    if op == "lte":
        return {"criteria1": f"<={_as_text(value)}"}
    if op == "in":
        items = list(value) if isinstance(value, (list, tuple)) else [value]
        if len(items) == 0:
            raise DataError("Operator 'in' requires a non-empty list")
        if len(items) == 1:
            return {"criteria1": items[0]}
        if len(items) == 2:
            return {
                "criteria1": items[0],
                "criteria2": items[1],
                "operator": XL_FILTER_OR,
            }
        return {"criteria1": list(items), "operator": XL_FILTER_VALUES}
    raise DataError(f"Operator {op!r} is not viewable")


def compile_autofilter_field_steps(
    where: Sequence[Mapping[str, Any]],
    *,
    column_to_field: Mapping[str, int],
) -> list[dict[str, Any]]:
    """Compile AND-combined where clauses into per-field AutoFilter steps.

    ``column_to_field`` maps header name → 1-based field index relative to the
    table header (not worksheet column A).
    """
    # Multiple clauses on the same column are not expressible as one AutoFilter
    # field without AdvancedFilter; refuse rather than apply a partial view.
    seen: dict[str, Mapping[str, Any]] = {}
    for clause in where:
        col = str(clause["column"])
        if col in seen:
            raise DataError(
                f"Multiple where clauses on column {col!r} cannot be shown as one "
                "AutoFilter field; sheet was not changed"
            )
        seen[col] = clause

    steps: list[dict[str, Any]] = []
    for col, clause in seen.items():
        if col not in column_to_field:
            raise DataError(f"Unknown column: {col!r}")
        field = int(column_to_field[col])
        if field < 1:
            raise DataError(f"Invalid field index for column {col!r}")
        crit = criteria_text_for_clause(clause)
        step: dict[str, Any] = {"field": field, "column": col, **crit}
        steps.append(step)
    return steps


def build_restore_token(
    *,
    sheet: str,
    prior_filters: Sequence[Mapping[str, Any]] = (),
    prior_sort: Optional[Sequence[Mapping[str, Any]]] = None,
    sort_applied: bool = False,
    columns_hidden: Sequence[int] = (),
    kind: str = RESTORE_KIND_LISTOBJECT,
    table_name: Optional[str] = None,
    range_a1: Optional[str] = None,
    prior_autofilter_mode: Optional[bool] = None,
    created_by_tool: bool = False,
) -> dict[str, Any]:
    """Opaque-ish restore payload for ``clear_table_view`` (JSON-serializable)."""
    kind_norm = (kind or "").strip().lower()
    if kind_norm not in _VALID_RESTORE_KINDS:
        raise DataError(f"Unsupported restore_token kind: {kind!r}")
    if not isinstance(sheet, str) or not sheet.strip():
        raise DataError("sheet is required for restore_token")
    if kind_norm == RESTORE_KIND_SNAPSHOT:
        if not created_by_tool:
            raise DataError(
                "created_by_tool must be true for snapshot restore_token"
            )
        return {
            "v": RESTORE_TOKEN_VERSION,
            "kind": RESTORE_KIND_SNAPSHOT,
            "sheet": sheet.strip(),
            "created_by_tool": True,
        }
    token: dict[str, Any] = {
        "v": RESTORE_TOKEN_VERSION,
        "kind": kind_norm,
        "sheet": sheet.strip(),
        "prior_filters": [dict(f) for f in prior_filters],
        "prior_sort": [dict(s) for s in prior_sort] if prior_sort is not None else None,
        "sort_applied": bool(sort_applied),
        "columns_hidden": [int(c) for c in columns_hidden],
    }
    if kind_norm == RESTORE_KIND_LISTOBJECT:
        if not isinstance(table_name, str) or not table_name.strip():
            raise DataError("table_name is required for listobject restore_token")
        token["table"] = table_name.strip()
    else:
        if not isinstance(range_a1, str) or not range_a1.strip():
            raise DataError("range_a1 is required for range restore_token")
        if prior_autofilter_mode is None:
            raise DataError(
                "prior_autofilter_mode is required for range restore_token"
            )
        token["range"] = normalize_excel_a1_range(range_a1)
        token["prior_autofilter_mode"] = bool(prior_autofilter_mode)
    return token


def parse_restore_token(token: Mapping[str, Any]) -> dict[str, Any]:
    """Validate a restore_token from ``apply_table_view``."""
    if not isinstance(token, Mapping):
        raise DataError("restore_token must be an object")
    if token.get("v") != RESTORE_TOKEN_VERSION:
        raise DataError(f"Unsupported restore_token version: {token.get('v')!r}")
    kind = str(token.get("kind", "")).strip().lower()
    if kind not in _VALID_RESTORE_KINDS:
        raise DataError(
            f"restore_token.kind {token.get('kind')!r} is not supported here"
        )
    sheet = token.get("sheet")
    if not isinstance(sheet, str) or not sheet.strip():
        raise DataError("restore_token.sheet is required")
    if kind == RESTORE_KIND_SNAPSHOT:
        if token.get("created_by_tool") is not True:
            raise DataError(
                "restore_token.created_by_tool must be true to delete a snapshot sheet"
            )
        return {
            "kind": RESTORE_KIND_SNAPSHOT,
            "sheet": sheet.strip(),
            "created_by_tool": True,
        }
    prior_filters = token.get("prior_filters")
    if not isinstance(prior_filters, list):
        raise DataError("restore_token.prior_filters must be a list")
    columns_hidden = token.get("columns_hidden")
    if not isinstance(columns_hidden, list):
        raise DataError("restore_token.columns_hidden must be a list")
    base: dict[str, Any] = {
        "kind": kind,
        "sheet": sheet.strip(),
        "prior_filters": prior_filters,
        "prior_sort": token.get("prior_sort"),
        "sort_applied": bool(token.get("sort_applied")),
        "columns_hidden": [int(c) for c in columns_hidden],
    }
    if kind == RESTORE_KIND_LISTOBJECT:
        table = token.get("table")
        if not isinstance(table, str) or not table.strip():
            raise DataError("restore_token.table is required")
        base["table"] = table.strip()
        return base
    range_a1 = token.get("range")
    if not isinstance(range_a1, str) or not range_a1.strip():
        raise DataError("restore_token.range is required")
    if "prior_autofilter_mode" not in token:
        raise DataError("restore_token.prior_autofilter_mode is required")
    base["range"] = normalize_excel_a1_range(range_a1)
    base["prior_autofilter_mode"] = bool(token.get("prior_autofilter_mode"))
    return base
