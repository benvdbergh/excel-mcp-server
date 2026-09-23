"""Where-clauses, pagination, filter/sort, and shared row-query payloads."""

from typing import Any, Mapping, Optional, Sequence

from excel_mcp.exceptions import DataError

# query_table defaults (documented in TOOLS.md)
DEFAULT_QUERY_TABLE_LIMIT = 100
MAX_QUERY_TABLE_COLUMNS_WITHOUT_EXPLICIT = 32
MAX_QUERY_TABLE_COLUMNS_WITH_OMIT_EMPTY = 128

_VIEWABLE_OPS = frozenset(
    {"eq", "neq", "contains", "in", "gt", "gte", "lt", "lte"}
)
_ALL_QUERY_OPS = _VIEWABLE_OPS | {"is_empty"}
_OPS_REQUIRING_VALUE = frozenset(
    {"eq", "neq", "contains", "in", "gt", "gte", "lt", "lte"}
)


def _try_number(value: Any) -> Optional[float]:
    """Parse a numeric value for comparison; ``bool`` is not treated as numeric."""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None
        try:
            return float(s)
        except ValueError:
            return None
    return None


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def _is_empty_cell(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _values_equal(left: Any, right: Any) -> bool:
    ln = _try_number(left)
    rn = _try_number(right)
    if ln is not None and rn is not None:
        return ln == rn
    return _as_text(left) == _as_text(right)


def _compare_ordered(left: Any, right: Any, op: str) -> bool:
    ln = _try_number(left)
    rn = _try_number(right)
    if ln is not None and rn is not None:
        a, b = ln, rn
    else:
        a, b = _as_text(left), _as_text(right)
    if op == "gt":
        return a > b
    if op == "gte":
        return a >= b
    if op == "lt":
        return a < b
    if op == "lte":
        return a <= b
    raise DataError(f"Unsupported comparison operator: {op!r}")


def evaluate_row_clause(row: Mapping[str, Any], clause: Mapping[str, Any]) -> bool:
    """Apply one ``where`` clause to a row dict keyed by column header."""
    column = clause["column"]
    op = clause["op"]
    cell = row.get(column)
    if op == "is_empty":
        return _is_empty_cell(cell)
    value = clause.get("value")
    if op == "eq":
        return _values_equal(cell, value)
    if op == "neq":
        return not _values_equal(cell, value)
    if op == "contains":
        return _as_text(value).lower() in _as_text(cell).lower()
    if op == "in":
        if not isinstance(value, (list, tuple)):
            raise DataError("Operator 'in' requires a list value")
        return any(_values_equal(cell, item) for item in value)
    if op in ("gt", "gte", "lt", "lte"):
        return _compare_ordered(cell, value, op)
    raise DataError(f"Unsupported where operator: {op!r}")


def normalize_where_clauses(
    where: Optional[Sequence[Mapping[str, Any]]],
    *,
    known_columns: Sequence[str],
) -> list[dict[str, Any]]:
    """Validate and normalize ``where`` clauses; raise ``DataError`` on bad input."""
    known = set(known_columns)
    if where is None:
        return []
    if not isinstance(where, (list, tuple)):
        raise DataError("where must be a list of clause objects")
    out: list[dict[str, Any]] = []
    for i, raw in enumerate(where):
        if not isinstance(raw, Mapping):
            raise DataError(f"Invalid where clause at index {i}: expected object")
        column = raw.get("column")
        op_raw = raw.get("op")
        if not isinstance(column, str) or not column:
            raise DataError(f"Invalid where clause at index {i}: column is required")
        if column not in known:
            raise DataError(f"Unknown column: {column!r}")
        if not isinstance(op_raw, str) or not op_raw:
            raise DataError(f"Invalid where clause at index {i}: op is required")
        op = op_raw.strip().lower()
        if op not in _ALL_QUERY_OPS:
            raise DataError(
                f"Invalid where operator: {op_raw!r}; expected one of "
                f"{', '.join(sorted(_ALL_QUERY_OPS))}"
            )
        clause: dict[str, Any] = {"column": column, "op": op}
        if op in _OPS_REQUIRING_VALUE:
            if "value" not in raw:
                raise DataError(
                    f"Invalid where clause at index {i}: value is required for op {op!r}"
                )
            value = raw["value"]
            if op == "in" and not isinstance(value, (list, tuple)):
                raise DataError(
                    f"Invalid where clause at index {i}: 'in' value must be a list"
                )
            clause["value"] = list(value) if op == "in" else value
        out.append(clause)
    return out


def resolve_query_columns(
    table_columns: Sequence[str],
    columns: Optional[Sequence[str]],
    *,
    omit_empty: bool = False,
) -> list[str]:
    """Resolve projection columns; enforce width threshold when omitted."""
    names = list(table_columns)
    if columns is None:
        if len(names) > MAX_QUERY_TABLE_COLUMNS_WITH_OMIT_EMPTY:
            raise DataError(
                f"Table has {len(names)} columns; pass an explicit columns list "
                f"(omit_empty allows at most "
                f"{MAX_QUERY_TABLE_COLUMNS_WITH_OMIT_EMPTY})"
            )
        if len(names) > MAX_QUERY_TABLE_COLUMNS_WITHOUT_EXPLICIT:
            if omit_empty:
                return names
            raise DataError(
                f"Table has {len(names)} columns; pass an explicit columns list "
                f"(threshold is {MAX_QUERY_TABLE_COLUMNS_WITHOUT_EXPLICIT})"
            )
        return names
    if not isinstance(columns, (list, tuple)):
        raise DataError("columns must be a list of column names")
    known = set(names)
    out: list[str] = []
    for name in columns:
        if not isinstance(name, str) or not name:
            raise DataError(f"Invalid column name: {name!r}")
        if name not in known:
            raise DataError(f"Unknown column: {name!r}")
        if name not in out:
            out.append(name)
    if (
        omit_empty
        and len(out) > MAX_QUERY_TABLE_COLUMNS_WITH_OMIT_EMPTY
    ):
        raise DataError(
            f"Projection has {len(out)} columns; omit_empty allows at most "
            f"{MAX_QUERY_TABLE_COLUMNS_WITH_OMIT_EMPTY}"
        )
    return out


def normalize_query_pagination(
    limit: Optional[int], offset: Optional[int]
) -> tuple[int, int]:
    """Apply default limit and validate offset/limit."""
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
        lim = DEFAULT_QUERY_TABLE_LIMIT
    else:
        try:
            lim = int(limit)
        except (TypeError, ValueError) as e:
            raise DataError(f"Invalid limit: {limit!r}") from e
        if lim < 0:
            raise DataError(f"Invalid limit: {limit!r}; must be >= 0")
    return lim, off


def build_view_applicability(
    clauses: Sequence[Mapping[str, Any]],
    *,
    offset: int,
    truncated: bool,
) -> dict[str, Any]:
    """Describe which parts of the query Excel can later show as a view."""
    clause_apps: list[dict[str, Any]] = []
    for clause in clauses:
        op = str(clause["op"])
        viewable = op in _VIEWABLE_OPS
        entry: dict[str, Any] = {
            "column": clause["column"],
            "op": op,
            "viewable": viewable,
        }
        if not viewable:
            entry["reason"] = (
                f"Operator {op!r} is not expressible as an Excel AutoFilter view"
            )
        clause_apps.append(entry)

    all_clauses_ok = all(c["viewable"] for c in clause_apps)
    page_complete = offset == 0 and not truncated
    overall = all_clauses_ok and page_complete
    result: dict[str, Any] = {
        "viewable": overall,
        "limit_viewable": False,
        "offset_viewable": False,
        "clauses": clause_apps,
    }
    if all_clauses_ok and not page_complete:
        reasons: list[str] = []
        if offset > 0:
            reasons.append("offset drops leading matching rows")
        if truncated:
            reasons.append("limit truncates matching rows")
        result["reason"] = "; ".join(reasons) or "limit/offset changed the result set"
    return result


def sort_table_rows(
    rows: Sequence[Mapping[str, Any]],
    sort: Optional[Mapping[str, Any]],
) -> list[dict[str, Any]]:
    """Stable in-memory sort using a normalized ``view_spec.sort`` (``by`` list)."""
    out = [dict(r) for r in rows]
    if not sort:
        return out
    keys = sort.get("by") if isinstance(sort, Mapping) else None
    if not isinstance(keys, (list, tuple)) or not keys:
        return out

    def _part(value: Any) -> tuple[int, Any]:
        if value is None:
            return (2, "")
        if isinstance(value, bool):
            return (0, int(value))
        if isinstance(value, (int, float)):
            return (0, float(value))
        return (1, str(value).lower())

    for entry in reversed(list(keys)):
        col = str(entry["column"])
        reverse = str(entry.get("order", "asc")).lower() == "desc"
        out.sort(key=lambda row, c=col: _part(row.get(c)), reverse=reverse)
    return out


def filter_table_rows(
    rows: Sequence[Mapping[str, Any]],
    where: Sequence[Mapping[str, Any]],
) -> list[dict[str, Any]]:
    """Shared AND-combined predicate used by file and COM backends."""
    if not where:
        return [dict(r) for r in rows]
    out: list[dict[str, Any]] = []
    for row in rows:
        if all(evaluate_row_clause(row, clause) for clause in where):
            out.append(dict(row))
    return out


def row_matches_search(row: Mapping[str, Any], search: str) -> bool:
    """True when any string cell in ``row`` contains ``search`` (case-insensitive)."""
    needle = search.lower()
    for value in row.values():
        if isinstance(value, str) and needle in value.lower():
            return True
    return False


def filter_rows_by_search(
    rows: Sequence[Mapping[str, Any]],
    search: Optional[str],
) -> list[dict[str, Any]]:
    """Keep rows whose loaded string cells contain ``search``; no-op when unset."""
    if search is None or search == "":
        return [dict(r) for r in rows]
    return [dict(r) for r in rows if row_matches_search(r, search)]


def drop_empty_projected_columns(
    headers: Sequence[str],
    rows: Sequence[Mapping[str, Any]],
) -> tuple[list[str], list[dict[str, Any]]]:
    """Drop projected columns that are empty (None or "") in every returned row."""
    keep = [
        h
        for h in headers
        if any(not _is_empty_cell(row.get(h)) for row in rows)
    ]
    projected = [{h: row.get(h) for h in keep} for row in rows]
    return keep, projected


def build_query_table_payload(
    *,
    headers: Sequence[str],
    matching_rows: Sequence[Mapping[str, Any]],
    where: Sequence[Mapping[str, Any]],
    limit: int,
    offset: int,
    target: Mapping[str, Any],
    omit_empty: bool = False,
) -> dict[str, Any]:
    """Page matching rows and attach ``view_spec`` / ``view_applicability``."""
    total = len(matching_rows)
    page = list(matching_rows[offset : offset + limit])
    truncated = (offset + len(page)) < total
    projected = [
        {h: row.get(h) for h in headers} for row in page
    ]
    headers_out = list(headers)
    if omit_empty:
        headers_out, projected = drop_empty_projected_columns(headers_out, projected)
    # Keep every clause, including ones Excel cannot show. Epic-14 must refuse
    # a partial view when ``view_applicability.viewable`` is false.
    view_spec = {
        "target": dict(target),
        "columns": list(headers_out),
        "where": [dict(c) for c in where],
        "sort": None,
    }
    return {
        "headers": list(headers_out),
        "rows": projected,
        "row_count": len(projected),
        "truncated": truncated,
        "view_spec": view_spec,
        "view_applicability": build_view_applicability(
            where, offset=offset, truncated=truncated
        ),
    }


def query_table_rows(
    *,
    table_name: str,
    table_columns: Sequence[str],
    rows: Sequence[Mapping[str, Any]],
    columns: Optional[Sequence[str]] = None,
    where: Optional[Sequence[Mapping[str, Any]]] = None,
    limit: Optional[int] = None,
    offset: Optional[int] = None,
    omit_empty: bool = False,
    search: Optional[str] = None,
) -> dict[str, Any]:
    """Validate args, filter, page, and build the ``query_table`` payload."""
    headers = resolve_query_columns(
        table_columns, columns, omit_empty=omit_empty
    )
    clauses = normalize_where_clauses(where, known_columns=table_columns)
    lim, off = normalize_query_pagination(limit, offset)
    # Filter may reference columns outside the projection; rows must carry them.
    needed = set(headers) | {c["column"] for c in clauses}
    missing = needed - set(table_columns)
    if missing:
        raise DataError(f"Unknown column: {sorted(missing)[0]!r}")
    matching = filter_table_rows(rows, clauses)
    if search:
        matching = filter_rows_by_search(matching, search)
    return build_query_table_payload(
        headers=headers,
        matching_rows=matching,
        where=clauses,
        limit=lim,
        offset=off,
        target={"kind": "table", "name": table_name},
        omit_empty=omit_empty,
    )


def query_region_rows(
    *,
    sheet: str,
    range_a1: str,
    region_columns: Sequence[str],
    rows: Sequence[Mapping[str, Any]],
    columns: Optional[Sequence[str]] = None,
    where: Optional[Sequence[Mapping[str, Any]]] = None,
    limit: Optional[int] = None,
    offset: Optional[int] = None,
    omit_empty: bool = False,
    search: Optional[str] = None,
) -> dict[str, Any]:
    """Validate args, filter, page, and build the ``query_region`` payload."""
    # Lazy import: layout imports cell helpers from this module.
    from .layout import normalize_excel_a1_range

    headers = resolve_query_columns(
        region_columns, columns, omit_empty=omit_empty
    )
    clauses = normalize_where_clauses(where, known_columns=region_columns)
    lim, off = normalize_query_pagination(limit, offset)
    needed = set(headers) | {c["column"] for c in clauses}
    missing = needed - set(region_columns)
    if missing:
        raise DataError(f"Unknown column: {sorted(missing)[0]!r}")
    matching = filter_table_rows(rows, clauses)
    if search:
        matching = filter_rows_by_search(matching, search)
    return build_query_table_payload(
        headers=headers,
        matching_rows=matching,
        where=clauses,
        limit=lim,
        offset=off,
        target={
            "kind": "region",
            "sheet": sheet,
            "range": normalize_excel_a1_range(range_a1),
        },
        omit_empty=omit_empty,
    )
