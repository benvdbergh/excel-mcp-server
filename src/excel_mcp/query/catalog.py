"""Table catalog builders and list-tables detail validation (openpyxl-free)."""

from typing import Any, Mapping, Optional

from excel_mcp.exceptions import DataError

from .layout import normalize_excel_a1_range

DETAIL_MINIMAL = "minimal"
DETAIL_SCHEMA = "schema"
DEFAULT_LIST_TABLES_DETAIL = DETAIL_SCHEMA
_VALID_LIST_TABLES_DETAIL = frozenset({DETAIL_MINIMAL, DETAIL_SCHEMA})


def validate_list_tables_detail(detail: str) -> str:
    """Return normalized detail mode or raise ``DataError``."""
    mode = (detail or DEFAULT_LIST_TABLES_DETAIL).strip().lower()
    if mode not in _VALID_LIST_TABLES_DETAIL:
        raise DataError(
            f"Invalid detail: {detail!r}; expected '{DETAIL_MINIMAL}' or '{DETAIL_SCHEMA}'"
        )
    return mode


def build_table_catalog_entry(
    *,
    sheet: str,
    name: str,
    range_: str,
    header_range: str,
    data_range: Optional[str],
    row_count: int,
    filter_applied: bool,
    columns: Optional[list[str]] = None,
) -> dict[str, Any]:
    """One ListObject catalog row (no cell values)."""
    entry: dict[str, Any] = {
        "sheet": sheet,
        "name": name,
        "range": normalize_excel_a1_range(range_),
        "header_range": normalize_excel_a1_range(header_range),
        "data_range": (
            normalize_excel_a1_range(data_range) if data_range else None
        ),
        "row_count": int(row_count),
        "filter_applied": bool(filter_applied),
    }
    if columns is not None:
        entry["columns"] = list(columns)
    return entry


def build_table_catalog_payload(tables: list[Mapping[str, Any]]) -> dict[str, Any]:
    """Stable workbook-scoped catalog envelope."""
    return {"tables": list(tables)}
