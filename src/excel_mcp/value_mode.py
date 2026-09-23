"""Validation for read value/metadata modes and file-backend formula warnings."""

from __future__ import annotations

from typing import Dict

VALID_VALUE_MODES: frozenset[str] = frozenset({"value", "text"})
VALID_METADATA_MODES: frozenset[str] = frozenset({"full", "compact"})

FILE_BACKEND_FORMULA_NOT_EVALUATED_CODE = "file_backend_formula_not_evaluated"
FILE_BACKEND_FORMULA_NOT_EVALUATED_MESSAGE = (
    "The file (openpyxl) backend does not evaluate Excel formulas; "
    "cached values may be missing (null). Prefer COM routing "
    "(workbook_transport=auto or com) when the workbook is open in Excel."
)


def validate_value_mode(value_mode: str) -> str:
    """Return ``value_mode`` if valid; raise ``ValueError`` with an actionable message."""
    if value_mode not in VALID_VALUE_MODES:
        allowed = ", ".join(f"'{m}'" for m in sorted(VALID_VALUE_MODES))
        raise ValueError(f"Invalid value_mode {value_mode!r}; expected one of: {allowed}")
    return value_mode


def validate_metadata_mode(metadata_mode: str) -> str:
    """Return ``metadata_mode`` if valid; raise ``ValueError`` with an actionable message."""
    if metadata_mode not in VALID_METADATA_MODES:
        allowed = ", ".join(f"'{m}'" for m in sorted(VALID_METADATA_MODES))
        raise ValueError(
            f"Invalid metadata_mode {metadata_mode!r}; expected one of: {allowed}"
        )
    return metadata_mode


def file_backend_formula_not_evaluated_warning() -> Dict[str, str]:
    """ADR 0010 warning when openpyxl reads formula cells without evaluation."""
    return {
        "code": FILE_BACKEND_FORMULA_NOT_EVALUATED_CODE,
        "message": FILE_BACKEND_FORMULA_NOT_EVALUATED_MESSAGE,
    }
