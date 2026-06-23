"""Validation for ``read_range_with_metadata`` / ``read_data_from_excel`` value_mode."""

from __future__ import annotations

VALID_VALUE_MODES: frozenset[str] = frozenset({"value", "text"})


def validate_value_mode(value_mode: str) -> str:
    """Return ``value_mode`` if valid; raise ``ValueError`` with an actionable message."""
    if value_mode not in VALID_VALUE_MODES:
        allowed = ", ".join(f"'{m}'" for m in sorted(VALID_VALUE_MODES))
        raise ValueError(f"Invalid value_mode {value_mode!r}; expected one of: {allowed}")
    return value_mode
