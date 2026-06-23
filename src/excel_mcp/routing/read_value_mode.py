"""Validation for ``read_range_with_metadata`` / ``read_data_from_excel`` parameters."""

from __future__ import annotations

VALID_VALUE_MODES: frozenset[str] = frozenset({"value", "text"})
VALID_METADATA_MODES: frozenset[str] = frozenset({"full", "compact"})


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
