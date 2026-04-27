"""Operator environment: workbook transport and COM strictness (ADR 0001, ADR 0005).

Variables here select **workbook** execution (file-backed vs COM automation) and COM
policy. They do **not** configure the MCP client↔server wire transport (stdio, SSE,
or streamable HTTP); see ADR 0001 for that vocabulary split.
"""

from __future__ import annotations

import os
from typing import Mapping, cast

from excel_mcp.routing.routing_backend import WorkbookTransport

EXCEL_MCP_TRANSPORT = "EXCEL_MCP_TRANSPORT"
EXCEL_MCP_COM_STRICT = "EXCEL_MCP_COM_STRICT"
EXCEL_MCP_COM_ALLOW_FILE_FALLBACK = "EXCEL_MCP_COM_ALLOW_FILE_FALLBACK"

_VALID_WORKBOOK_TRANSPORTS: frozenset[str] = frozenset({"auto", "file", "com"})
_TRUTHY = frozenset({"1", "true", "yes"})
_FALSY = frozenset({"0", "false", "no"})


def read_workbook_transport(
    environ: Mapping[str, str] | None = None,
) -> WorkbookTransport:
    """Read ``EXCEL_MCP_TRANSPORT``: ``auto`` | ``file`` | ``com`` (case-insensitive).

    Default ``auto`` when the variable is unset or empty/whitespace.

    Raises:
        ValueError: If the value is not one of the allowed workbook transport modes.
            The message distinguishes **workbook** transport from MCP **wire**
            transport (ADR 0001).
    """
    env = os.environ if environ is None else environ
    raw = env.get(EXCEL_MCP_TRANSPORT, "")
    normalized = raw.strip().lower()
    if not normalized:
        return "auto"
    if normalized not in _VALID_WORKBOOK_TRANSPORTS:
        raise ValueError(
            f"Invalid {EXCEL_MCP_TRANSPORT}={raw!r}: expected one of "
            f"'auto', 'file', 'com' (workbook file/COM routing per ADR 0001). "
            f"This is not the MCP host wire transport (stdio/SSE/HTTP)."
        )
    return cast(WorkbookTransport, normalized)


def read_com_strict(environ: Mapping[str, str] | None = None) -> bool:
    """Read ``EXCEL_MCP_COM_STRICT``.

    Truthy for ``1``, ``true``, ``yes`` (case-insensitive). Falsy for ``0``,
    ``false``, ``no``, unset, or empty/whitespace.

    When unset or empty, returns ``True`` (strict COM policy by default; ADR 0005).

    Raises:
        ValueError: If set to a non-empty value that is neither truthy nor falsy.
    """
    env = os.environ if environ is None else environ
    raw = env.get(EXCEL_MCP_COM_STRICT)
    if raw is None:
        return True
    token = raw.strip().lower()
    if not token:
        return True
    if token in _TRUTHY:
        return True
    if token in _FALSY:
        return False
    raise ValueError(
        f"Invalid {EXCEL_MCP_COM_STRICT}={raw!r}: use '1'/'true'/'yes' or "
        f"'0'/'false'/'no' (case-insensitive), or leave unset for strict default."
    )


def read_com_allow_file_fallback(environ: Mapping[str, str] | None = None) -> bool:
    """Read ``EXCEL_MCP_COM_ALLOW_FILE_FALLBACK``.

    Truthy only for ``1``, ``true``, ``yes`` (case-insensitive). Falsy when unset,
    empty/whitespace, or ``0`` / ``false`` / ``no``.

    Raises:
        ValueError: If set to a non-empty value that is neither truthy nor falsy.
    """
    env = os.environ if environ is None else environ
    raw = env.get(EXCEL_MCP_COM_ALLOW_FILE_FALLBACK)
    if raw is None:
        return False
    token = raw.strip().lower()
    if not token:
        return False
    if token in _TRUTHY:
        return True
    if token in _FALSY:
        return False
    raise ValueError(
        f"Invalid {EXCEL_MCP_COM_ALLOW_FILE_FALLBACK}={raw!r}: use '1'/'true'/'yes' "
        f"or '0'/'false'/'no' (case-insensitive), or leave unset."
    )


def resolve_workbook_transport(
    override: str | None,
    environ: Mapping[str, str] | None = None,
) -> WorkbookTransport:
    """Per-call ``workbook_transport`` or env default (``EXCEL_MCP_TRANSPORT``).

    ``None`` or whitespace-only ``override`` selects :func:`read_workbook_transport`.
    Otherwise ``override`` must be ``auto`` | ``file`` | ``com`` (case-insensitive).
    """
    if override is None:
        return read_workbook_transport(environ)
    token = override.strip().lower()
    if not token:
        return read_workbook_transport(environ)
    if token not in _VALID_WORKBOOK_TRANSPORTS:
        raise ValueError(
            f"Invalid workbook_transport={override!r}: expected 'auto', 'file', or "
            f"'com' (workbook routing per ADR 0001; not MCP wire transport)."
        )
    return cast(WorkbookTransport, token)


def effective_com_strict(environ: Mapping[str, str] | None = None) -> bool:
    """Effective COM strict flag for routing (combines strict + file-fallback allow).

    Returns ``False`` when either:

    * ``EXCEL_MCP_COM_ALLOW_FILE_FALLBACK`` is truthy (operator allows documented
      file fallback where non-strict routing would apply), **or**
    * ``EXCEL_MCP_COM_STRICT`` is explicitly falsy (``0`` / ``false`` / ``no``).

    Otherwise returns ``True``. Precedence: allowing file fallback forces
    non-strict *effective* behavior for the router whenever that flag is on;
    explicit non-strict (``read_com_strict`` false) also yields non-strict
    regardless of the fallback flag.

    Unset ``EXCEL_MCP_COM_STRICT`` still defaults to strict ``True``; unset
    fallback defaults to ``False``, so the default effective result is ``True``.

    Reads ``os.environ`` at call time (no process-wide cache).
    """
    return read_com_strict(environ) and not read_com_allow_file_fallback(environ)
