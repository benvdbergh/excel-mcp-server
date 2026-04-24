"""Path allowlist and containment policy (FR-11).

``EXCEL_MCP_ALLOWED_PATHS``
---------------------------
Optional. When **unset** or **whitespace-only**, the allowlist is **inactive**
(default): stdio keeps legacy behavior (any absolute path after
``os.path.normpath`` in ``get_excel_path`` unless allowlist is enabled separately);
SSE/HTTP uses only the ``EXCEL_FILES_PATH`` jail.

When set to a non-empty value, entries are separated by ``os.pathsep`` (same
as ``PATH``: semicolon on Windows, colon on POSIX). Each non-empty segment is
trimmed, passed through ``os.path.expanduser``, then ``os.path.realpath`` to
form a canonical directory root. Empty segments are ignored.

**stdio** (``EXCEL_FILES_PATH`` unset): after the workbook path is resolved with
``resolve_target``, it must lie **inside** at least one allowed root (same
containment rule as the SSE jail: ``resolved_path_is_within``).

**SSE / streamable HTTP** (``EXCEL_FILES_PATH`` set): the resolved path must be
inside ``realpath(EXCEL_FILES_PATH)`` **and**, when the allowlist is active,
inside at least one allowed root (**intersection**). If the allowlist is
inactive, behavior is jail-only (unchanged).

Invalid or unreadable root entries are skipped; if none remain, the allowlist
is effectively **on** with **no** roots and all paths fail the allowlist check
(fail-closed).
"""

from __future__ import annotations

import os


def resolved_path_is_within(base: str, candidate: str) -> bool:
    """True if ``candidate`` is the same path as ``base`` or strictly inside it.

    Both paths are canonicalized with ``os.path.realpath``. Uses
    ``os.path.commonpath`` for containment (same semantics as the historical
    ``EXCEL_FILES_PATH`` jail in ``server.py``).
    """
    base_rp = os.path.realpath(base)
    candidate_rp = os.path.realpath(candidate)
    if candidate_rp == base_rp:
        return True
    try:
        return os.path.commonpath([base_rp, candidate_rp]) == base_rp
    except ValueError:
        return False


def _allowlist_roots() -> tuple[str, ...] | None:
    """``None`` if allowlist env is unset/empty; else tuple of ``realpath`` roots (may be empty)."""
    raw = os.environ.get("EXCEL_MCP_ALLOWED_PATHS")
    if raw is None or not raw.strip():
        return None
    parts = [p.strip() for p in raw.split(os.pathsep) if p.strip()]
    if not parts:
        return None
    roots: list[str] = []
    for p in parts:
        expanded = os.path.expanduser(p)
        try:
            roots.append(os.path.realpath(expanded))
        except OSError:
            continue
    return tuple(roots)


def allowlist_enforced() -> bool:
    """True when allowlist rules apply (env set with at least one path segment after ``os.pathsep`` split).

    Includes the case where every root failed to resolve (empty tuple): paths still
    must go through ``resolve_target`` + ``assert_path_allowed`` and will be rejected.
    """
    return _allowlist_roots() is not None


def path_is_allowed(resolved: str, *, jail_realpath: str | None = None) -> bool:
    """Return whether a **resolved** absolute path passes jail + optional allowlist.

    ``resolved`` should already be normalized (e.g. output of ``resolve_target``).

    * If ``jail_realpath`` is set (SSE/HTTP), the path must be inside the jail.
    * If an allowlist is active (non-``None`` from internal parsing: env set
      with at least one segment after split), the path must be inside at least
      one allowlist root. An empty tuple of roots (all entries invalid) denies
      all paths for the allowlist portion.
    * If no allowlist is configured (env unset/empty), only the jail rule
      applies when ``jail_realpath`` is set; when it is ``None`` (stdio), returns
      ``True`` (caller should not rely on this for unvalidated paths—use
      ``get_excel_path``).
    """
    roots = _allowlist_roots()

    if jail_realpath is not None:
        if not resolved_path_is_within(jail_realpath, resolved):
            return False

    if roots is None:
        return True

    if not roots:
        return False

    for root in roots:
        if resolved_path_is_within(root, resolved):
            return True
    return False


def assert_path_allowed(resolved: str, *, jail_realpath: str | None = None) -> None:
    """Raise ``ValueError`` if ``path_is_allowed`` would return ``False``."""
    if path_is_allowed(resolved, jail_realpath=jail_realpath):
        return
    roots = _allowlist_roots()
    if jail_realpath is not None and not resolved_path_is_within(jail_realpath, resolved):
        raise ValueError(
            f"Invalid path: {resolved!r} escapes EXCEL_FILES_PATH jail "
            f"({jail_realpath!r})"
        )
    if roots is not None:
        raise ValueError(
            f"Invalid path: {resolved!r} is not within any directory listed in "
            "EXCEL_MCP_ALLOWED_PATHS"
        )
    raise ValueError(f"Invalid path: {resolved!r}")
