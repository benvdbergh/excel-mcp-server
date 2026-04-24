"""Single entry for workbook target path normalization (FR-1).

Relative path resolution order
-------------------------------

When ``path`` is relative:

1. If ``search_roots`` is a non-empty tuple: for each ``root`` in order, form
   ``candidate = os.path.join(root, path)``. The first ``candidate`` for which
   ``os.path.isfile(os.path.expanduser(candidate))`` is true is returned as
   ``os.path.realpath(os.path.expanduser(candidate))``.
2. Otherwise, or if no root yields an existing file: resolve against
   ``cwd`` when provided, else ``os.getcwd()`` — return
   ``os.path.realpath(os.path.expanduser(os.path.join(base, path)))`` (the path
   need not exist yet; used for sandbox joins and new-file flows).

When ``path`` is absolute: return ``os.path.realpath(os.path.expanduser(path))``.

Empty ``path`` or a NUL character (``"\\x00"``) in ``path`` raises ``ValueError``
(aligned with ``get_excel_path``).

Interaction with ``get_excel_path`` (stdio vs SSE)
---------------------------------------------------

* **stdio** (``EXCEL_FILES_PATH`` unset): callers must pass an absolute path.
  When ``EXCEL_MCP_ALLOWED_PATHS`` is unset/empty, ``get_excel_path`` still
  returns ``os.path.normpath`` only (legacy). When the allowlist is active,
  ``get_excel_path`` uses ``resolve_target`` then ``path_policy.assert_path_allowed``
  so policy matches file/COM normalization (FR-11).
* **SSE / HTTP jail** (``EXCEL_FILES_PATH`` set): relative paths are finalized
  via ``resolve_target(..., cwd=realpath(EXCEL_FILES_PATH))`` with no
  ``search_roots``, matching prior ``realpath(join(base, path))`` behavior.
  If ``EXCEL_MCP_ALLOWED_PATHS`` is set, the result must also lie inside one of
  those roots (intersection with the jail); see ``excel_mcp.path_policy``.
"""

from __future__ import annotations

import os


def resolve_target(
    path: str,
    *,
    cwd: str | None = None,
    search_roots: tuple[str, ...] | None = None,
) -> str:
    if not path or "\x00" in path:
        raise ValueError(f"Invalid path: {path!r}")

    if os.path.isabs(path):
        return os.path.realpath(os.path.expanduser(path))

    if search_roots:
        for root in search_roots:
            candidate = os.path.join(root, path)
            expanded = os.path.expanduser(candidate)
            if os.path.isfile(expanded):
                return os.path.realpath(expanded)

    base = cwd if cwd is not None else os.getcwd()
    joined = os.path.join(base, path)
    return os.path.realpath(os.path.expanduser(joined))
