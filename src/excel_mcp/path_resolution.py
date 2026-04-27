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

Cloud workbook locators (Story 9-1 / ADR 0006)
----------------------------------------------

A **cloud workbook locator** is an ``https:`` URL treated as an opaque workbook
identity for COM routing when the workbook is open from SharePoint / Microsoft 365.
Validation and canonical form are implemented by
:func:`parse_cloud_workbook_locator`; :func:`resolve_target` must only be used
for filesystem paths (unchanged).

Interaction with ``get_excel_path`` (stdio vs SSE)
---------------------------------------------------

* **stdio** (``EXCEL_FILES_PATH`` unset): callers may pass a validated ``https:``
  workbook locator (see ``parse_cloud_workbook_locator``) or a disk path.
  When ``EXCEL_MCP_ALLOWED_PATHS`` is unset/empty, ``get_excel_path`` returns the
  canonical cloud locator for HTTPS targets without ``os.path.realpath``.
  When the path allowlist is active, https locators must satisfy
  ``EXCEL_MCP_ALLOWED_URL_PREFIXES`` (see ``excel_mcp.path_policy``).
* **SSE / HTTP jail** (``EXCEL_FILES_PATH`` set): cloud workbook URLs are not
  supported (``ValueError``); use local paths under the jail root.
"""

from __future__ import annotations

import os
import re
from urllib.parse import quote, unquote, urlparse, urlunparse

# v1: only https workbook locators (ADR 0006).
_ALLOWED_CLOUD_SCHEMES = frozenset({"https"})


def is_cloud_workbook_locator(s: str) -> bool:
    """Return True if ``s`` is a syntactically valid v1 cloud workbook locator."""
    try:
        parse_cloud_workbook_locator(s)
        return True
    except ValueError:
        return False


def parse_cloud_workbook_locator(locator: str) -> str:
    """Parse and return canonical cloud workbook locator string for downstream use.

    Rules (v1): ``https`` scheme only; no NUL; non-empty host via
    :func:`urllib.parse.urlparse`; scheme allowlist; reject obviously malformed
    inputs. Canonicalization: lowercases scheme and netloc (host/port identity),
    normalizes path with decode/re-encode and leading slash, preserves query and
    fragment.

    Raises:
        ValueError: If ``locator`` is not a valid cloud workbook locator.
    """
    if not locator or "\x00" in locator:
        raise ValueError("Invalid cloud workbook locator: empty or contains NUL character")
    s = locator.strip()
    if not s:
        raise ValueError("Invalid cloud workbook locator: empty")

    parsed = urlparse(s)
    scheme = parsed.scheme.lower()
    if scheme not in _ALLOWED_CLOUD_SCHEMES:
        raise ValueError(
            f"Invalid cloud workbook locator: only https is supported (v1); got scheme {parsed.scheme!r}"
        )

    if not parsed.netloc or not parsed.netloc.strip():
        raise ValueError("Invalid cloud workbook locator: missing host (netloc)")

    host = parsed.hostname
    if host is None or not str(host).strip():
        raise ValueError("Invalid cloud workbook locator: missing or invalid host")

    netloc = parsed.netloc.lower()
    if re.search(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", netloc):
        raise ValueError("Invalid cloud workbook locator: control characters in netloc")

    path_decoded = unquote(parsed.path or "")
    if not path_decoded:
        norm_path = "/"
    else:
        if not path_decoded.startswith("/"):
            path_decoded = "/" + path_decoded
        # Slash consistency + percent-encoding for non-ASCII / spaces.
        norm_path = quote(path_decoded, safe="/")

    return urlunparse(("https", netloc, norm_path, parsed.params, parsed.query, parsed.fragment))


def _norm_disk_path_for_com(path: str) -> str:
    """Canonical disk path for COM identity comparison (matches COM service disk norm).

    Kept here (not imported from COM modules) so :func:`normalize_workbook_target_for_com`
    stays usable without circular imports.
    """
    expanded = os.path.expanduser(path)
    try:
        canonical = os.path.realpath(expanded)
    except OSError:
        canonical = os.path.abspath(expanded)
    return os.path.normcase(os.path.normpath(canonical))


def normalize_workbook_target_for_com(path: str) -> str:
    """Normalize a workbook target string for COM comparison with ``Workbook.FullName``.

    * **Https cloud locators** — canonical form from :func:`parse_cloud_workbook_locator`.
    * **Local paths** — same normalization as disk-based COM matching (``realpath``,
      ``normcase`` / ``normpath``), aligned with ``resolve_target`` for absolute paths.

    Raises:
        ValueError: Empty path, NUL, or invalid cloud locator (same as parse helper).
    """
    if not path or "\x00" in path:
        raise ValueError(f"Invalid workbook target: empty or contains NUL character")
    if is_cloud_workbook_locator(path):
        return parse_cloud_workbook_locator(path)
    return _norm_disk_path_for_com(path)


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
