"""Shared workbook host identity for COM open detection and attach (ADR 0006).

Excel ``Workbook.FullName`` for SharePoint / Microsoft 365 workbooks is often an
``https://`` URL, not a local synced path. A caller's local OneDrive/SharePoint
sync path is **not** equivalent to Excel's https ``FullName`` for matching — use
``excel_list_open_workbooks`` (ADR 0009) to obtain the exact COM locator.
"""

from __future__ import annotations

import numbers
import os
from typing import Any
from urllib.parse import urljoin

from excel_mcp.path_resolution import normalize_workbook_target_for_com


def _coerce_workbook_count(val: Any) -> int:
    if isinstance(val, bool):
        return 0
    if isinstance(val, numbers.Integral):
        return int(val)
    return 0


def normalized_workbook_fullname(wb: Any) -> str | None:
    """Return normalized ``Workbook.FullName`` or ``None`` if unusable."""
    try:
        return normalize_workbook_target_for_com(str(wb.FullName))
    except Exception:
        return None


def protected_view_candidate_paths(pv: Any) -> list[str]:
    """Normalized identities for a ``ProtectedViewWindow`` (COM object)."""
    out: list[str] = []
    try:
        wb = pv.Workbook
        fn = normalized_workbook_fullname(wb)
        if fn:
            out.append(fn)
    except Exception:
        pass
    try:
        sp, sn = str(pv.SourcePath), str(pv.SourceName)
        if sp and sn:
            sps = sp.strip()
            if sps.lower().startswith("https://"):
                base = sp if sp.endswith("/") else sp + "/"
                combined = urljoin(
                    base, str(sn).replace("\\", "/").lstrip("/")
                )
                out.append(normalize_workbook_target_for_com(combined))
            else:
                out.append(
                    normalize_workbook_target_for_com(os.path.join(sp, sn))
                )
        elif sp:
            out.append(normalize_workbook_target_for_com(sp))
    except Exception:
        pass
    return out


def count_workbook_collection_matches(xl: Any, target: str) -> int:
    """Count ``Application.Workbooks`` entries whose identity equals ``target``.

    Protected View windows are intentionally excluded (FR-9): PV-held workbooks
    are not members of ``Workbooks`` until the user clicks **Enable Editing**.
    """
    try:
        n = _coerce_workbook_count(getattr(xl.Workbooks, "Count", 0))
    except Exception:
        n = 0
    matches = 0
    for i in range(1, n + 1):
        try:
            wb = xl.Workbooks.Item(i)
        except Exception:
            continue
        norm_full = normalized_workbook_fullname(wb)
        if norm_full == target:
            matches += 1
    return matches


def protected_view_matches_target(xl: Any, target: str) -> bool:
    """True when any Protected View window identity equals ``target``."""
    try:
        pvw = xl.ProtectedViewWindows
        n_pv = _coerce_workbook_count(getattr(pvw, "Count", 0))
    except Exception:
        return False
    for i in range(1, n_pv + 1):
        try:
            pv = pvw.Item(i)
            for cand in protected_view_candidate_paths(pv):
                if cand == target:
                    return True
        except Exception:
            continue
    return False


def workbook_in_protected_view(xl: Any, wb: Any) -> bool:
    """True if ``wb`` is the workbook shown in a ``ProtectedViewWindow``."""
    want = normalized_workbook_fullname(wb)
    if not want:
        return False
    try:
        pvw = xl.ProtectedViewWindows
        n = _coerce_workbook_count(getattr(pvw, "Count", 0))
    except Exception:
        return False
    for i in range(1, n + 1):
        try:
            pv = pvw.Item(i)
            pw = pv.Workbook
            got = normalized_workbook_fullname(pw)
            if got and got == want:
                return True
        except Exception:
            continue
    return False


__all__ = [
    "count_workbook_collection_matches",
    "normalized_workbook_fullname",
    "protected_view_candidate_paths",
    "protected_view_matches_target",
    "workbook_in_protected_view",
]
