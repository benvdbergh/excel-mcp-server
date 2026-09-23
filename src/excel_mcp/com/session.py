"""COM workbook session helpers: match, open, close, list-open (FR-10)."""

from __future__ import annotations

import json
import numbers
import os
from typing import Any, Dict, List, Optional, Tuple

from excel_mcp.path.resolution import normalize_workbook_target_for_com

# Stable routing errors (FR-9 / ADR 0005) — wording fixed for automation clients.
_ERR_COM_MULTIPLE_MATCH = (
    "Error: Multiple Excel workbooks match this path; close duplicates or use a single instance."
)

_ERR_COM_UNSAVED_PATH = (
    "Error: Workbook has no saved path on disk; save the workbook to a known path before COM routing."
)

_ERR_COM_NOT_OPEN = (
    "Error: Workbook not open in Excel (path does not match any open workbook FullName)"
)

_ERR_COM_PROTECTED_VIEW = (
    "Error: Workbook is open in Protected View in Excel; enable editing before COM routing."
)

_ERR_COM_READ_ONLY = (
    "Error: Workbook is read-only in Excel; COM routing cannot modify this workbook."
)

def _norm_workbook_path(path: str) -> str:
    """Canonical disk path for comparison with Excel ``Workbook.FullName`` (FR-1).

    Uses :func:`os.path.realpath` so junctions/symlinks align with the host path
    Excel reports, while staying consistent with ``resolve_target`` for absolute
    paths when the allowlist path is active.
    """
    expanded = os.path.expanduser(path)
    try:
        canonical = os.path.realpath(expanded)
    except OSError:
        canonical = os.path.abspath(expanded)
    return os.path.normcase(os.path.normpath(canonical))

def _coerce_com_count(val: Any, default: int = 0) -> int:
    """Coerce Excel COM ``Count`` to ``int``; reject non-numeric test doubles (e.g. ``MagicMock``)."""
    if isinstance(val, bool):
        return default
    if isinstance(val, numbers.Integral):
        return int(val)
    return default

def _com_bool_is_true(val: Any) -> bool:
    """True only for a real COM ``VARIANT_BOOL`` / Python ``bool`` (mocks are ignored)."""
    return isinstance(val, bool) and val

def _workbook_fullname_norm(wb: Any) -> Optional[str]:
    from excel_mcp.routing.workbook_host_identity import normalized_workbook_fullname

    return normalized_workbook_fullname(wb)

_LIST_OPEN_WORKBOOKS_DETAIL_LEVELS = frozenset({"minimal", "active_context"})

def _normalize_list_open_workbooks_detail(
    detail: Optional[str],
) -> Tuple[str, Optional[str]]:
    """Return ``(level, error_message)``; error_message is set when invalid."""
    if detail is None:
        return "minimal", None
    level = detail.strip().lower()
    if level not in _LIST_OPEN_WORKBOOKS_DETAIL_LEVELS:
        return (
            "",
            f"Error: Invalid detail {detail!r}; use 'minimal' or 'active_context'.",
        )
    return level, None

def _com_active_context_fields(xl: Any) -> Dict[str, Any]:
    """Active workbook, sheet name, and selection address from ``Excel.Application``."""
    active_workbook: Optional[Dict[str, str]] = None
    active_sheet: Optional[str] = None
    selection: Optional[str] = None

    try:
        aw = xl.ActiveWorkbook
        if aw is not None:
            active_workbook = {
                "full_name": str(aw.FullName),
                "name": str(aw.Name),
            }
    except Exception:
        pass

    try:
        sheet = xl.ActiveSheet
        if sheet is not None:
            active_sheet = str(sheet.Name)
    except Exception:
        pass

    try:
        sel = xl.Selection
        if sel is not None:
            addr = str(getattr(sel, "Address", "")).replace("$", "")
            if addr:
                selection = addr
    except Exception:
        pass

    return {
        "active_workbook": active_workbook,
        "active_sheet": active_sheet,
        "selection": selection,
    }

def collect_workbooks_matching_path(xl: Any, target: str) -> List[Any]:
    """All open COM workbooks whose on-disk path equals ``target`` (normalized).

    Includes workbooks open only in **Protected View** — those are not members of
    ``Application.Workbooks``; match via ``Application.ProtectedViewWindows`` and
    ``ProtectedViewWindow.Workbook`` / ``SourcePath`` + ``SourceName`` (Excel COM).
    """
    from excel_mcp.routing.workbook_host_identity import protected_view_candidate_paths

    matches: List[Any] = []
    try:
        count = _coerce_com_count(getattr(xl.Workbooks, "Count", 0))
    except Exception:
        count = 0
    for i in range(1, count + 1):
        try:
            wb = xl.Workbooks.Item(i)
            if _workbook_fullname_norm(wb) == target:
                matches.append(wb)
        except Exception:
            continue

    try:
        pvw = xl.ProtectedViewWindows
        n_pv = _coerce_com_count(getattr(pvw, "Count", 0))
    except Exception:
        n_pv = 0
    for i in range(1, n_pv + 1):
        try:
            pv = pvw.Item(i)
            for cand in protected_view_candidate_paths(pv):
                if cand == target:
                    matches.append(pv.Workbook)
                    break
        except Exception:
            continue

    # Same COM instance must not be counted twice if ever visible from both paths.
    seen: set[int] = set()
    unique: List[Any] = []
    for wb in matches:
        k = id(wb)
        if k not in seen:
            seen.add(k)
            unique.append(wb)
    return unique

def get_open_workbook_com(filepath: str, *, for_write: bool = False) -> Tuple[Any, Optional[str]]:
    """Return ``(workbook_com, None)`` or ``(None, error_message)``.

    Never-saved workbooks (e.g. ``Book1``) expose a **non-disk** ``FullName`` and an
    empty ``Workbook.Path``; they cannot be matched to a caller-supplied absolute
    file path. When exactly one such workbook is open and nothing matches ``target``,
    return :data:`_ERR_COM_UNSAVED_PATH` instead of a generic not-open message.
    """
    from excel_mcp.routing.workbook_host_identity import workbook_in_protected_view

    # Linux CI: no pywin32. Tests may inject a fake win32com via ``sys.modules``; only
    # short-circuit on import failure, not on ``sys.platform`` (test_read_class_com_wiring
    # fakes win32 routing but has no win32com mock).
    try:
        import win32com.client  # lazy: worker thread only
    except ModuleNotFoundError:
        return None, _ERR_COM_NOT_OPEN

    target = normalize_workbook_target_for_com(filepath)
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return None, "Error: No running Excel application found"

    matches = collect_workbooks_matching_path(xl, target)
    if len(matches) > 1:
        return None, _ERR_COM_MULTIPLE_MATCH

    if len(matches) == 0:
        # Fail-closed: only infer "unsaved path" when a single workbook is open and
        # it has no saved directory (``Path`` is empty per Excel COM).
        try:
            wb_count = _coerce_com_count(getattr(xl.Workbooks, "Count", 0))
        except Exception:
            wb_count = 0
        never_saved_open = 0
        for i in range(1, wb_count + 1):
            try:
                wb = xl.Workbooks.Item(i)
                p = str(getattr(wb, "Path", "") or "").strip()
                if p == "":
                    never_saved_open += 1
            except Exception:
                continue
        if wb_count == 1 and never_saved_open == 1:
            return None, _ERR_COM_UNSAVED_PATH
        return None, _ERR_COM_NOT_OPEN

    wb_com = matches[0]
    if workbook_in_protected_view(xl, wb_com):
        return None, _ERR_COM_PROTECTED_VIEW
    if for_write:
        try:
            if _com_bool_is_true(getattr(wb_com, "ReadOnly", False)):
                return None, _ERR_COM_READ_ONLY
        except Exception:
            pass
    return wb_com, None

def open_workbook_in_excel_com(filepath: str) -> str:
    try:
        import win32com.client
    except ModuleNotFoundError:
        return (
            "Error: COM workbook automation requires Windows with "
            "optional dependency excel-com-mcp[com] (pywin32)."
        )

    is_url = str(filepath).lower().startswith("https://")
    if not is_url:
        p = os.path.abspath(os.path.expanduser(filepath))
        if not os.path.isfile(p):
            return f"Error: File not found: {filepath}"
        open_arg = p
    else:
        open_arg = str(filepath)

    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        try:
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.Visible = True
        except Exception as exc:
            return f"Error: Cannot start or attach to Excel: {exc}"
    try:
        xl.Workbooks.Open(open_arg)
    except Exception as exc:
        return f"Error: {exc}"
    return f"Opened workbook in Excel: {open_arg}"

def close_workbook_in_excel_com(filepath: str, save: bool) -> str:
    wb_com, err = get_open_workbook_com(filepath, for_write=save)
    if err:
        return err
    try:
        wb_com.Close(SaveChanges=bool(save))
    except Exception as exc:
        return f"Error: {exc}"
    what = "saved and closed" if save else "closed without saving"
    return f"Workbook closed in Excel ({what}): {filepath}"

def list_open_workbooks_com(detail: Optional[str] = None) -> str:
    """Walk ``Workbooks`` for ``GetActiveObject("Excel.Application")``."""
    level, detail_err = _normalize_list_open_workbooks_detail(detail)
    if detail_err is not None:
        return detail_err
    try:
        import win32com.client
    except ModuleNotFoundError:
        return (
            "Error: COM workbook automation requires Windows with "
            "optional dependency excel-com-mcp[com] (pywin32)."
        )

    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return "Error: No running Excel application found"

    try:
        wb_count = _coerce_com_count(getattr(xl.Workbooks, "Count", 0))
    except Exception:
        wb_count = 0

    active_norm: Optional[str] = None
    try:
        aw = xl.ActiveWorkbook
        if aw is not None:
            active_norm = _workbook_fullname_norm(aw)
    except Exception:
        active_norm = None

    entries: List[Dict[str, Any]] = []
    for i in range(1, wb_count + 1):
        try:
            wb = xl.Workbooks.Item(i)
        except Exception:
            continue
        try:
            fn_raw = str(wb.FullName)
            short_name = str(wb.Name)
        except Exception:
            continue
        wb_norm = _workbook_fullname_norm(wb)
        is_act = (
            active_norm is not None
            and wb_norm is not None
            and wb_norm == active_norm
        )
        entries.append(
            {
                "full_name": fn_raw,
                "name": short_name,
                "is_active": bool(is_act),
            }
        )

    payload: Dict[str, Any] = {"workbooks": entries}
    if level == "active_context":
        payload.update(_com_active_context_fields(xl))
    return json.dumps(payload, ensure_ascii=False)

