"""A1/region helpers and plain-grid island/layout math (openpyxl-free)."""

from __future__ import annotations

import re
from typing import Any, Mapping, Optional, Sequence

from excel_mcp.exceptions import DataError

from .rows import _as_text, _is_empty_cell, _try_number

_RANGE_RE = re.compile(r"^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$")


def normalize_excel_a1_range(addr: str) -> str:
    """Strip sheet qualifiers and ``$`` from an Excel A1 range address."""
    s = str(addr).strip()
    if "!" in s:
        s = s.split("!", 1)[1]
    return s.replace("$", "")


def _column_letter(col_idx: int) -> str:
    """Convert a 1-based column index to an Excel column letter (no openpyxl)."""
    if col_idx < 1:
        raise ValueError(f"Invalid column index {col_idx}")
    result: list[str] = []
    while col_idx:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def _column_index(col: str) -> int:
    """Convert an Excel column letter to a 1-based index (no openpyxl)."""
    idx = 0
    for ch in col.upper():
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Invalid column letter {col!r}")
        idx = idx * 26 + (ord(ch) - 64)
    if idx < 1:
        raise ValueError(f"Invalid column letter {col!r}")
    return idx


def _range_boundaries(range_string: str) -> tuple[int, int, int, int]:
    """Parse ``A1`` / ``A1:B2`` into ``(min_col, min_row, max_col, max_row)``."""
    msg = f"{range_string} is not a valid coordinate or range"
    m = _RANGE_RE.match(range_string)
    if not m:
        raise ValueError(msg)
    min_col_s, min_row_s, max_col_s, max_row_s = m.groups()
    min_col = _column_index(min_col_s)
    min_row = int(min_row_s)
    if max_col_s is None:
        return min_col, min_row, min_col, min_row
    max_col = _column_index(max_col_s)
    max_row = int(max_row_s)
    return min_col, min_row, max_col, max_row


def a1_range_from_bounds(min_col: int, min_row: int, max_col: int, max_row: int) -> str:
    """Format absolute 1-based bounds as an A1 range."""
    return (
        f"{_column_letter(min_col)}{min_row}:"
        f"{_column_letter(max_col)}{max_row}"
    )


def region_id_for(sheet: str, range_a1: str) -> str:
    """Stable region id: ``{sheet}!{range}``."""
    return f"{sheet}!{normalize_excel_a1_range(range_a1)}"


def parse_region_id(region_id: str, *, sheet_name: str) -> str:
    """Extract the A1 range from a region id; require matching sheet."""
    raw = (region_id or "").strip()
    if "!" not in raw:
        raise DataError(f"Invalid region_id: {region_id!r}")
    sheet, rng = raw.split("!", 1)
    if sheet != sheet_name:
        raise DataError(
            f"region_id sheet {sheet!r} does not match sheet_name {sheet_name!r}"
        )
    if not rng.strip():
        raise DataError(f"Invalid region_id: {region_id!r}")
    return normalize_excel_a1_range(rng)


def guess_header_from_row(values: Sequence[Any]) -> tuple[bool, list[str]]:
    """Return ``(header_guess, columns)``. Columns only when mostly text."""
    non_empty = [v for v in values if not _is_empty_cell(v)]
    if not non_empty:
        return False, []
    text_count = sum(1 for v in non_empty if _try_number(v) is None)
    if text_count <= len(non_empty) / 2:
        return False, []
    return True, [_as_text(v) if not _is_empty_cell(v) else "" for v in values]


def _island_overlaps_box(
    island: tuple[int, int, int, int],
    box: tuple[int, int, int, int],
) -> bool:
    """True when an island ``(min_row, min_col, max_row, max_col)`` hits a range box.

    ``box`` is openpyxl order: ``(min_col, min_row, max_col, max_row)``.
    """
    min_r, min_c, max_r, max_c = island
    b_min_c, b_min_r, b_max_c, b_max_r = box
    return not (
        max_r < b_min_r or min_r > b_max_r or max_c < b_min_c or min_c > b_max_c
    )


def _subtract_exclusion(
    island: tuple[int, int, int, int],
    box: tuple[int, int, int, int],
) -> list[tuple[int, int, int, int]]:
    """Remove one excluded box from an island, returning the leftover rectangles."""
    if not _island_overlaps_box(island, box):
        return [island]
    min_r, min_c, max_r, max_c = island
    b_min_c, b_min_r, b_max_c, b_max_r = box
    pieces: list[tuple[int, int, int, int]] = []
    if min_r < b_min_r:
        pieces.append((min_r, min_c, b_min_r - 1, max_c))
    if max_r > b_max_r:
        pieces.append((b_max_r + 1, min_c, max_r, max_c))
    mid_r1 = max(min_r, b_min_r)
    mid_r2 = min(max_r, b_max_r)
    if mid_r1 <= mid_r2 and min_c < b_min_c:
        pieces.append((mid_r1, min_c, mid_r2, b_min_c - 1))
    if mid_r1 <= mid_r2 and max_c > b_max_c:
        pieces.append((mid_r1, b_max_c + 1, mid_r2, max_c))
    return pieces


def _shrink_to_occupied(
    island: tuple[int, int, int, int],
    occupied: Sequence[Sequence[bool]],
    *,
    origin_row: int,
    origin_col: int,
) -> Optional[tuple[int, int, int, int]]:
    """Shrink a rectangle to occupied cells, or return None when none remain."""
    min_r, min_c, max_r, max_c = island
    found_r1 = found_c1 = found_r2 = found_c2 = None
    nrows = len(occupied)
    ncols = len(occupied[0]) if nrows else 0
    for r in range(min_r, max_r + 1):
        ir = r - origin_row
        if ir < 0 or ir >= nrows:
            continue
        for c in range(min_c, max_c + 1):
            ic = c - origin_col
            if ic < 0 or ic >= ncols or not occupied[ir][ic]:
                continue
            if found_r1 is None:
                found_r1 = found_r2 = r
                found_c1 = found_c2 = c
            else:
                found_r1 = min(found_r1, r)
                found_r2 = max(found_r2, r)
                found_c1 = min(found_c1, c)
                found_c2 = max(found_c2, c)
    if found_r1 is None:
        return None
    assert found_c1 is not None and found_r2 is not None and found_c2 is not None
    return (found_r1, found_c1, found_r2, found_c2)


def _islands_outside_exclusions(
    islands: Sequence[tuple[int, int, int, int]],
    boxes: Sequence[tuple[int, int, int, int]],
    occupied: Sequence[Sequence[bool]],
    *,
    origin_row: int,
    origin_col: int,
) -> list[tuple[int, int, int, int]]:
    """Drop ListObject cells from island bounds, then shrink to occupied cells."""
    out: list[tuple[int, int, int, int]] = []
    for island in islands:
        pieces = [island]
        for box in boxes:
            next_pieces: list[tuple[int, int, int, int]] = []
            for piece in pieces:
                next_pieces.extend(_subtract_exclusion(piece, box))
            pieces = next_pieces
        for piece in pieces:
            shrunk = _shrink_to_occupied(
                piece, occupied, origin_row=origin_row, origin_col=origin_col
            )
            if shrunk is not None:
                out.append(shrunk)
    return out


def _bounds_overlap_or_contain(
    r: int, c: int, boxes: Sequence[tuple[int, int, int, int]]
) -> bool:
    for min_col, min_row, max_col, max_row in boxes:
        if min_row <= r <= max_row and min_col <= c <= max_col:
            return True
    return False


def discover_occupied_islands(
    occupied: Sequence[Sequence[bool]],
    *,
    origin_row: int,
    origin_col: int,
) -> list[tuple[int, int, int, int]]:
    """Find rectangular islands separated by fully blank rows or columns.

    Returns absolute 1-based ``(min_row, min_col, max_row, max_col)`` tuples.
    """
    if not occupied:
        return []
    nrows = len(occupied)
    ncols = len(occupied[0]) if nrows else 0
    if nrows == 0 or ncols == 0:
        return []

    row_has = [any(occupied[r][c] for c in range(ncols)) for r in range(nrows)]
    bands: list[tuple[int, int]] = []
    r = 0
    while r < nrows:
        if not row_has[r]:
            r += 1
            continue
        r0 = r
        while r < nrows and row_has[r]:
            r += 1
        bands.append((r0, r - 1))

    islands: list[tuple[int, int, int, int]] = []
    for r0, r1 in bands:
        col_has = [
            any(occupied[rr][c] for rr in range(r0, r1 + 1)) for c in range(ncols)
        ]
        c = 0
        while c < ncols:
            if not col_has[c]:
                c += 1
                continue
            c0 = c
            while c < ncols and col_has[c]:
                c += 1
            c1 = c - 1
            min_r = max_r = min_c = max_c = None
            for rr in range(r0, r1 + 1):
                for cc in range(c0, c1 + 1):
                    if occupied[rr][cc]:
                        if min_r is None:
                            min_r = max_r = rr
                            min_c = max_c = cc
                        else:
                            min_r = min(min_r, rr)
                            max_r = max(max_r, rr)
                            min_c = min(min_c, cc)
                            max_c = max(max_c, cc)
            if min_r is not None:
                islands.append(
                    (
                        origin_row + min_r,
                        origin_col + min_c,
                        origin_row + max_r,
                        origin_col + max_c,
                    )
                )
    return islands


def build_region_layout_entry(
    *,
    sheet: str,
    min_row: int,
    min_col: int,
    max_row: int,
    max_col: int,
    header_values: Sequence[Any],
) -> dict[str, Any]:
    """One non-table island with a documented header guess."""
    range_a1 = a1_range_from_bounds(min_col, min_row, max_col, max_row)
    header_guess, columns = guess_header_from_row(header_values)
    header_range = a1_range_from_bounds(min_col, min_row, max_col, min_row)
    entry: dict[str, Any] = {
        "id": region_id_for(sheet, range_a1),
        "kind": "region",
        "sheet": sheet,
        "range": range_a1,
        "header_range": header_range,
        "header_row": min_row,
        "columns": columns,
        "header_guess": header_guess,
    }
    return entry


def build_sheet_layout_payload(
    tables: Sequence[Mapping[str, Any]],
    regions: Sequence[Mapping[str, Any]],
) -> dict[str, Any]:
    """Stable sheet-scoped layout envelope."""
    return {"tables": list(tables), "regions": list(regions)}


def layout_regions_from_grid(
    *,
    sheet: str,
    values: Sequence[Sequence[Any]],
    origin_row: int,
    origin_col: int,
    excluded_ranges: Sequence[str],
) -> list[dict[str, Any]]:
    """Discover islands outside ``excluded_ranges`` and attach header guesses."""
    if not values:
        return []
    nrows = len(values)
    ncols = max((len(row) for row in values), default=0)
    if nrows == 0 or ncols == 0:
        return []

    boxes: list[tuple[int, int, int, int]] = []
    for ref in excluded_ranges:
        try:
            boxes.append(_range_boundaries(normalize_excel_a1_range(ref)))
        except Exception:
            continue

    occupied: list[list[bool]] = []
    for ir in range(nrows):
        row_vals = values[ir] if ir < len(values) else []
        row_flags: list[bool] = []
        for ic in range(ncols):
            abs_r = origin_row + ir
            abs_c = origin_col + ic
            cell = row_vals[ic] if ic < len(row_vals) else None
            in_table = _bounds_overlap_or_contain(abs_r, abs_c, boxes)
            row_flags.append(not in_table and not _is_empty_cell(cell))
        occupied.append(row_flags)

    islands = _islands_outside_exclusions(
        discover_occupied_islands(
            occupied, origin_row=origin_row, origin_col=origin_col
        ),
        boxes,
        occupied,
        origin_row=origin_row,
        origin_col=origin_col,
    )
    regions: list[dict[str, Any]] = []
    for min_r, min_c, max_r, max_c in islands:
        header_vals: list[Any] = []
        ir = min_r - origin_row
        for abs_c in range(min_c, max_c + 1):
            ic = abs_c - origin_col
            row_vals = values[ir] if 0 <= ir < len(values) else []
            header_vals.append(row_vals[ic] if 0 <= ic < len(row_vals) else None)
        regions.append(
            build_region_layout_entry(
                sheet=sheet,
                min_row=min_r,
                min_col=min_c,
                max_row=max_r,
                max_col=max_c,
                header_values=header_vals,
            )
        )
    return regions


def rows_from_header_and_data(
    headers: Sequence[str],
    data_matrix: Sequence[Sequence[Any]],
) -> list[dict[str, Any]]:
    """Zip data rows to header keys (pad/truncate to header width)."""
    out: list[dict[str, Any]] = []
    for raw in data_matrix:
        out.append({name: raw[i] if i < len(raw) else None for i, name in enumerate(headers)})
    return out


def resolve_region_query_source(
    *,
    matrix: Sequence[Sequence[Any]],
    range_a1: str,
    require_header_guess: bool,
) -> tuple[list[str], list[dict[str, Any]], str]:
    """Interpret first matrix row as headers (or require a text header guess)."""
    rng = normalize_excel_a1_range(range_a1)
    if not matrix:
        raise DataError(f"Region range {rng!r} is empty")
    header_row = list(matrix[0])
    if require_header_guess:
        ok, columns = guess_header_from_row(header_row)
        if not ok:
            raise DataError(
                "Region has no header guess; pass an explicit range that includes "
                "a header row"
            )
    else:
        columns = [
            _as_text(v) if not _is_empty_cell(v) else "" for v in header_row
        ]
    if not columns:
        raise DataError(f"Region range {rng!r} has no header row")
    data = [list(r) for r in matrix[1:]]
    rows = rows_from_header_and_data(columns, data)
    return columns, rows, rng
