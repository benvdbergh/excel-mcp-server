# Pre-fork architecture (starting point)

This document describes the **as-is** architecture of this repository before workbook transport routing (file vs COM) and related PRD work. It is the baseline for comparing the **target** architecture in `target-architecture.md`.

## Purpose and scope

The Excel MCP Server exposes **Model Context Protocol (MCP)** tools so clients can create, read, and mutate `.xlsx` workbooks using **Python** and **openpyxl**. All workbook access is **file-based**: paths resolve to on-disk files, which are opened and saved with openpyxl. There is **no** integration with a running Microsoft Excel instance via COM.

## Runtime and deployment

| Aspect | Detail |
|--------|--------|
| **Language** | Python ≥ 3.10 |
| **Entry** | Typer CLI: `excel-mcp-server` → `excel_mcp.__main__:app` |
| **MCP framework** | `FastMCP` (`mcp.server.fastmcp`) |
| **Core I/O** | `openpyxl` only (no xlwings / pywin32 in default dependencies) |
| **Manifest** | `manifest.json` lists MCP tools; aligns with `@mcp.tool` function names in `server.py` |

## Wire transports (not workbook transport)

The server can be started in three **network/process** modes. These control how MCP messages are carried (stdio vs HTTP/SSE), **not** whether Excel uses files or COM.

| Mode | Function | `EXCEL_FILES_PATH` | Path rules for `filepath` argument |
|------|----------|--------------------|-----------------------------------|
| **stdio** | `run_stdio()` | Left `None` | Client must pass an **absolute** path; `os.path.normpath` only |
| **SSE** | `run_sse()` | Set from `EXCEL_FILES_PATH` env (default `./excel_files`) | Path must be **relative** to that directory; `realpath` + directory containment check |
| **Streamable HTTP** | `run_streamable_http()` | Same as SSE | Same as SSE |

Logging uses a **file** handler (`excel-mcp.log` under repo root) so stdio stdout stays valid for MCP JSON-RPC only.

## Module layout (`src/excel_mcp`)

```
excel_mcp/
  __main__.py      # Typer CLI: stdio | sse | streamable_http
  server.py        # FastMCP instance, path resolution, all @mcp.tool handlers
  workbook.py      # Create/open/metadata; load_workbook / Workbook / save
  sheet.py         # Sheet copy/delete/rename, merge, ranges, row/col ops
  data.py          # read_excel_range*, write_data
  formatting.py    # format_range via get_or_create_workbook
  calculations.py  # apply_formula via get_or_create_workbook
  validation.py    # Formula/range validation against loaded workbook
  chart.py, tables.py, pivot.py  # Feature modules: load → mutate → save
  cell_utils.py, cell_validation.py  # Helpers; cell_validation is worksheet-bound
  exceptions.py    # Domain exception hierarchy
```

**Coupling pattern:** Feature modules call `load_workbook(filepath)` or `get_or_create_workbook(filepath)` directly. There is **no** shared workbook session service or transport abstraction.

## Path resolution and trust

All tools that need a file go through **`get_excel_path(filename)`** in `server.py`:

- Rejects empty paths and NUL bytes.
- **stdio:** absolute paths only → `normpath`; no sandbox root (full trust in client-supplied absolute paths).
- **SSE / HTTP:** relative paths only, resolved under `EXCEL_FILES_PATH`, with **`_resolved_path_is_within`** using `realpath` and `os.path.commonpath` to prevent traversal outside the jail.

`validation.py` does **not** enforce paths; it validates formulas and ranges for workbooks already opened by path elsewhere.

## MCP tool surface

Tools are registered **only** in `server.py` via `@mcp.tool(...)`. Handler names equal MCP tool names (FastMCP convention). There are **25** tools, all **synchronous** `def` handlers returning `str` (success message or `"Error: ..."` for many caught domain exceptions).

**Typical handler flow:**

1. `full_path = get_excel_path(filepath)`
2. Call into `workbook` / `sheet` / `data` / `formatting` / `calculations` / `chart` / `pivot` / `tables` / `validation`
3. Catch domain exceptions → return `f"Error: {str(e)}"`; some paths re-raise generic `Exception`

**Exception:** `get_data_validation_info` performs **`load_workbook` inline** in `server.py` instead of delegating through a shared workbook access layer.

## Read vs write at the file layer

| Style | Examples |
|-------|----------|
| **load → mutate → save** | Most sheet ops, `write_data`, chart/table/pivot, formatting |
| **get_or_create → mutate → save** | `formatting.format_range`, `calculations.apply_formula` (may create new file) |
| **load → read → close** (intended) | `data.read_excel_range*`, `get_workbook_info` |
| **load → read** (lifecycle inconsistent in some helpers) | e.g. `get_merged_ranges`, some validation reads may omit `wb.close()` |

## Dependencies (`pyproject.toml`)

- `mcp[cli]`, `fastmcp`, `openpyxl`, `typer`
- No optional COM extras in the pre-fork baseline.

## Non-goals in this baseline

- Routing between file I/O and Excel COM.
- Per-tool `workbook_transport` / `save_after_write` parameters.
- Structured logs for chosen backend (`transport`, `reason`, `duration_ms`).
- Windows-only COM tests or optional `[com]` install.

## Summary

The pre-fork system is a **monolithic MCP server module** (`server.py`) on top of **distributed openpyxl file access** across feature modules, with a **single path gate** (`get_excel_path`) and **two containment modes** (trusted absolute stdio vs jailed relative HTTP/SSE). This is the structural baseline the fork extends with **`RoutingBackend`**, **`FileWorkbookService`**, **`ComWorkbookService`**, and shared normalization and observability.
