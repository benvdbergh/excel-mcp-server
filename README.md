<p align="center">
  <img src="https://raw.githubusercontent.com/haris-musa/excel-mcp-server/main/assets/logo.png" alt="Excel MCP Server Logo" width="300"/>
</p>

[![PyPI version](https://img.shields.io/pypi/v/excel-com-mcp.svg)](https://pypi.org/project/excel-com-mcp/)
[![Total Downloads](https://static.pepy.tech/badge/excel-com-mcp)](https://pepy.tech/project/excel-com-mcp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Install MCP Server](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/install-mcp?name=excel-com-mcp&config=eyJjb21tYW5kIjoidXZ4IGV4Y2VsLWNvbS1tY3Agc3RkaW8ifQ%3D%3D)

A Model Context Protocol (MCP) server that lets you manipulate Excel files without needing Microsoft Excel installed. Create, read, and modify Excel workbooks with your AI agent.

## Features

- 📊 **Excel Operations**: Create, read, update workbooks and worksheets
- 📈 **Data Manipulation**: Formulas, formatting, charts, pivot tables, and Excel tables
- 🔍 **Data Validation**: Built-in validation for ranges, formulas, and data integrity
- 🎨 **Formatting**: Font styling, colors, borders, alignment, and conditional formatting
- 📋 **Table Operations**: Create and manage Excel tables with custom styling
- 📊 **Chart Creation**: Generate various chart types (line, bar, pie, scatter, etc.)
- 🔄 **Pivot Tables**: Create dynamic pivot tables for data analysis
- 🔧 **Sheet Management**: Copy, rename, delete worksheets with ease
- 🔌 **Triple transport support**: stdio, SSE (deprecated), and streamable HTTP
- 🌐 **Remote & Local**: Works both locally and as a remote service

## Development and CI parity

Install **dev** extras so local runs match PR and release gates ([`docs/architecture/ci-cd-packaging-governance.md`](docs/architecture/ci-cd-packaging-governance.md)):

```bash
pip install -e ".[dev]"
python -m pytest
hatch build
python -m twine check dist/*
```

From the repository root with **uv** (uses `uv.lock` if present):

```bash
uv sync --extra dev
uv run python -m pytest
uv run hatch build
uv run python -m twine check dist/*
```

The PyPI **distribution name** is **`excel-com-mcp`** (same as `[project].name` in `pyproject.toml`). A legacy console entrypoint **`excel-mcp-server`** is also installed. Examples below use **`excel-com-mcp`** for **`uvx`** and MCP JSON so they stay aligned with `manifest.json` → `server.mcp_config` (`command` / `args`).

## Operator documentation map

| Artifact | Purpose |
|----------|---------|
| **This README** | Transports, env vars, **filepath** rules (disk vs SharePoint URL), Cursor/`uv` local MCP setup, allowlists |
| **[`TOOLS.md`](TOOLS.md)** | Per-tool reference; same `filepath` and `workbook_transport` rules apply to every workbook tool |
| **`manifest.json`** | MCP catalog metadata and `mcp_config` (`uvx excel-com-mcp stdio`) |
| **[`.cursor/mcp.json`](.cursor/mcp.json)** | Optional Cursor workspace server using `${workspaceFolder}` + `uv run --project …` (see [Stdio Transport](#1-stdio-transport-for-local-use)) |
| **[`CHANGELOG.md`](CHANGELOG.md)** | Version-to-version release notes and breaking changes |
| **[`docs/plan/transport-routing/IMPLEMENTATION-ROADMAP.md`](docs/plan/transport-routing/IMPLEMENTATION-ROADMAP.md)** | Epic/story delivery status for workbook routing |

### Upgrading from 0.2.x

**Epic 11 / ADR 0008** (release **0.3.0**): **`save_after_write` is removed** from mutating tools—call **`save_workbook`** when you need disk persistence. **Read tools are COM-first** (same as writes) when `workbook_transport` is `auto` or `com` and Excel has the workbook open; for on-disk snapshots use `workbook_transport=file` on reads or save first. New **lifecycle** tools: **`excel_open_workbook`**, **`excel_close_workbook`** (Windows + COM). Details: [`CHANGELOG.md`](CHANGELOG.md), [`TOOLS.md`](TOOLS.md), [ADR 0008](docs/architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md).

### Upgrading from 0.3.x

**Epic 12 / ADR 0009** (release **0.4.0**): **`excel_list_open_workbooks`** lists workbooks open in Excel and returns exact **`full_name`** locators for **`get_workbook_metadata`**, reads, and writes (replaces ad-hoc VBA/Immediate discovery). COM-only; see [`TOOLS.md`](TOOLS.md) and [`CHANGELOG.md`](CHANGELOG.md).

---

## Usage

The server supports three transport methods:

### 1. Stdio Transport (for local use)

```bash
uvx excel-com-mcp stdio
```

```json
{
   "mcpServers": {
      "excel": {
         "command": "uvx",
         "args": ["excel-com-mcp", "stdio"]
      }
   }
}
```

**Local clone in Cursor (this repo):** MCP often starts `uv` with **no project working directory**, so `uv run --extra com …` fails with *“`--extra com` has no effect when used outside of a project`* and *`program not found`*. Pass the project explicitly with **`--project`** (absolute path to the folder that contains `pyproject.toml`):

```json
{
   "mcpServers": {
      "excel-mcp-local": {
         "command": "uv",
         "args": [
            "run",
            "--project",
            "C:/Users/YOU/mcp/excel-mcp-server",
            "--extra",
            "com",
            "excel-com-mcp",
            "stdio"
         ]
      }
   }
}
```

On Windows you can use `C:\\\\Users\\\\YOU\\\\...` instead of forward slashes. Omit `"com"` on non-Windows installs. You can add `"cwd"` with the same path as a hint for other tools, but **`--project` is what fixes `uv run`**.

If Cursor logs **`Failed to spawn: excel-com-mcp`** / **`program not found`**, the client did not resolve the project for `uv run`: add **`--project`** with an absolute path to this repo (or use the workspace file [`.cursor/mcp.json`](.cursor/mcp.json), which passes `${workspaceFolder}`). Relying on **`cwd` alone is not enough** in many Cursor builds.

### 2. SSE Transport (Server-Sent Events - Deprecated)

```bash
uvx excel-com-mcp sse
```

**SSE transport connection**:
```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/sse",
      }
   }
}
```

### 3. Streamable HTTP Transport (Recommended for remote connections)

```bash
uvx excel-com-mcp streamable-http
```

**Streamable HTTP transport connection**:
```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/mcp",
      }
   }
}
```

## Environment Variables & File Path Handling

### SSE and Streamable HTTP Transports

When running the server with the **SSE or Streamable HTTP protocols**, you **must set the `EXCEL_FILES_PATH` environment variable on the server side**. This variable tells the server where to read and write Excel files.
- If not set, it defaults to `./excel_files`.
- With these transports, tool `filepath` values must be **relative** to that directory (e.g. `reports/q1.xlsx`); absolute paths and directory traversal are rejected.

You can also set the `FASTMCP_PORT` environment variable to control the port the server listens on (default is `8017` if not set).
- Example (Windows PowerShell):
  ```powershell
  $env:EXCEL_FILES_PATH="E:\MyExcelFiles"
  $env:FASTMCP_PORT="8007"
  uvx excel-com-mcp streamable-http
  ```
- Example (Linux/macOS):
  ```bash
  EXCEL_FILES_PATH=/path/to/excel_files FASTMCP_PORT=8007 uvx excel-com-mcp streamable-http
  ```

### Stdio Transport

When using the **stdio protocol**, the file path is provided with each tool call, so you do **not** need to set `EXCEL_FILES_PATH` on the server. The server will use the path sent by the client for each operation.

**Workbook identity for COM (Windows):** `filepath` may be a **local absolute path** *or* a **SharePoint-style `https://` URL** that matches `Workbook.FullName` in a running Excel instance. Opening the file, Microsoft 365 sign-in, and sync are handled by **Excel / Office**; the MCP does not perform SharePoint or Graph OAuth. For **read/write via the file backend** (`openpyxl`), use a real filesystem path. When **`EXCEL_MCP_ALLOWED_PATHS`** is enabled, `https` locators also require **`EXCEL_MCP_ALLOWED_URL_PREFIXES`** (see below).

#### SharePoint / Microsoft 365: use the same string as Excel’s `FullName`

Excel often reports cloud-backed workbooks with an **`https://…sharepoint.com/…`** identity even when a synced copy exists on disk. The MCP compares your `filepath` to COM `Workbook.FullName` **after normalization** (ADR [0006](docs/architecture/adr/0006-cloud-workbook-locator-sharepoint-urls.md)). If they differ, routing fails (`Workbook not open…`), and **`auto`** may fall back to the **file** backend and hit **`Permission denied`** while Excel holds the file.

1. In Excel: **Alt+F11** → **Immediate** window → run: `? ActiveWorkbook.FullName` and press Enter.
2. Use that **exact** returned string (including `https://`) as **`filepath`** on tools.
3. Prefer **`workbook_transport=com`** (or **`auto`** once the identity matches and open detection can see the workbook) for edits in the Excel session.

Use the **local absolute path** only when `FullName` is a normal file path, or when you intentionally use the **file** backend (e.g. headless `openpyxl`).

### Path normalization (`resolve_target`)

Internally, workbook targets are normalized with **`resolve_target`** in `excel_mcp.path_resolution` (single entry point for **FR-1** / future COM path comparison). It uses `os.path.realpath` for stable absolute paths; relative resolution order (`search_roots`, then `cwd`) is documented in that module.

- **stdio, allowlist off:** tool paths must still be **absolute**; the server returns `os.path.normpath` only (legacy behavior for existing clients).
- **stdio, allowlist on** and **SSE / streamable HTTP:** paths are finalized with `resolve_target` before jail and allowlist checks.

Integrators can reuse **`path_is_allowed`** / **`assert_path_allowed`** from `excel_mcp.path_policy` so file and future COM backends share the same policy.

### Optional path allowlist (`EXCEL_MCP_ALLOWED_PATHS`)

For tighter control (aligned with **FR-11**), set **`EXCEL_MCP_ALLOWED_PATHS`** to one or more allowed **directory** roots, separated by **`os.pathsep`** (semicolon on Windows, colon on macOS/Linux—the same rule as the `PATH` environment variable). Whitespace around each entry is trimmed; `~` is expanded per entry.

When unset or blank, behavior matches the pre-fork defaults: stdio accepts any absolute path (subject to existing validation), and SSE/HTTP use only the `EXCEL_FILES_PATH` jail.

When set:

- **stdio:** each workbook path is resolved with `resolve_target`, then must lie **inside** at least one listed root (directory containment, same idea as the remote jail).
- **SSE / streamable HTTP:** the resolved path must be inside **`EXCEL_FILES_PATH`** *and* inside at least one allowlist root (**intersection**).

If the variable is non-empty but **no** root path resolves (typos, missing drive letters, unreadable paths), the allowlist is treated as **active with zero valid roots** and paths are **rejected until the environment is corrected** (fail-closed).

Examples:

```powershell
# Windows: two roots (note the semicolon)
$env:EXCEL_MCP_ALLOWED_PATHS = "E:\Workbooks;E:\SharedTemplates"
uvx excel-com-mcp stdio
```

```bash
# Linux / macOS: colon-separated
EXCEL_MCP_ALLOWED_PATHS=/var/excel-in:/var/excel-out uvx excel-com-mcp stdio
```

### URL prefix allowlist for cloud locators (`EXCEL_MCP_ALLOWED_URL_PREFIXES`)

When **`EXCEL_MCP_ALLOWED_PATHS`** is set (path allowlist **on**), **`https://` cloud workbook locators** must also match at least one entry in **`EXCEL_MCP_ALLOWED_URL_PREFIXES`**, or they are rejected (fail-closed). This is separate from directory containment: filesystem roots do not authorize arbitrary SharePoint URLs.

- Segments are separated with **semicolons (`;`)** on every platform (not `PATH`-style `os.pathsep`: `https://` URLs contain colons, so colon-separated lists break on Linux/macOS).
- Each segment is a **prefix URL** (scheme `https`, host, and optional path) after the same canonicalization as workbook URLs. Prefer a **trailing slash** on path prefixes (e.g. `https://contoso.sharepoint.com/sites/Team/`) so only that hierarchy is allowed.
- If the path allowlist is on but this variable is missing, empty, or has no valid `https` entries, **https workbook targets are denied**.

When the path allowlist is **off**, `https` locators are accepted without `EXCEL_MCP_ALLOWED_URL_PREFIXES` (subject to normal `https` validation in `get_excel_path`).

### Workbook transport and COM policy (not MCP wire transport)

These environment variables control **workbook** routing (file-backed ``openpyxl`` path vs COM automation when wired in later stories). They do **not** select the MCP client↔server **wire** transport (stdio, SSE, or streamable HTTP); that is configured by how you launch the server (see above). See ADR 0001 for the vocabulary split.

| Variable | Meaning |
| -------- | ------- |
| ``EXCEL_MCP_TRANSPORT`` | Workbook mode: ``auto``, ``file``, or ``com`` (case-insensitive). Default ``auto`` when unset or empty. Invalid values raise at read time. Parsed by ``excel_mcp.routing.read_workbook_transport``. |
| ``EXCEL_MCP_COM_STRICT`` | When ``1`` / ``true`` / ``yes`` (case-insensitive): strict COM policy. When ``0`` / ``false`` / ``no``, or explicitly relaxed: non-strict. **Unset or empty defaults to strict** (``True``). Parsed by ``read_com_strict``. |
| ``EXCEL_MCP_COM_ALLOW_FILE_FALLBACK`` | When ``1`` / ``true`` / ``yes``: operators allow documented file fallback in scenarios where non-strict routing would apply (ADR 0005). Unset or empty: ``False``. Parsed by ``read_com_allow_file_fallback``. |
| ``EXCEL_MCP_ALLOWED_URL_PREFIXES`` | When ``EXCEL_MCP_ALLOWED_PATHS`` is set: **required** for `https` workbook locators. **Semicolon-separated** **https** URL prefixes on all OSes. See [URL prefix allowlist](#url-prefix-allowlist-for-cloud-locators-excel_mcp_allowed_url_prefixes). |

**Effective strictness for the router** is ``effective_com_strict()``: ``False`` if file fallback is allowed **or** ``EXCEL_MCP_COM_STRICT`` is explicitly falsy; otherwise ``True``. Allowing file fallback forces non-strict effective behavior whenever that flag is on.

### Optional MCP tool parameters (workbook routing)

Most workbook tools accept optional **`workbook_transport`** (``auto`` \| ``file`` \| ``com``). When omitted, transport defaults to ``EXCEL_MCP_TRANSPORT``. **Discovery** (**`excel_list_open_workbooks`**) and **lifecycle** tools **`excel_open_workbook`** / **`excel_close_workbook`** do not use this parameter (they are COM-only per ADR 0009 / ADR 0008). For **persistence** after mutations, call **`save_workbook`** (ADR 0008). These names refer to **workbook** execution routing (ADR 0001), not MCP wire transport.

### Routing observability

Routed workbook operations (via ``execute_routed_workbook_operation`` in ``excel_mcp.routing.routed_dispatch``) emit **one JSON object per dispatch** on logger ``excel-mcp.routing`` at INFO (no stdout). Fields follow ADR 0001 vocabulary:

- **workbook_transport** — requested mode: ``auto``, ``file``, or ``com``.
- **workbook_backend** — resolved backend after the selection matrix: ``file`` or ``com``.
- **routing_reason** — stable reason string from ``RoutingBackend`` (e.g. ``forced_file``, ``full_name_match``, ``auto_workbook_not_open_file``, ``v1_file_forced``).
- **duration_ms** — wall time for resolve plus executed file I/O (when applicable).
- **workbook_path** — redacted path (basename only by default; set ``EXCEL_MCP_LOG_FULL_PATHS=1`` for full path in break-glass scenarios).
- **operation_name** — routed contract method name (e.g. ``read_range_with_metadata``).
- **mcp_tool_name** — optional registered MCP tool name when supplied by the caller.
- **v1_file_forced** — `true` when **ADR 0004** forces the **file** backend for a tool (chart / pivot v1) regardless of `auto`→COM for other writes.

### ADR 0008 / ADR 0003 — COM-first reads and when disk matters

Read-class tools use the **same routing** as writes (**COM-first** when `transport=auto`/`com`, the workbook matches an open host, and COM is viable; otherwise **openpyxl** file path). Live grid reads therefore follow **Excel** when COM is selected, not necessarily the last saved file. If you rely on **on-disk** snapshots (external tools, `workbook_transport=file`), or after COM writes you need the file to match the host, call **`save_workbook`** before file-backed operations. **`create_chart`** / **`create_pivot_table`** remain **file-forced** (ADR 0004). Historical “file-default reads” (ADR 0007) are superseded; see [`docs/architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md`](docs/architecture/adr/0008-com-first-default-and-file-lifecycle-tools.md) and [`docs/architecture/adr/0003-read-path-com-parity.md`](docs/architecture/adr/0003-read-path-com-parity.md).

### FR-9 — Protected View, read-only, duplicate instances

COM operations return **clear, fail-closed** errors when Excel blocks writes (e.g. **Protected View**, **read-only**). If **multiple open workbooks** resolve to the **same path** across Excel instances, routing fails closed with an error asking the operator to **close duplicates** or use a **single** Excel instance.

### NFR-2 — routing latency

**p95 routing overhead** is **not** continuously benchmarked in CI. See [`docs/performance/routing-nfr2-note.md`](docs/performance/routing-nfr2-note.md) for an honest scope note and an optional local micro-benchmark idea.

### NFR-4 — elevation

The server **does not** request **administrator elevation** by default (COM automation targets the user’s normal Excel session; see also the blueprint §5 “do not start Excel as admin without explicit user opt-in”).

**Planning / delivery status:** workbook transport epics and stories are tracked in [`docs/plan/transport-routing/IMPLEMENTATION-ROADMAP.md`](docs/plan/transport-routing/IMPLEMENTATION-ROADMAP.md). Phases **1–9** and **Epic 11** (COM-first reads, explicit **`save_workbook`**, lifecycle tools, docs/tests) are **delivered** per that roadmap. **Epic 7** remains the vertical slice for COM write-class parity, **`save_workbook`**, **FR-9** errors, and **ADR 0004** v1 file-forced **`create_chart`** / **`create_pivot_table`** (logs: `routing_reason` **`v1_file_forced`** when applicable).

**Optional Windows COM (`[com]`):** to install pywin32 for COM-backed workbook routing, use `pip install excel-com-mcp[com]` (or the equivalent for your installer). pywin32 is distributed under the [PSF License Agreement](https://github.com/mhammond/pywin32/blob/main/LICENSE.txt) (same terms as CPython).

### COM execution threading (Windows)

COM apartment rules require Excel automation from a **consistent thread**. The server uses ``excel_mcp.com_executor.ComThreadExecutor``: a **single worker thread** pulls jobs from a queue; ``submit(fn, *args, **kwargs)`` runs ``fn`` on that thread and **blocks** the caller until the result is ready (or an exception is propagated), so synchronous MCP tool handlers stay compatible without turning every tool ``async``. On Windows, that worker thread calls ``pythoncom.CoInitialize()`` before processing jobs and ``CoUninitialize()`` on shutdown so COM APIs such as ``GetActiveObject("Excel.Application")`` behave like a normal main-thread script (without this, automation from a plain background thread often fails even when Excel is running). The executor does **not** start Excel by itself. For tests or clean process teardown, call ``shutdown(wait=True)``; abrupt exit may still cut off in-flight work—see the module docstring on ``com_executor`` for limitations (including no reentrant ``submit`` from inside a job on the worker).

### Windows manual smoke (COM write path)

- Install the optional stack: ``pip install "excel-com-mcp[com]"`` (or your package equivalent) so ``pywin32`` is available.
- Start **Microsoft Excel** manually and open the target ``.xlsx`` using **File → Open** (a running instance with the workbook loaded is required; the server does not launch Excel).
- Run the MCP server (e.g. stdio) on the same Windows machine with routing env vars as needed (defaults: ``EXCEL_MCP_TRANSPORT=auto`` when unset).
- Call **``write_data_to_excel``** with an **absolute** path to that file, ``workbook_transport=com`` or ``auto``, and a small ``data`` grid; with ``auto``, the workbook must be detected as open in Excel for COM to win.
- Optional: call **`excel_list_open_workbooks`** when you need the exact **`filepath`**/`FullName` Excel reports (replacing VBA Immediate); then **`get_workbook_metadata`** / reads with that string.
- Optional: use **`excel_open_workbook`** (or **`create_workbook(..., open_in_excel=true)`**) so Excel has the path open for **auto**→COM routing.
- Read tools (e.g. ``read_data_from_excel``) use the **same** COM/file matrix; after COM writes, call **`save_workbook`** if you need the **on-disk** file to match the host.
- Confirm routing in ``excel-mcp.log``: one JSON line per dispatch with ``workbook_backend`` ``com`` and a stable ``routing_reason`` (e.g. ``full_name_match`` / ``forced_com``) for writes routed to COM.
- For release-style sign-off, follow **[`docs/plan/transport-routing/MANUAL-WINDOWS-RC-CHECKLIST.md`](docs/plan/transport-routing/MANUAL-WINDOWS-RC-CHECKLIST.md)** (Protected View, read-only, duplicate instance, save-then-read, chart/pivot ``v1_file_forced`` rows).

### CI locally (contributors)

GitHub Actions runs the same gates as **PR CI** via [`.github/workflows/reusable-validate-and-test.yml`](.github/workflows/reusable-validate-and-test.yml): editable install with **dev** extras, **pytest**, **`hatch build`**, and **`twine check dist/*`**.

```bash
python -m pip install --upgrade pip
pip install -e ".[dev]"
pytest
hatch build
twine check dist/*
```

No Excel or Windows COM is required for this default path (Linux CI).

## Available Tools

The server provides a comprehensive set of Excel manipulation tools. See [TOOLS.md](TOOLS.md) for complete documentation of all available tools.

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=benvdbergh/excel-mcp-server&type=Date)](https://www.star-history.com/#benvdbergh/excel-mcp-server&Date)

## License

MIT License - see [LICENSE](LICENSE) for details.
