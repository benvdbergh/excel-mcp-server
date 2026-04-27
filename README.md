<p align="center">
  <img src="https://raw.githubusercontent.com/haris-musa/excel-mcp-server/main/assets/logo.png" alt="Excel MCP Server Logo" width="300"/>
</p>

[![PyPI version](https://img.shields.io/pypi/v/excel-mcp-server.svg)](https://pypi.org/project/excel-mcp-server/)
[![Total Downloads](https://static.pepy.tech/badge/excel-mcp-server)](https://pepy.tech/project/excel-mcp-server)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![smithery badge](https://smithery.ai/badge/@haris-musa/excel-mcp-server)](https://smithery.ai/server/@haris-musa/excel-mcp-server)
[![Install MCP Server](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/install-mcp?name=excel-mcp-server&config=eyJjb21tYW5kIjoidXZ4IGV4Y2VsLW1jcC1zZXJ2ZXIgc3RkaW8ifQ%3D%3D)

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

## Usage

The server supports three transport methods:

### 1. Stdio Transport (for local use)

```bash
uvx excel-mcp-server stdio
```

```json
{
   "mcpServers": {
      "excel": {
         "command": "uvx",
         "args": ["excel-mcp-server", "stdio"]
      }
   }
}
```

### 2. SSE Transport (Server-Sent Events - Deprecated)

```bash
uvx excel-mcp-server sse
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
uvx excel-mcp-server streamable-http
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
  uvx excel-mcp-server streamable-http
  ```
- Example (Linux/macOS):
  ```bash
  EXCEL_FILES_PATH=/path/to/excel_files FASTMCP_PORT=8007 uvx excel-mcp-server streamable-http
  ```

### Stdio Transport

When using the **stdio protocol**, the file path is provided with each tool call, so you do **not** need to set `EXCEL_FILES_PATH` on the server. The server will use the path sent by the client for each operation.

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
uvx excel-mcp-server stdio
```

```bash
# Linux / macOS: colon-separated
EXCEL_MCP_ALLOWED_PATHS=/var/excel-in:/var/excel-out uvx excel-mcp-server stdio
```

### Workbook transport and COM policy (not MCP wire transport)

These environment variables control **workbook** routing (file-backed ``openpyxl`` path vs COM automation when wired in later stories). They do **not** select the MCP client↔server **wire** transport (stdio, SSE, or streamable HTTP); that is configured by how you launch the server (see above). See ADR 0001 for the vocabulary split.

| Variable | Meaning |
| -------- | ------- |
| ``EXCEL_MCP_TRANSPORT`` | Workbook mode: ``auto``, ``file``, or ``com`` (case-insensitive). Default ``auto`` when unset or empty. Invalid values raise at read time. Parsed by ``excel_mcp.routing.read_workbook_transport``. |
| ``EXCEL_MCP_COM_STRICT`` | When ``1`` / ``true`` / ``yes`` (case-insensitive): strict COM policy. When ``0`` / ``false`` / ``no``, or explicitly relaxed: non-strict. **Unset or empty defaults to strict** (``True``). Parsed by ``read_com_strict``. |
| ``EXCEL_MCP_COM_ALLOW_FILE_FALLBACK`` | When ``1`` / ``true`` / ``yes``: operators allow documented file fallback in scenarios where non-strict routing would apply (ADR 0005). Unset or empty: ``False``. Parsed by ``read_com_allow_file_fallback``. |
| ``EXCEL_MCP_SAVE_AFTER_WRITE_DEFAULT`` | Default for optional tool parameter ``save_after_write`` when omitted on mutating tools. ``1`` / ``true`` / ``yes`` → default **true** (extra ``save_workbook`` after file-backed writes). **Unset or empty defaults to false** (FR-8: no extra flush until requested). Parsed by ``read_save_after_write_default``. |

**Effective strictness for the router** is ``effective_com_strict()``: ``False`` if file fallback is allowed **or** ``EXCEL_MCP_COM_STRICT`` is explicitly falsy; otherwise ``True``. Allowing file fallback forces non-strict effective behavior whenever that flag is on.

### Optional MCP tool parameters (workbook routing)

Every workbook tool accepts optional **`workbook_transport`** (``auto`` \| ``file`` \| ``com``) and **`save_after_write`** (boolean). When omitted, transport defaults to ``EXCEL_MCP_TRANSPORT`` and the save flag defaults per ``EXCEL_MCP_SAVE_AFTER_WRITE_DEFAULT``. **Read-only tools** ignore ``save_after_write`` (no extra save). These names refer to **workbook** execution routing (ADR 0001), not MCP wire transport.

### Routing observability

Routed workbook operations (via ``execute_routed_workbook_operation`` in ``excel_mcp.routing.routed_dispatch``) emit **one JSON object per dispatch** on logger ``excel-mcp.routing`` at INFO (no stdout). Fields follow ADR 0001 vocabulary:

- **workbook_transport** — requested mode: ``auto``, ``file``, or ``com``.
- **workbook_backend** — resolved backend after the selection matrix: ``file`` or ``com``.
- **routing_reason** — stable reason string from ``RoutingBackend`` (e.g. ``forced_file``, ``full_name_match``).
- **duration_ms** — wall time for resolve plus executed file I/O (when applicable).
- **workbook_path** — redacted path (basename only by default; set ``EXCEL_MCP_LOG_FULL_PATHS=1`` for full path in break-glass scenarios).
- **operation_name** — routed contract method name (e.g. ``read_range_with_metadata``).
- **mcp_tool_name** — optional registered MCP tool name when supplied by the caller.

**Planning / delivery status:** workbook transport epics and stories (Phases 1–5) are tracked in [`docs/plan/transport-routing/IMPLEMENTATION-ROADMAP.md`](docs/plan/transport-routing/IMPLEMENTATION-ROADMAP.md) and marked **done** in the epic/story frontmatter through Epic 5; COM packaging (Epics 6–7) is still **draft**.

## Available Tools

The server provides a comprehensive set of Excel manipulation tools. See [TOOLS.md](TOOLS.md) for complete documentation of all available tools.

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=haris-musa/excel-mcp-server&type=Date)](https://www.star-history.com/#haris-musa/excel-mcp-server&Date)

## License

MIT License - see [LICENSE](LICENSE) for details.
