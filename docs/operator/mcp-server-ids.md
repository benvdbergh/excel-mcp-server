# MCP server identity registry

Single source of truth for **Excel MCP** server names in Cursor and related configs. The **office-xlsx** skill references this table — do not duplicate it in skill files.

**Package:** PyPI distribution **`excel-com-mcp`** (`manifest.json` → pinned `uvx excel-com-mcp==0.5.0 stdio`; see [README install decision matrix](../../README.md#install-decision-matrix)).

## Registry

| Agent server id | `mcp.json` key | Config pattern | Package | Notes |
|-----------------|----------------|----------------|---------|-------|
| **`user-excel`** | `excel` | `uvx excel-com-mcp==0.5.0 stdio` | PyPI | Global default; pin version — unpinned `@latest` is legacy |
| **`user-excel-local`** | `excel-local` | `uv run --project <fork> --extra com excel-com-mcp stdio` | Fork | Recommended for COM development and SharePoint allowlists |
| *(same as row above)* | *(workspace)* | Same as `excel-local` | Fork | Cursor caches tool descriptors under `mcps/user-excel-local/` (`serverName`: `excel-local`) |

### Naming rules

- **`mcp.json` key** — the JSON object key under `mcpServers` (operator-chosen; use **`excel`** and **`excel-local`** for consistency).
- **Agent server id** — what agents and the MCP tool panel use. Cursor maps user-level servers to **`user-<mcp.json key>`** (see `mcps/<id>/SERVER_METADATA.json`: `serverIdentifier` vs `serverName`).
- **Do not** use legacy keys such as `excel-mcp-local` in new configs; rename to **`excel-local`**.

On Windows fork installs, pass **`--project`** with an absolute path to the repo root (folder containing `pyproject.toml`). Omit `--extra com` on non-Windows. See [README § Stdio Transport](../../README.md#1-stdio-transport-for-local-use).

## Canonical config snippets

**Global PyPI (`excel` → `user-excel`):**

```json
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["excel-com-mcp==0.5.0", "stdio"]
    }
  }
}
```

**Local fork (`excel-local` → `user-excel-local`):**

```json
{
  "mcpServers": {
    "excel-local": {
      "command": "uv",
      "args": [
        "run",
        "--project",
        "${workspaceFolder}",
        "--extra",
        "com",
        "excel-com-mcp",
        "stdio"
      ]
    }
  }
}
```

Optional: add `env` (e.g. `EXCEL_MCP_TRANSPORT`, `EXCEL_MCP_ALLOWED_URL_PREFIXES` for SharePoint). Workspace file [`.cursor/mcp.json`](../../.cursor/mcp.json) ships the `excel-local` fork pattern with example operator env for this repo.

## Deployment profiles (this operator)

| Host | Excel MCP | Server id agents see |
|------|-----------|----------------------|
| **Global Cursor** (`~/.cursor/mcp.json`) | **`excel`** (PyPI `uvx`) + **`excel-local`** (fork) | `user-excel`, `user-excel-local` |
| **excel-mcp-server workspace** | **`excel-local`** via project MCP when configured | `user-excel-local` |
| **Ai-Vault project** | **`excel-local`** fork pin (`C:/Users/vandenbb/mcp/excel-mcp-server`) + SharePoint URL prefix | `user-excel-local` |

Prefer **`user-excel-local`** / **`excel-local`** when COM routing, fork fixes, or SharePoint URL allowlists matter. Use **`user-excel`** for quick PyPI-only file workflows without a local clone.

## Related docs

- [README operator documentation map](../../README.md#operator-documentation-map)
- [TOOLS.md](../../TOOLS.md) — tool reference (same `filepath` / `workbook_transport` rules for every server id)
- **office-xlsx** skill `references/excel-mcp-server.md` — execution contract (links here for server identity only)
