# ADR 0001: Workbook transport vs MCP wire transport naming

## Status

Accepted

## Context

The product introduces routing between **file-based** workbook access and **COM** (Excel host) access. The MCP Python stack already uses the word **transport** for how the server speaks MCP to clients (`stdio`, `sse`, `streamable-http` via `mcp.run(transport=...)`).

Reusing the term `transport` for file-vs-COM without qualification will confuse maintainers, operators, and documentation readers.

## Decision

1. **MCP wire transport** — Keep existing names: stdio, SSE, streamable HTTP; documented as “client connection mode.”
2. **Workbook backend / workbook transport** — Use distinct vocabulary in code, logs, and tool parameters, for example:
   - `workbook_transport` with values `auto`, `file`, `com`
   - Environment variable `EXCEL_MCP_TRANSPORT` documented explicitly as “workbook transport,” not MCP wire.

## Consequences

- Tool schemas and README must use clear headings (“MCP connection” vs “Excel workbook access”).
- Code search for `transport=` must distinguish FastMCP’s wire parameter from workbook routing.
