# ADR 0005: COM strict mode and optional file fallback

## Status

Accepted

## Context

Operators need predictable behavior:

- **`EXCEL_MCP_COM_STRICT`** (PRD US-5): When the user expects COM (`workbook_transport=com` or policy implies COM), the server must **error** instead of silently using the file backend if the workbook is not open in Excel.
- The PRD also mentions an optional **`allow_com_fallback`** (naming TBD) for documented silent fallback — env-only vs per-tool exposure is unresolved.

## Decision

1. **`EXCEL_MCP_COM_STRICT=1`:** When effective mode is **`com`** and the workbook is **not** detected as open (or COM is unavailable), return a **documented error** (structured / typed per error-handling ADR). **No** silent file write in strict mode.
2. **Fallback flag:** If implemented, default **off**; name either `EXCEL_MCP_COM_ALLOW_FILE_FALLBACK=1` (env-first) or an optional tool parameter — **record final name and surface** in README when implemented. Prefer **env-only** for v1 to keep tool schemas smaller unless integrators require per-call override.

## Consequences

- Router tests must cover: `com` + closed book + strict → error; `com` + closed book + fallback on → file (if ever implemented).
- Client-visible error messages must be stable enough for automation (avoid changing strings every release without notice).
