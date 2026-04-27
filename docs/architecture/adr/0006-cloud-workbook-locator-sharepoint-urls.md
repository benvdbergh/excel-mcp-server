# ADR 0006: Cloud workbook locators (SharePoint / OneDrive URLs) for COM routing

## Status

Accepted

## Context

Excel often reports **`Workbook.FullName`** as an **https URL** when a workbook is opened from **SharePoint** or **Microsoft 365** (“cloud-first”), not as a local filesystem path. The fork’s workbook routing was built around **normalized disk paths** (`resolve_target`, `os.path.realpath`, `EXCEL_MCP_ALLOWED_PATHS` containment) and COM matching via `_norm_workbook_path` (FR-1).

Operators may pass a **synced local path** while Excel’s COM identity is the **cloud URL**, so COM routing fails to match. Separately, **`get_excel_path`** rejects non-absolute strings under stdio; `https://` URLs are not Windows “absolute paths,” so valid cloud identities are rejected before routing.

**Authentication:** Opening and syncing the file is **Excel’s responsibility** (signed-in Office identity). The MCP uses **COM against the running Excel process** and does not call SharePoint REST/Graph for workload auth. No supplemental MCP credential is required for the default “workbook already open” scenario.

## Decision

1. Introduce a **cloud workbook locator**: a limited set of non-file schemes (initially **`https:`** only) treated as **opaque document identifiers** for COM routing, not as paths to open with `openpyxl`/file backend.
2. **`filepath` parameter** may therefore be either a **canonical filesystem path** (existing behavior) or an **allowed cloud locator string** (new).
3. **File backend** (`workbook_transport=file`, or auto resolved to file) must **reject** cloud locators with an explicit error (no silent misuse).
4. **Normalization** for COM comparison must use a **URL-safe normalizer** (e.g. `urllib.parse` canonicalization: scheme/host case, path unquoting, slash consistency), distinct from `_norm_workbook_path`, so `FullName` strings match stable operator input.
5. **`EXCEL_MCP_ALLOWED_PATHS`:** When the allowlist is active, cloud locators require an **operator-defined URL prefix allowlist** (new env, e.g. semicolon-separated `EXCEL_MCP_ALLOWED_URL_PREFIXES`) or equivalent fail-closed rule—filesystem roots alone cannot authorize `https://` targets.

## Consequences

- README and TOOLS/manifest tool descriptions: document **both** local absolute paths and **SharePoint-style URLs** as valid **`filepath`** for **COM / auto** when Excel hosts the workbook.
- Tests: unit tests for URL normalization equivalence; COM matching tests with doubles continue to live beside existing path-equivalence tests.
- SSE/HTTP jail (`EXCEL_FILES_PATH`) interaction: confirm whether cloud locators are **stdio-only** or need explicit policy (default: **COM-focused**, document restrictions).

## Links

- FR-1 / path policy: Epic-2, `docs/architecture/target-architecture.md`
- COM matching: `src/excel_mcp/routing/com_workbook_service.py`
