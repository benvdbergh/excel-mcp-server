# Changelog

All notable changes to this project are documented in this file. The format is informal; align version bumps with [Semantic Versioning](https://semver.org/) and [docs/architecture/release-versioning-policy.md](docs/architecture/release-versioning-policy.md).

## 0.2.0 — 2026-04-28

### Added

- Cloud (SharePoint-style) `https://` workbook locators for stdio/COM on Windows; `normalize_workbook_target_for_com` matches Excel `Workbook.FullName`; `EXCEL_MCP_ALLOWED_URL_PREFIXES` when `EXCEL_MCP_ALLOWED_PATHS` is set (semicolon-separated URL prefixes on **all** OSes).

### Changed

- Operator docs: README, `TOOLS.md`, `manifest.json`, MCP server `instructions`; Epic 9 / ADR 0006; workspace [`.cursor/mcp.json`](.cursor/mcp.json) for local `uv run --project`.

### Fixed

- `EXCEL_MCP_ALLOWED_URL_PREFIXES` parsing no longer uses `os.pathsep` on POSIX (colons in `https://` broke Linux/macOS CI and production allowlists).
