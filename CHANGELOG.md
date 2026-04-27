# Changelog

All notable changes to this project are documented in this file. The format is informal; align version bumps with [Semantic Versioning](https://semver.org/) and [docs/architecture/release-versioning-policy.md](docs/architecture/release-versioning-policy.md).

## 0.3.0 — 2026-04-2

### Added

- **COM session lifecycle (ADR 0008):** `excel_open_workbook` and `excel_close_workbook` (Windows + COM) to bind Excel host state; `create_workbook(..., open_in_excel=true)` for post-create open.
- **Full read-class COM wiring** via `com_do_op` and `ComWorkbookService` parity with file-backed contracts where applicable.

### Changed

- **COM-first default routing (Epic 11 / ADR 0008):** read-class tools use the same **COM-first / file fallback** matrix as writes when `workbook_transport` is `auto` or `com` (supersedes ADR 0007 file-default reads). Live grid reads follow Excel when COM wins; use `save_workbook` before relying on on-disk snapshots or `workbook_transport=file`.
- **Explicit save only:** `save_after_write` removed from all mutating tool signatures and env; call **`save_workbook`** when persistence is required.

### Breaking

- Any client or prompt that passed **`save_after_write`** must drop it and use **`save_workbook`** after writes.
- Agents expecting **file-default reads** on Windows with Excel open should set **`workbook_transport=file`** for disk snapshots or call **`save_workbook`** then read, per ADR 0008.

### Docs

- README, `TOOLS.md`, routing observability, and [`docs/plan/transport-routing/IMPLEMENTATION-ROADMAP.md`](docs/plan/transport-routing/IMPLEMENTATION-ROADMAP.md) updated for Epic 11; [`manifest.json`](manifest.json) catalog includes lifecycle tools and `0.3.0`.

## 0.2.0 — 2026-04-28

### Added

- Cloud (SharePoint-style) `https://` workbook locators for stdio/COM on Windows; `normalize_workbook_target_for_com` matches Excel `Workbook.FullName`; `EXCEL_MCP_ALLOWED_URL_PREFIXES` when `EXCEL_MCP_ALLOWED_PATHS` is set (semicolon-separated URL prefixes on **all** OSes).

### Changed

- Operator docs: README, `TOOLS.md`, `manifest.json`, MCP server `instructions`; Epic 9 / ADR 0006; workspace [`.cursor/mcp.json`](.cursor/mcp.json) for local `uv run --project`.

### Fixed

- `EXCEL_MCP_ALLOWED_URL_PREFIXES` parsing no longer uses `os.pathsep` on POSIX (colons in `https://` broke Linux/macOS CI and production allowlists).
