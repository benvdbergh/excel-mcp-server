# Changelog

All notable changes to this project are documented in this file. The format is informal; align version bumps with [Semantic Versioning](https://semver.org/) and [docs/architecture/release-versioning-policy.md](docs/architecture/release-versioning-policy.md).

## [Unreleased]

## 0.5.0 — 2026-06-23

Epics **1–4** (agent reliability, read fidelity, hardening, operator documentation): BEN-120 through BEN-140.

### Added

- **Routing metadata envelope (ADR 0010, BEN-120):** Optional `include_routing_metadata` on `read_data_from_excel`. When `true`, successful responses wrap the tool payload in `{ "result", "_meta", "warnings" }` with `workbook_transport`, `workbook_backend`, `routing_reason`, and `duration_ms`. Default `false` preserves legacy JSON parsers. Supersedes BEN-138 (duplicate).
- **SharePoint open-detection (BEN-131):** Shared `workbook_host_identity` helpers align `FullName` normalization between COM open-detection and attach paths (ADR 0006); Protected View excluded from `auto`→COM until Enable Editing.
- **Display-text reads (BEN-125, BEN-126):** COM `Range.Text` path when `value_mode=text`; optional `value_mode` on `read_data_from_excel` (`value` \| `text`, default `value`). Response echoes `value_mode` at root.
- **Bulk read tool (BEN-130):** `export_worksheet_table` — compact header row + data matrix with optional `max_rows` cap and `truncated` flag; routes on file and COM backends.
- **COM recalc tool (BEN-129):** `evaluate_range` forces Excel recalculation on sheet or range before reads; COM-only, does not persist to disk.
- **Compact reads (BEN-136):** Optional `metadata_mode` (`full` \| `compact`) on `read_data_from_excel`; `compact` omits per-cell validation metadata (default `full`).
- **Discovery detail levels (BEN-140):** `excel_list_open_workbooks` implements `detail` (`minimal` \| `active_context`); `active_context` adds active workbook, sheet, and selection (absorbs BEN-127).

### Changed

- **COM Value2 resiliency (BEN-133):** Wider sampling for large-range sparse `Value2` anomalies before per-cell direct-read fallback.

### Breaking

- **Schema cleanup (BEN-135):** Removed unused `preview_only` parameter from read tools. Clients that still pass `preview_only` must drop it.

### Fixed

- **`.xlsm` file-backend reads (BEN-134):** When the file backend reads formula cells on `.xlsm`, emit ADR 0010 warning `file_backend_formula_not_evaluated` (via `include_routing_metadata` envelope).

### Docs

- **ADR 0010 (BEN-151):** MCP tool response envelope contract.
- **Operator documentation (BEN-132, BEN-137):** `read_data_from_excel` docstring and README hero aligned with COM-first routing, SharePoint `https://` locators, and discovery workflow.
- **Install decision matrix (BEN-121):** README documents local fork vs pinned PyPI vs consumer workspace; global unpinned `uvx excel-com-mcp@latest` marked legacy with migration snippet.
- **TOOLS.md** — `include_routing_metadata`, `_meta`, `value_mode`, `metadata_mode`, `export_worksheet_table`, `evaluate_range`, and `excel_list_open_workbooks` `detail`.

### Tests

- **SharePoint + formatted cells (BEN-139):** Mocked integration tests for `https` FullName routing and COM read fidelity (CI-safe without live Excel).

### Operator

- **Ai-Vault consumer pin (BEN-150):** Verified `excel-local` in `Ai-Vault/.cursor/mcp.json` pins `--project C:/Users/vandenbb/mcp/excel-mcp-server` with `EXCEL_MCP_ALLOWED_URL_PREFIXES` for `https://kion.sharepoint.com/` — no config change required for this release.

## 0.4.1 — 2026-05-08

### Fixed

- **COM range read resiliency:** `read_data_from_excel` on COM transport now detects sparse/blank anomalies from bulk `Range.Value2` and falls back to direct per-cell reads when mismatches are detected. This addresses cases where grouped/outlined sheet regions returned mostly `null` values despite populated cells.

### Tests

- Added COM regression coverage for fallback behavior in `tests/test_com_workbook_service.py`.

## 0.4.0 — 2026-04-27

### Added

- **Open workbook discovery ([ADR 0009](docs/architecture/adr/0009-open-workbook-discovery-tool.md)):** MCP tool **`excel_list_open_workbooks`** returns JSON enumerating **`Application.Workbooks`** (`full_name`, `name`, `is_active`), COM-only on the executor thread. Use each returned **`full_name`** as **`filepath`** on **`get_workbook_metadata`**, reads, and writes.

### Docs

- README (upgrade notes), **`TOOLS.md`**, **`manifest.json`**, [`MANUAL-WINDOWS-RC-CHECKLIST.md`](docs/plan/transport-routing/MANUAL-WINDOWS-RC-CHECKLIST.md), [`IMPLEMENTATION-ROADMAP.md`](docs/plan/transport-routing/IMPLEMENTATION-ROADMAP.md): discovery workflow, Windows checklist, epic status; [`release-versioning-policy.md`](docs/architecture/release-versioning-policy.md) last reviewed.

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
