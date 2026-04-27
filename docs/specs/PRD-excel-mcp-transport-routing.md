---
title: Product Requirements Document
type: prd
product: Excel MCP Server (fork) — File vs COM transport routing
project: excel-com-mcp
version: 0.1
status: draft
created: 2026-04-24
updated: 2026-04-27
source_blueprint: docs/excel-mcp-fork-com-vs-file-routing.md
scale_profile: growth
---

# Product Requirements Document

## Executive Summary

This PRD defines product requirements for a **forked Excel MCP server** that routes each workbook operation to either **file-based I/O** (existing behavior) or **COM-based automation** (Excel host) depending on whether the target workbook is **already open in a local Microsoft Excel instance** on Windows. The outcome is fewer **last-writer-wins / cloud merge surprises** when agents edit workbooks the user has open in Excel, while preserving **headless and CI-friendly** file transport when Excel does not own the session.

**Primary deliverable:** a `RoutingBackend` with explicit configuration (`auto` | `file` | `com`), path-normalized open-workbook detection, shared internal service API across `FileWorkbookService` and `ComWorkbookService`, and backward-compatible tool schemas.

## Product Vision

Agents and IDE integrations interact with Excel workbooks through MCP tools in a way that **respects the user’s active Excel session** when one exists: visible workbooks update live, saves follow Excel’s normal persistence and sync behavior. When no such session exists, the same tools continue to operate reliably via **direct file access** without requiring Excel to be running.

## Problem Statement and Opportunity

**Problem**

- File-based MCP writes can diverge from what the user sees in **Excel desktop or a browser-backed session**, increasing risk of OneDrive merge conflicts, stale UI, and unexpected overwrites.
- COM-based editing aligns with user mental models for “live” spreadsheets but is inappropriate for batch, headless, or non-Windows environments.

**Opportunity**

- Introduce a **router** in front of existing implementations so upstream file logic remains a **backend module** with minimal churn, while adding a **COM backend** that implements the same internal contract.
- Offer **explicit transport controls** and observability so power users and operators can predict behavior and debug routing decisions.

**Baseline (success relative to today)**

- Today: all mutating operations use a single transport (typically file). After delivery: `transport=auto` selects COM when the resolved workbook path matches an open Excel workbook, else file—without breaking existing callers that omit the new optional parameters.

## Target Users

| Persona | Needs |
|---------|--------|
| **IDE / agent operator** | Predictable edits; optional “live Excel” path; clear errors when COM is blocked (protected view, read-only). |
| **Developer / integrator** | Stable tool contracts; flags for `transport` and save policy; logs for support. |
| **Maintainer / OSS contributor** | Clear layering (`FileWorkbookService`, `ComWorkbookService`, `RoutingBackend`); tests that run without Excel on CI. |

## User Stories

| ID | Story | Priority |
|----|--------|----------|
| US-1 | As an operator, when my workbook is **open in Excel** and I use a mutating MCP tool with default settings, I want edits to apply **through Excel (COM)** so the UI updates immediately. | P0 |
| US-2 | As an operator, when the workbook is **not** open in Excel, I want the MCP to use **file-based** I/O so automation works headlessly and in CI. | P0 |
| US-3 | As a power user, I want `EXCEL_MCP_TRANSPORT` and per-call `transport` so I can force **file** or **com** or keep **auto**, avoiding silent mis-routing. | P0 |
| US-4 | As an operator, I want path matching to treat **OneDrive / junction / short vs long paths** consistently so the correct backend is chosen. | P1 |
| US-5 | As a power user, I want optional **`EXCEL_MCP_COM_STRICT=1`** so that when I expect COM, the server **errors** instead of falling back to file if the workbook is not found in Excel. | P2 |
| US-6 | As an integrator, I want **structured logs** (transport chosen, reason, duration) for every routed operation to troubleshoot production issues. | P1 |

### User story acceptance criteria (representative)

- **US-1 AC:** Given Windows + Excel installed, given workbook W open in Excel, given `transport=auto` (or unset defaulting to auto), when a supported mutating tool targets W’s resolved path, then the COM backend is used and the visible Excel grid reflects the change within the same user session without requiring a manual file reload.
- **US-2 AC:** Given workbook W not listed in `Application.Workbooks` (or Excel not running), given `transport=auto`, when a mutating tool targets W, then the file backend is used and the `.xlsx` on disk reflects the change; no COM connection is required.
- **US-3 AC:** Given `transport=file`, any open-in-Excel state yields file backend only. Given `transport=com` and workbook not open in Excel, the tool returns a **documented error** (unless an explicitly documented optional fallback flag is enabled).
- **US-4 AC:** Given two path strings that resolve to the same file identity (per documented normalization rules, optionally file ID), routing treats them as the same workbook for open detection.

## Functional Requirements

| ID | Requirement | Priority |
|----|-------------|----------|
| FR-1 | Provide `resolve_target(path)` (or equivalent) that returns a **normalized absolute path** for routing and comparison, with documented order for **relative path** resolution (`cwd`, then documented vault/workspace roots). | P0 |
| FR-2 | Implement `workbook_open_in_excel(resolved_path) -> bool` on **Windows** using running `Excel.Application` and enumeration of open workbooks; compare **resolved full paths** per normalization rules in the blueprint (short/long, drive casing, symlink/junction targets). | P0 |
| FR-3 | Provide `RoutingBackend` / `resolve_workbook_backend(path, transport_mode)` that implements the **transport selection matrix** in the blueprint (`auto`, `file`, `com`, optional com fallback flag). | P0 |
| FR-4 | Extract or maintain **file I/O** behind `FileWorkbookService` with a **narrow shared API** (e.g., `read_range`, `write_range`, `apply_formula`, `get_metadata`—exact surface to match upstream inventory in implementation plan). | P0 |
| FR-5 | Add `ComWorkbookService` (Windows-gated) implementing the **same method signatures** as `FileWorkbookService` for all **routed** tools agreed in the tool inventory. | P0 |
| FR-6 | **Serialize COM calls** on a single thread (dedicated executor or lock) compatible with async MCP servers; document threading model. | P0 |
| FR-7 | Expose optional tool parameters: `transport: "auto" \| "file" \| "com"` (default `"auto"`), and optional `save_after_write: boolean` for COM path; file path behavior documented when ignored. | P0 |
| FR-8 | Default **COM write persistence:** `auto_save=false` at host level unless `save_after_write=true` (product default aligns with blueprint). | P1 |
| FR-9 | **Ambiguous cases** follow documented policies: duplicate instances (foreground / fail-closed), unsaved new book (`Saved == False`, unstable path) treated as **not routable by path** until a future enhancement, read-only/protected view surfaced as **clear errors**. | P1 |
| FR-10 | Do **not** auto-start Excel for routing by default; starting Excel (including invisible instances) is **opt-in only** if ever offered. | P0 |
| FR-11 | Preserve or introduce **path allowlist / workspace root** options consistent with upstream trust model; document that COM can affect **any open workbook** matching allowed paths. | P1 |
| FR-12 | On **non-Windows** platforms, COM module is skipped; router defaults to file for `auto`/`file`; `com` yields a clear **unsupported** error. | P0 |

### Tool routing expectations (from blueprint)

Implementation planning must map each MCP tool to **read vs write** routing; minimum product intent:

- **Write-class tools** (e.g., `write_data_to_excel`, `apply_formula`, formatting/merge/row-col operations): **must** participate in routing per FR-3–FR-5.
- **Read-class tools** may remain file-first for performance unless product chooses COM parity; if so, document trade-offs in ADR.

Traceability: detailed per-tool table lives in `docs/excel-mcp-fork-com-vs-file-routing.md` §6; implementation shall not drop tools from inventory without explicit PRD amendment.

## Non-Functional Requirements

| ID | Category | Requirement | Verification |
|----|----------|-------------|--------------|
| NFR-1 | Reliability | For `transport=auto`, incorrect backend selection rate **0** on golden path tests (mocked COM + deterministic paths) for defined equivalence classes. | Automated test suite |
| NFR-2 | Performance | Open-workbook detection + backend selection (excluding first-time Excel COM cold start) completes in **p95 ≤ 500 ms** on a reference Windows dev profile with ≤10 open workbooks. | Benchmark / manual note in test docs |
| NFR-3 | Observability | Each routed call emits **structured log** fields: `transport`, `reason` (e.g., `full_name_match`, `file_identity_match`, `forced_file`, `forced_com`), `duration_ms`, `workbook_path` (redacted if policy requires). | Log inspection |
| NFR-4 | Security | No elevation: MCP must not start or attach to Excel **as administrator** without a separate explicit user opt-in outside default install. | Code review + config docs |
| NFR-5 | Compatibility | Default install remains **lightweight**: COM dependencies optional via `extras_require` or equivalent packaging (`[com]`), documented in README. | Packaging CI |
| NFR-6 | Testability | **CI (no Excel):** router + file backend tests pass; COM paths covered by **mocks**; optional **manual** integration cases documented for Windows + Excel. | CI config + test markers |

## Acceptance Criteria (release-level)

1. With `EXCEL_MCP_TRANSPORT=auto` (default), integration tests (mocked) prove routing to **COM** when `workbook_open_in_excel` is true and to **file** when false.
2. Forced modes: `transport=file` always uses file; `transport=com` uses COM or returns documented error when not open (no silent file write unless explicitly documented fallback is enabled).
3. Relative and junction-heavy paths: at least one automated test demonstrates normalization behavior per documented rules.
4. README documents: transport matrix, env vars (`EXCEL_MCP_TRANSPORT`, `EXCEL_MCP_COM_STRICT`), threading/save policy, and Windows/COM prerequisites.
5. Manual test checklist from blueprint §7 is reproduced or linked in `docs/` and executed once before marking release-ready for COM features.

## Success Metrics

| Metric | Target | Horizon |
|--------|--------|---------|
| Routed write operations using COM when workbook open (`auto`) | **100%** on golden manual scenarios (§7 checklist) | First release candidate |
| Support tickets / issues citing “Excel didn’t update” or “merge conflict from MCP” | **Downward trend** vs pre-routing baseline (qualitative until baseline exists) | 30 days post dogfood |
| CI green without Excel | **100%** runs on default pipeline | Ongoing |
| p95 routing overhead | **≤ 500 ms** (NFR-2) | Release |

## Out of Scope

- macOS Excel COM parity.
- Driving **LibreOffice** via UNO.
- Replacing Graph API / SharePoint-only workflows.
- Matching unsaved new workbooks (`Book1`) by window title or heuristics (optional future; explicitly excluded from v1 unless reopened).
- Office Add-in + WebSocket bridge (listed as alternative only).

## Dependencies

| Dependency | Impact |
|------------|--------|
| Upstream **user-excel** (or equivalent) MCP codebase fork | Inventory of tools and file I/O boundaries (blueprint Step A). |
| **Windows + Microsoft Excel** installed | COM integration tests (manual); runtime for COM path. |
| **xlwings** or **pywin32** | COM implementation choice; affects packaging and license notes. |
| Python packaging (`extras_require` / optional deps) | Keeps default install small (NFR-5). |

## Risks, Assumptions, and Open Questions

| Item | Type | Mitigation / owner |
|------|------|-------------------|
| OneDrive path aliasing / sync delays | Risk | Path normalization + optional file identity; dogfood; document limitations. |
| COM threading / apartment issues | Risk | Single-thread executor; tests; ADR for chosen library. |
| `read` tools stay file-based while Excel shows different calculated state | Risk / open | Decide per-tool COM parity in architecture; document agent-facing caveats. |
| Duplicate Excel instances | Risk | Policy: foreground preference or fail-closed; test. |
| Protected View / read-only | Risk | Detect and return actionable errors (FR-9). |

**Unresolved decisions (escalate to architecture / implementation plan)**

- Exact shared interface surface after upstream inventory (FR-4).
- Whether **read** tools default to COM when open (blueprint marks optional).
- Optional `allow_com_fallback` naming and exposure on tools vs env only.

## Quality Gates

| Gate | Status intent |
|------|----------------|
| **Gate 1 — Spec ready** | This PRD: measurable goals, testable FR/NFR, explicit out-of-scope, risks/deps documented. **Pass** when stakeholders agree scope and priorities P0/P1. |
| **Gate 2 — Build ready** | Phased plan exists (inventory → file extract → COM service → router → schema); ADRs for COM library and read-tool parity; test strategy mapped to AC. **Track in technical plan.** |
| **Gate 3 — Release ready** | README + ops notes; observability fields live; manual Windows checklist completed for RC; optional extras install documented. |

## Timeline

Phasing aligns with blueprint §5 (inventory → COM module → router → schema → security/ops). Specific calendar dates are owned by **project-planning** after dependency on fork baseline is confirmed.

## Notes

- **Source:** Requirements trace to `docs/excel-mcp-fork-com-vs-file-routing.md` (architecture diagram, matrix, policies, testing, rollout).
- **Rollout:** Ship with `transport` default `auto`, log chosen backend (per blueprint §8); optional upstream contribution path noted as non-blocking.
- **Handoff:** `project-planning` should decompose FR-1–FR-12 and US-1–US-6 into epics/stories preserving IDs; `software-architecture` should own ADRs for COM stack, threading, and read-path parity.

## Engineering progress (fork implementation)

Phased delivery is tracked in [`docs/plan/transport-routing/IMPLEMENTATION-ROADMAP.md`](../plan/transport-routing/IMPLEMENTATION-ROADMAP.md). **Epics 1–6** are **implemented in code** (including optional `[com]` / pywin32, `ComThreadExecutor`, `ComWorkbookService` skeleton, and routed COM dispatch for supported writes). **Epic 7** (broad COM write parity, `save_workbook` MCP tool, docs/CI/release hardening) remains **planned**.

## References

- Internal: `docs/excel-mcp-fork-com-vs-file-routing.md`
- External: [How to manage merge conflicts in Excel cloud files](https://support.microsoft.com/en-us/office/how-to-manage-merge-conflicts-in-excel-cloud-files-535fb3f2-e7c9-4701-bdcd-0c447d284a6f) (context)
