---
kind: epic
id: EPIC-2
title: Path normalization and unified allowlist
status: done
depends_on:
  - EPIC-1
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/target-architecture.md
  - path: docs/architecture/pre-fork-architecture.md
slice: vertical
acceptance_criteria:
  - resolve_target (or equivalent) returns a documented normalized absolute path used consistently for file, allowlist, and COM comparison (FR-1, US-4).
  - Allowlist / workspace containment uses the same normalized path for stdio and HTTP/SSE modes (FR-11).
created: "2026-04-24"
updated: "2026-04-25"
---

# Epic-2: Path normalization and unified allowlist

## Delivery summary

Epic acceptance criteria are satisfied in code and tests:

| Theme | Where |
|-------|--------|
| `resolve_target`, FR-1 / US-4 | `src/excel_mcp/path_resolution.py`, `tests/test_resolve_target.py`, `tests/test_path_equivalence.py` |
| Unified allowlist / jail, FR-11 | `src/excel_mcp/path_policy.py`, `get_excel_path` in `src/excel_mcp/server.py`, `tests/test_path_allowlist.py` |
| Operator docs | `README.md` — path normalization and `EXCEL_MCP_ALLOWED_PATHS` |

Stories: [2-1](../Stories/Story-2-1-implement-resolve-target-with-documented-relative-path-order.md), [2-2](../Stories/Story-2-2-unify-allowlist-and-workspace-containment-on-normalized-paths.md), [2-3](../Stories/Story-2-3-path-equivalence-automated-tests-casing-junctions-short-vs-long.md) — all **done**.

## Description

Implement **`resolve_target`** and extend trust boundaries so **one path pipeline** feeds file I/O, policy checks, and future `workbook_open_in_excel` comparisons—without yet changing workbook transport behavior beyond safer, documented resolution.

## Objectives

- Meet **FR-1** (relative path resolution order documented and implemented).
- Meet **US-4** / **NFR-1** inputs via normalization tests (junctions, casing, short vs long paths where applicable on Windows).
- Align **FR-11** with target architecture §2 (multiple roots / unified checks).

## User stories (links)

- [Story-2-1](../Stories/Story-2-1-implement-resolve-target-with-documented-relative-path-order.md)
- [Story-2-2](../Stories/Story-2-2-unify-allowlist-and-workspace-containment-on-normalized-paths.md)
- [Story-2-3](../Stories/Story-2-3-path-equivalence-automated-tests-casing-junctions-short-vs-long.md)

## Dependencies (narrative)

Depends on **EPIC-1** so path inputs and tool entry points are known and stable.

## Related sources

- `docs/architecture/target-architecture.md` — §1 `resolve_target`, §2 allowlist.
- `docs/architecture/pre-fork-architecture.md` — `get_excel_path` baseline.
