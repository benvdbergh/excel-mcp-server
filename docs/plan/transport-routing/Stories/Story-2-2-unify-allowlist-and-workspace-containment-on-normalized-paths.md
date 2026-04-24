---
kind: story
id: STORY-2-2
title: Unify allowlist and workspace containment on normalized paths
status: done
parent: EPIC-2
depends_on:
  - STORY-2-1
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/target-architecture.md
  - path: docs/architecture/pre-fork-architecture.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - Same normalized path is used for allowlist checks for file and future COM targets (FR-11).
  - Traversal and jail semantics for HTTP/SSE remain correct after refactor.
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-2-2: Unify allowlist and workspace containment on normalized paths

## Description

Generalize trust boundaries so **multiple workspace roots** or stricter stdio policy can apply consistently, and **COM** cannot attach to open workbooks outside policy (**FR-11**, target architecture §2).

## User story

As a **security-conscious operator**, I want **COM and file backends** to respect the **same path policy** so that **auto mode** cannot reach unexpected workbooks.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Preserve pre-fork behavior as default where PRD does not require stricter policy; document any tightening.
- Prepare for `workbook_open_in_excel` to filter by allowed paths before returning true.

## Delivered

- `src/excel_mcp/path_policy.py` — `EXCEL_MCP_ALLOWED_PATHS` (`os.pathsep`), `resolved_path_is_within`, `path_is_allowed`, `assert_path_allowed`, `allowlist_enforced`.
- `src/excel_mcp/server.py` — `get_excel_path` uses policy helpers; stdio uses `resolve_target` when allowlist active.
- `tests/test_path_allowlist.py` — stdio allowlist pass/fail, SSE jail ∩ allowlist, `path_is_allowed` export.
- `README.md` — operator documentation for `EXCEL_MCP_ALLOWED_PATHS` and path normalization overview.

## Dependencies (narrative)

Depends on **STORY-2-1** for normalized path output.
