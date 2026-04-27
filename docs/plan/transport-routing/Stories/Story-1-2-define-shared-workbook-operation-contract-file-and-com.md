---
kind: story
id: STORY-1-2
title: Define shared workbook operation contract (File and COM)
status: done
parent: EPIC-1
depends_on:
  - STORY-1-1
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/target-architecture.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - Protocol or facade lists operations (e.g. read_range, write_range, apply_formula, metadata) matching agreed inventory; gaps explicitly marked follow-up.
  - Contract is reviewable in code (typed interface or module doc) and referenced by EPIC-3 facade tasks.
  - Automated acceptance tests live under a descriptive test module (see Verification); story docs reference that module by path, not by story id in the filename.
created: "2026-04-24"
updated: "2026-04-25"
---

# Story-1-2: Define shared workbook operation contract (File and COM)

## Description

Freeze the **internal method surface** both `FileWorkbookService` and `ComWorkbookService` will implement for routed operations (**FR-4**, **FR-5**), informed by the inventory from **STORY-1-1**.

## User story

As an **implementer**, I want a **stable internal API** so that **file extraction and COM parity** proceed in parallel without constant signature churn.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Include `tool_kind` / operation metadata hooks needed for future opt-in COM reads (**ADR 0003** consequences).
- Prefer operation-oriented names distinct from MCP wire `transport` (**ADR 0001**).

## Verification

- **Implementation:** [`src/excel_mcp/routing/workbook_operation_contract.py`](../../../../src/excel_mcp/routing/workbook_operation_contract.py)
- **Acceptance tests (descriptive name):** [`tests/test_shared_workbook_operation.py`](../../../../tests/test_shared_workbook_operation.py)

## Dependencies (narrative)

Depends on **STORY-1-1** for the classified inventory.
