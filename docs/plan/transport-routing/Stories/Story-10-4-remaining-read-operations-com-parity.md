---
kind: story
id: STORY-10-4
title: COM parity for remaining read contract operations
status: draft
parent: EPIC-10
depends_on:
  - STORY-10-2
traces_to:
  - path: docs/architecture/com-read-class-tools-design.md
  - path: src/excel_mcp/routing/com_workbook_service.py
  - path: src/excel_mcp/routing/workbook_operation_contract.py
  - path: src/excel_mcp/workbook.py
  - path: src/excel_mcp/sheet.py
  - path: src/excel_mcp/cell_validation.py
  - path: src/excel_mcp/validation.py
slice: vertical
invest_check:
  independent: false
  negotiable: true
  valuable: true
  estimable: false
  small: false
  testable: true
acceptance_criteria:
  - workbook_metadata, read_merged_cell_ranges, read_worksheet_data_validation, validate_sheet_range, and validate_formula_syntax on ComWorkbookService are implemented to match file JSON/string contracts or documented deltas.
  - Edge policies reuse existing COM helpers for protected view / read-only where consistent with FR-9 style errors (cross-ref Story-7-2 patterns).
  - Tests cover each operation at unit level (mock COM or thin harness); enumerate gaps for manual Excel validation in STORY-10-5 if needed.
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-10-4: COM parity for remaining read contract operations

## Description

Implement the **remaining read methods** on **`ComWorkbookService`** that are still stubs: **`workbook_metadata`**, **`read_merged_cell_ranges`**, **`read_worksheet_data_validation`**, **`validate_sheet_range`**, **`validate_formula_syntax`**. Follow **`FileWorkbookService`** delegation targets (`get_workbook_info`, `get_merged_ranges`, validation helpers) as semantic references per design note §4.3–4.7.

## User story

As an **operator**, I want **every read-class MCP tool** to function over **COM** when routing selects the Excel host, not only primary grid reads.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- **validate_formula_syntax** may reuse **pure** validation stacks where possible; align with **`validate_formula_in_cell_operation`** behavior from disk loads.
- **Data validation over COM** may differ subtly from openpyxl; prefer matching exported JSON schema with explicit field notes when parity is impossible.

## Dependencies (narrative)

Depends on **STORY-10-2**. Prefer coordinating shared COM workbook/sheet resolution helpers with **STORY-10-3** to avoid duplication.
