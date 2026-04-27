---
kind: story
id: STORY-6-1
title: Optional com extra and import guards in pyproject
status: done
parent: EPIC-6
depends_on:
  - STORY-1-2
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/architecture/adr/0002-com-automation-stack.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - pyproject.toml defines optional-dependencies com with chosen COM stack (ADR 0002 updated when choice is final).
  - Default install on Linux CI does not require COM packages (NFR-5, NFR-6).
created: "2026-04-24"
updated: "2026-04-27"
---

# Story-6-1: Optional com extra and import guards in pyproject

## Description

Add **`[com]`** optional dependencies and **import guards** so COM code never loads on unsupported platforms (**FR-12**, **NFR-5**).

## User story

As a **packager**, I want a **lightweight default install** so that **headless Linux deployments** stay simple.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Record license implications in README if pywin32 or xlwings requires notices.
- Coordinate version pins with security review.

## Dependencies (narrative)

Depends on **STORY-1-2** so COM module boundaries align with the shared contract.

## Delivered

- `pyproject.toml`: optional group `com` with `pywin32>=307`.
- `src/excel_mcp/com_support.py`: `COM_STACK_AVAILABLE`, `is_com_runtime_supported()`; no `win32com` import on non-Windows.
- ADR 0002 documents pywin32 as the chosen stack; README notes optional install and license.
