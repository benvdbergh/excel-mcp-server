---
kind: story
id: STORY-7-5
title: README, CI, NFR-2 note, and manual Windows checklist
status: draft
parent: EPIC-7
depends_on:
  - STORY-7-2
  - STORY-7-3
  - STORY-7-4
traces_to:
  - path: docs/specs/PRD-excel-mcp-transport-routing.md
  - path: docs/excel-mcp-fork-com-vs-file-routing.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - README documents transport matrix, env vars, threading, save policy, Windows prerequisites, and MCP vs workbook transport distinction (PRD AC4, ADR 0001).
  - Default CI passes without Excel; COM paths mocked (NFR-6).
  - Manual checklist from blueprint §7 is linked or reproduced under docs/ and executed once before RC sign-off (PRD AC5, Gate 3).
  - NFR-2 p95 routing overhead captured as benchmark note or automated timing harness where feasible.
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-7-5: README, CI, NFR-2 note, and manual Windows checklist

## Description

Close **release-level acceptance**: documentation, **CI strategy**, **performance note**, and **manual validation** on Windows with Excel (**PRD** acceptance criteria, **Gate 3**).

## User story

As a **maintainer**, I want **documented operations and CI** so that **contributors and users** know how to install, test, and run COM features safely.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Include **NFR-4** security note (no admin elevation by default).
- Cross-link **IMPLEMENTATION-ROADMAP.md** from README optional—keep single source of truth for planning under `docs/plan/transport-routing/`.

## Dependencies (narrative)

Depends on **STORY-7-2**, **STORY-7-3**, and **STORY-7-4** so documented behavior matches implemented edge policies, chart/pivot scope, and save tool.
