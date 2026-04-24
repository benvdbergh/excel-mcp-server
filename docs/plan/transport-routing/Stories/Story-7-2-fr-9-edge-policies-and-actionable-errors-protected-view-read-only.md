---
kind: story
id: STORY-7-2
title: FR-9 edge policies and actionable errors (protected view, read-only, duplicates)
status: draft
parent: EPIC-7
depends_on:
  - STORY-7-1
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
  - Protected view, read-only, and duplicate Excel instance cases return documented, stable error messages suitable for automation (ADR 0005 consequences).
  - Unsaved new workbook path policy matches PRD out-of-scope statement; errors are explicit, not silent.
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-7-2: FR-9 edge policies and actionable errors (protected view, read-only, duplicates)

## Description

Implement **ambiguous and failure cases** for COM routing: **protected view**, **read-only**, **duplicate instances**, and **unsaved path** handling (**FR-9**, PRD out-of-scope for `Book1`-style matching).

## User story

As an **operator**, I want **clear errors** when Excel cannot safely apply a mutation so that **agents do not corrupt** my session.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Prefer typed/domain exceptions mapped to MCP-visible errors.
- Document **foreground vs fail-closed** choice for duplicate instances in README.

## Dependencies (narrative)

Depends on **STORY-7-1** for baseline COM mutations.
