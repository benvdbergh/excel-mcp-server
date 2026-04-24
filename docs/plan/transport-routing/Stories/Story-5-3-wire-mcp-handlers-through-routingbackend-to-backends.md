---
kind: story
id: STORY-5-3
title: Wire MCP handlers through RoutingBackend to backends
status: draft
parent: EPIC-5
depends_on:
  - STORY-4-3
  - STORY-5-2
traces_to:
  - path: docs/architecture/target-architecture.md
slice: vertical
invest_check:
  independent: true
  negotiable: false
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - Handlers obtain resolved path, effective transport, and dispatch via RoutingBackend to FileWorkbookService for executed operations in this phase.
  - Integration-style tests (still mock COM) cover at least one write and one read path through the router.
created: "2026-04-24"
updated: "2026-04-24"
---

# Story-5-3: Wire MCP handlers through RoutingBackend to backends

## Description

Refactor **`server.py`** tool handlers to use **`RoutingBackend`** as the single dispatch gate (target architecture layered view), preserving backward compatibility for callers that omit new parameters.

## User story

As a **maintainer**, I want **all routed tools** to pass through **one dispatch layer** so that **logging and transport rules** cannot be bypassed accidentally.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Read-class tools remain file-backed per **ADR 0003** but still log routing context if applicable.
- Preserve existing error string patterns where PRD does not require change; document breaking changes if any.

## Dependencies (narrative)

Depends on **STORY-5-2** (parameters) and **STORY-4-3** (logging on router).
