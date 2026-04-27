---
kind: story
id: STORY-8-4
title: Manual PyPI publish and harden release-published path
status: draft
parent: EPIC-8
depends_on:
  - STORY-8-2
traces_to:
  - path: docs/architecture/ci-cd-packaging-governance.md
  - path: docs/architecture/release-versioning-policy.md
  - path: README.md
  - path: manifest.json
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - "`workflow_dispatch` publish workflow exists with `release_ref` (and optional TestPyPI vs PyPI selection via environment or input), calls reusable validate-and-test first, then uses `pypa/gh-action-pypi-publish`."
  - "Only the publish job sets `id-token: write`; publish uses pinned Python consistent with CI."
  - "Existing `publish.yml` on `release: published` either invokes the same reusable validation before publish or is replaced by a pattern documented in governance §3."
  - "Release versioning doc §5 (automation options) references which triggers are active after merge."
  - "README documents no-install MCP wiring for agentic hosts (e.g. Cursor, Claude Code, comparable MCP clients) using the published distribution identity: commands and JSON examples use the same package name as `pyproject.toml` `[project].name` and the stdio entrypoint; `manifest.json` `server.mcp_config` matches that command/args for registry-oriented installs."
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-8-4: Manual PyPI publish and harden release-published path

## Description

Add a **manual trusted publish** path and ensure **GitHub Release → PyPI** (if retained) never skips quality gates, per [ci-cd-packaging-governance.md](../../../architecture/ci-cd-packaging-governance.md) §6–§7 and [release-versioning-policy.md](../../../architecture/release-versioning-policy.md) §5.

## User story

As a **maintainer**, I want **publish from an explicit ref with OIDC** so that **PyPI uploads are gated, auditable, and consistent with PR CI**.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- This closes the loop from **PyPI artifact** to **documented consumer config** so “publish succeeded” implies operators can copy from README (and optional manifest) without name drift after a fork/rename.
- Coordinate **GitHub Environment** name (`release` / `pypi`) with PyPI **trusted publisher** settings.
- **STORY-8-3** is optional before this story; publish can rebuild from ref. Prefer validate → build → publish in one workflow if artifact reuse is not required for v1.

## Dependencies (narrative)

Requires **STORY-8-2**. May proceed in parallel with **STORY-8-3** once reusable workflow exists.
