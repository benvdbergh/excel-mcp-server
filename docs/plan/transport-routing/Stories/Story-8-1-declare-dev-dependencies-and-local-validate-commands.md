---
kind: story
id: STORY-8-1
title: Declare dev dependencies and local validate commands
status: done
parent: EPIC-8
depends_on: []
traces_to:
  - path: docs/architecture/ci-cd-packaging-governance.md
  - path: docs/architecture/release-versioning-policy.md
slice: vertical
invest_check:
  independent: true
  negotiable: true
  valuable: true
  estimable: true
  small: true
  testable: true
acceptance_criteria:
  - "`[project.optional-dependencies] dev` (or PEP 621 dependency-groups) lists pytest and any linters/build tools CI will invoke."
  - "Contributors can install dev deps in one documented command (e.g. `pip install -e \".[dev]\"` or uv equivalent)."
  - "`pytest` passes locally with only dev + default extras (no Windows COM requirement for default suite)."
  - "Governance doc §9 gap (undeclared pytest) is closed by this change or explicitly deferred with issue link."
created: "2026-04-27"
updated: "2026-04-27"
---

# Story-8-1: Declare dev dependencies and local validate commands

## As delivered

- **`[project.optional-dependencies].dev`** in `pyproject.toml` includes **pytest**, **twine**, **hatch** (and aligns with CI install **`pip install -e ".[dev]"`**).
- **README** § *Development and CI parity* documents **`pytest`**, **`hatch build`**, **`twine check dist/*`** (plus **`uv sync --extra dev`** path).
- **`docs/architecture/ci-cd-packaging-governance.md` §9** documents closed dev-deps gap for this fork.

## Description

Close the **reproducibility gap** between contributors, future CI, and release gates by declaring **test and build-check** dependencies in `pyproject.toml` and documenting the **local validate** command sequence aligned with [ci-cd-packaging-governance.md](../../../architecture/ci-cd-packaging-governance.md) (pytest, hatch build, twine check).

## User story

As a **contributor**, I want **declared dev dependencies and documented commands** so that **my local results match CI and release** without guessing installed packages.

## Acceptance criteria

See frontmatter `acceptance_criteria`.

## Technical notes

- Prefer **`dev`** optional extra over undocumented `pip install pytest` in README only.
- If optional **`twine`** / **`build`** are added only for CI, document that in Story-8-2 workflow or keep them in `dev` for local parity.

## Dependencies (narrative)

None. Unblocks **Story-8-2**.
