# Release management and versioning policy

**Last reviewed:** 2026-04-27

This document defines **semantic versioning**, **tagging**, **changelog / release notes**, and **automation options** for **`excel-com-mcp`** (this fork’s PyPI distribution). It mirrors the discipline of the reference **`workflows`** repository’s alpha release docs, adapted for **Python** and **PyPI** (`pyproject.toml` as the version source of truth).

## 1) Semantic versioning (SemVer)

The public **Python package version** lives in **`[project] version`** in `pyproject.toml` and must match the **intended** Git tag for that release (without the leading `v` on the tag, the numeric part should align, e.g. tag `v0.1.9` → version `0.1.9`).

### Pre-1.0 (`0.y.z`)

While the project remains **pre-1.0**:

- **Breaking changes do not require jumping to `1.0.0`.** Consumers should still expect evolution.
- Suggested bump semantics (maintainer discretion):
  - **`0.(y+1).0` (minor):** breaking or externally significant behavior changes, new major surface area, or incompatible defaults.
  - **`0.y.(z+1)` (patch):** backward-compatible fixes, documentation that affects usage, low-risk improvements.

### Release channels on PyPI (no npm dist-tags)

PyPI does not offer npm-style **dist-tags** on the same version. Use one or both of:

1. **PEP 440 pre-releases** — e.g. `0.2.0a1`, `0.2.0b1`, `0.2.0rc1` for iteration cuts; `pip` resolves stable releases unless a pre-release is explicitly requested or pinned.
2. **TestPyPI** — separate index for dry-run publishing and integrators who opt in.

Iteration vs “accepted baseline” can still be expressed **in Git** with tags such as `v0.y.z-alpha.N` for candidates and `v0.y.z` when that line is accepted, while the **uploaded** version strings follow PEP 440 (see [ci-cd-packaging-governance.md](ci-cd-packaging-governance.md) for workflow separation).

## 2) Conventional Commits as release intent

Use [Conventional Commits](https://www.conventionalcommits.org/) as **hints** for the next version bump. If a release contains mixed commits, apply the **highest** required bump.

| Commit type | Typical bump (pre-1.0) | Notes |
|-------------|-------------------------|--------|
| `feat` | minor (`0.y+1.0`) or patch if explicitly non-breaking small surface | Default: minor for user-visible capability. |
| `fix` | patch | User-visible defect correction. |
| `feat!`, `fix!`, `BREAKING CHANGE:` footer | minor | Breaking stays within `0.y.z` until 1.0.0 policy changes. |
| `perf` | patch or none | Bump when externally observable. |
| `refactor` | none by default | Bump only if behavior changes for consumers. |
| `docs`, `test` | none | Unless release notes policy says otherwise. |
| `build`, `ci`, `chore`, `style` | none | Internal-only unless it affects artifacts or runtime. |

## 3) Final release commit checklist

Before tagging or publishing:

1. **Working tree** — only intended files changed.
2. **Quality gates** (local or CI): full **`pytest`** suite; **`hatch build`**; **`twine check dist/*`** where applicable.
3. **Version** — `pyproject.toml` version matches the intended tag and PyPI slot.
4. **Optional governed packaging** — run **Release packaging** manual workflow (once present) with `release_ref` set to the candidate tag or SHA; inspect artifacts.
5. **Optional governed publish** — run **Release PyPI publish** manual workflow (once present) with explicit `release_ref` and environment (PyPI vs TestPyPI).
6. **Release narrative** — update changelog / GitHub Release description (template below).
7. **Tag** — push annotated or signed tags per team practice (`v0.y.z` or `v0.y.z-alpha.N`).

## 4) Changelog and release notes

### Process

1. Collect merged changes since the previous tag.
2. Group under **Added**, **Changed**, **Fixed**, **Docs**, **Internal**.
3. Call out **Breaking / impact** for consumers even in `0.y.z`.
4. Attach the same summary to the **GitHub Release** body when using `release: published` automation.

### Template block

```markdown
## 0.y.z - YYYY-MM-DD

### Added
- ...

### Changed
- ...

### Fixed
- ...

### Docs
- ...

### Internal
- ...

### Breaking / impact notes
- None.

### Validation run
- `pytest`
- `hatch build`
- `twine check dist/*` (if used)
```

## 5) Automation options

Pick one primary model; all can satisfy the same governance if checks stay centralized.

| Model | Behavior | Fit |
|-------|----------|-----|
| **Manual (closest to `workflows`)** | Maintainer bumps `pyproject.toml`, updates changelog, pushes tag, dispatches **manual publish** workflow with `release_ref`. | Maximum control; minimal automation. |
| **GitHub Release triggers publish** | Keep **`release: published`** workflow; ensure it **calls the same reusable validation** as PR CI before `gh-action-pypi-publish`. | Familiar for OSS; requires discipline on release notes. |
| **release-please** | Opens release PRs; manages changelog and version bumps; tag or release triggers publish. | Good GitHub-native automation for Python. |
| **semantic-release** | Fully commit-driven bumps and changelog from conventional commits. | Strong convention discipline required. |

**Recommendation:** start with **reusable CI + manual dispatch publish** (or hardened **release-published** path), then add **release-please** if release volume grows.

**Active automation in this repository**

| Trigger | Workflow | Role |
|---------|----------|------|
| **`push` / `pull_request`** to `main` / `master` | **`ci.yml`** | Calls **`reusable-validate-and-test.yml`** (PR quality gate; stable job name **`validate-and-test`**). |
| **`release: published`** | **`publish.yml`** | Trusted publish to PyPI after the same reusable validation; GitHub Environment **`release`**. |
| **`workflow_dispatch`** | **`release-packaging.yml`** | Build **`dist/*`** from **`release_ref`** for inspection. |
| **`workflow_dispatch`** | **`release-pypi-publish.yml`** | Manual trusted publish from **`release_ref`** to **PyPI** or **TestPyPI** (`pypi_target`). |

This table is the concrete mapping for §5’s “manual + release-published” model; adjust branch protection required checks to the displayed **`CI / validate-and-test`** (or equivalent) name in GitHub once workflows are enabled.

## 6) CI/CD governance cross-reference

Workflow names, permissions, artifact retention, and PyPI OIDC details: **[ci-cd-packaging-governance.md](ci-cd-packaging-governance.md)**.

## 7) Maintainer decisions to keep explicit

- Criteria for moving from **pre-releases** to the next **patch/minor** stable line on PyPI.
- Whether **TestPyPI** is mandatory for every release or only for risky changes.
- Date or quality bar for declaring **`1.0.0`** and tightening breaking-change rules to SemVer major bumps.

## 8) Related documents

- [ci-cd-packaging-governance.md](ci-cd-packaging-governance.md)
- [target-architecture.md](target-architecture.md)
- [adr/README.md](adr/README.md)
- Implementation plan: `docs/plan/transport-routing/Epics/Epic-8-governed-ci-cd-pypi-and-release-pipelines.md`
