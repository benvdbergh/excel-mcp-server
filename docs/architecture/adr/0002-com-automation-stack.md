# ADR 0002: COM automation stack (Windows Excel)

## Status

Accepted

## Context

Implementing `ComWorkbookService` requires driving Excel via COM on Windows. Common options:

| Option | Pros | Cons |
|--------|------|------|
| **pywin32** (`win32com.client`) | Direct control, widely used, no extra runtime | Verbose API, apartment/threading discipline on the caller |
| **xlwings** | Ergonomic Python surface, aligns with Excel object model | Additional dependency; license and deployment story must match the project |

The PRD lists both as candidates (`docs/specs/PRD-excel-mcp-transport-routing.md` Dependencies).

## Decision

**Record the choice here when implementation starts** — default recommendation for this codebase:

- Prefer **pywin32** for a **minimal optional extra** and explicit control inside the single-thread COM executor, **unless** the team values xlwings’ ergonomics enough to accept the extra dependency.

The implemented choice must be reflected in:

- `pyproject.toml` optional dependency group (e.g. `com`)
- README prerequisites for Windows
- Single-thread COM worker design (ADR cross-reference: target architecture § COM execution model)

## Consequences

- License and distribution notes belong in README / NOTICE if required by the chosen stack.
- All COM entry points stay behind import guards and `[com]` so Linux CI remains unaffected.
