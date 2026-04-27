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

**pywin32** (`win32com.client`) is the chosen COM stack for this codebase: it ships as the optional **`[com]`** extra in `pyproject.toml`, keeps the default install free of Windows-only wheels, and pairs with the single-thread COM executor design (explicit apartment/threading discipline on the caller).

xlwings remains a documented alternative in the context table above but is **not** a default dependency.

The implemented choice is reflected in:

- `pyproject.toml` optional dependency group (e.g. `com`)
- README prerequisites for Windows
- Single-thread COM worker design (ADR cross-reference: target architecture § COM execution model)

## Consequences

- License and distribution notes belong in README / NOTICE if required by the chosen stack.
- All COM entry points stay behind import guards and `[com]` so Linux CI remains unaffected.
- Callers on **non-main threads** must follow Win32 COM rules: the dedicated COM worker initializes the apartment (`pythoncom.CoInitialize()` in `ComThreadExecutor`) so pywin32 calls match the single-thread design in target architecture §7.
