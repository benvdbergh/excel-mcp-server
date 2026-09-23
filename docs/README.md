# Documentation map

Operator setup and the tool contract live in the repository root: [README](../README.md) and [TOOLS.md](../TOOLS.md). This folder holds architecture, specs, and the delivery record.

| Path | Use it for |
|------|------------|
| [architecture/target-architecture.md](architecture/target-architecture.md) | Current package layout, routing layers, and import rules |
| [architecture/README.md](architecture/README.md) | Index of architecture docs and ADRs |
| [architecture/adr/](architecture/adr/) | Decision records (workbook transport, COM-first, envelopes) |
| [specs/](specs/) | Product requirements: transport routing, table query and views |
| [operator/mcp-server-ids.md](operator/mcp-server-ids.md) | Cursor server ids and `mcp.json` keys |
| [performance/routing-nfr2-note.md](performance/routing-nfr2-note.md) | What is and is not benchmarked in CI |
| [plan/transport-routing/IMPLEMENTATION-ROADMAP.md](plan/transport-routing/IMPLEMENTATION-ROADMAP.md) | Epic and story delivery status |
| [index.html](index.html) | Short GitHub Pages splash; the README is the operator contract |

Story and epic files under `plan/transport-routing/` record where code lived when that story closed. After the package split, current paths are in [target-architecture.md](architecture/target-architecture.md), not in those evidence notes.

[pre-fork-architecture.md](architecture/pre-fork-architecture.md) and [excel-mcp-fork-com-vs-file-routing.md](excel-mcp-fork-com-vs-file-routing.md) are the baseline and the original fork blueprint.
