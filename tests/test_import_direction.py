"""AST-based import-direction rules for excel_mcp package layers.

Parses ``.py`` sources under ``src/excel_mcp`` with the ``ast`` module; does not
import the packages under test. Lazy imports inside functions count the same as
top-level imports.
"""

from __future__ import annotations

import ast
import unittest
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]
_PKG_ROOT = _REPO_ROOT / "src" / "excel_mcp"


def _py_files_under(root: Path) -> list[Path]:
    return sorted(p for p in root.rglob("*.py") if p.is_file())


def _containing_package_parts(path: Path) -> tuple[str, ...]:
    """Package parts for relative-import resolution (parent of the module file)."""
    rel = path.relative_to(_PKG_ROOT)
    if rel.name == "__init__.py":
        return tuple(rel.parent.parts)
    return tuple(rel.parent.parts)


def _all_import_roots(path: Path, tree: ast.AST) -> set[str]:
    """Every imported module name (absolute, plus relative resolved under excel_mcp)."""
    pkg_parts = _containing_package_parts(path)
    roots: set[str] = set()

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                roots.add(alias.name)
        elif isinstance(node, ast.ImportFrom):
            if node.level and node.level > 0:
                up = node.level - 1
                if up > len(pkg_parts):
                    # Escapes the excel_mcp package — treat as unresolved violation marker.
                    roots.add("__relative_escape__")
                    continue
                base_parts = ("excel_mcp",) + pkg_parts[: len(pkg_parts) - up]
                if node.module:
                    roots.add(".".join(base_parts + tuple(node.module.split("."))))
                else:
                    roots.add(".".join(base_parts))
                    for alias in node.names:
                        if alias.name != "*":
                            roots.add(".".join(base_parts + (alias.name,)))
            elif node.module:
                roots.add(node.module)
    return roots


def _first_segment_after_excel_mcp(mod: str) -> str | None:
    if mod == "excel_mcp":
        return None
    if not mod.startswith("excel_mcp."):
        return None
    return mod[len("excel_mcp.") :].split(".", 1)[0]


def _scan_package(subdir: str) -> list[tuple[Path, set[str]]]:
    out: list[tuple[Path, set[str]]] = []
    for path in _py_files_under(_PKG_ROOT / subdir):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        out.append((path, _all_import_roots(path, tree)))
    return out


def _scan_module(filename: str) -> list[tuple[Path, set[str]]]:
    path = _PKG_ROOT / filename
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    return [(path, _all_import_roots(path, tree))]


class TestImportDirection(unittest.TestCase):
    def _assert_no_forbidden(
        self,
        label: str,
        scanned: list[tuple[Path, set[str]]],
        forbidden: frozenset[str],
    ) -> None:
        violations: list[str] = []
        for path, roots in scanned:
            for mod in sorted(roots):
                if mod == "__relative_escape__":
                    violations.append(
                        f"{path.relative_to(_REPO_ROOT)}: relative import escapes excel_mcp"
                    )
                    continue
                top = _first_segment_after_excel_mcp(mod)
                if top in forbidden:
                    violations.append(f"{path.relative_to(_REPO_ROOT)}: imports {mod}")
        if violations:
            self.fail(f"{label} forbidden import(s):\n" + "\n".join(violations))

    def _assert_no_third_party(
        self,
        label: str,
        scanned: list[tuple[Path, set[str]]],
        banned: frozenset[str],
    ) -> None:
        violations: list[str] = []
        for path, roots in scanned:
            for mod in sorted(roots):
                if mod.split(".", 1)[0] in banned:
                    violations.append(f"{path.relative_to(_REPO_ROOT)}: imports {mod}")
        if violations:
            self.fail(
                f"{label} must not import {sorted(banned)}:\n" + "\n".join(violations)
            )

    def test_query_forbidden_imports(self) -> None:
        scanned = _scan_package("query")
        self._assert_no_forbidden(
            "excel_mcp.query",
            scanned,
            frozenset({"fileio", "com", "routing", "server"}),
        )
        self._assert_no_third_party(
            "excel_mcp.query", scanned, frozenset({"openpyxl", "win32com"})
        )

    def test_fileio_forbidden_imports(self) -> None:
        self._assert_no_forbidden(
            "excel_mcp.fileio",
            _scan_package("fileio"),
            frozenset({"routing", "com", "server"}),
        )

    def test_com_forbidden_imports(self) -> None:
        self._assert_no_forbidden(
            "excel_mcp.com",
            _scan_package("com"),
            frozenset({"fileio", "server"}),
        )

    def test_routing_forbidden_imports(self) -> None:
        self._assert_no_forbidden(
            "excel_mcp.routing",
            _scan_package("routing"),
            frozenset({"fileio", "com"}),
        )

    def test_path_forbidden_imports(self) -> None:
        scanned = _scan_package("path")
        self._assert_no_forbidden(
            "excel_mcp.path",
            scanned,
            frozenset({"routing", "fileio", "com", "server"}),
        )
        self._assert_no_third_party(
            "excel_mcp.path", scanned, frozenset({"openpyxl", "win32com"})
        )

    def test_cells_forbidden_imports(self) -> None:
        scanned = _scan_module("cells.py")
        self._assert_no_forbidden(
            "excel_mcp.cells",
            scanned,
            frozenset({"routing", "fileio", "com", "server"}),
        )
        self._assert_no_third_party(
            "excel_mcp.cells", scanned, frozenset({"openpyxl", "win32com"})
        )

    def test_formula_syntax_and_value_mode_forbidden_imports(self) -> None:
        forbidden = frozenset({"fileio", "com", "routing", "server"})
        for name in ("formula_syntax.py", "value_mode.py"):
            scanned = _scan_module(name)
            self._assert_no_forbidden(f"excel_mcp.{Path(name).stem}", scanned, forbidden)


if __name__ == "__main__":
    unittest.main()
