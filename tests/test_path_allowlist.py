import os
import sys
import tempfile
import unittest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

import excel_mcp.server as server  # noqa: E402
from excel_mcp.path_policy import path_is_allowed  # noqa: E402
from excel_mcp.path_resolution import resolve_target  # noqa: E402


class TestPathAllowlist(unittest.TestCase):
    def setUp(self):
        self._prev_allowed = os.environ.get("EXCEL_MCP_ALLOWED_PATHS")

    def tearDown(self):
        server.EXCEL_FILES_PATH = None
        if self._prev_allowed is None:
            os.environ.pop("EXCEL_MCP_ALLOWED_PATHS", None)
        else:
            os.environ["EXCEL_MCP_ALLOWED_PATHS"] = self._prev_allowed

    def test_stdio_allowlist_passes_inside_root(self):
        server.EXCEL_FILES_PATH = None
        with tempfile.TemporaryDirectory() as allowed:
            os.environ["EXCEL_MCP_ALLOWED_PATHS"] = allowed
            inner = os.path.join(allowed, "book.xlsx")
            open(inner, "wb").close()
            resolved = resolve_target(inner)
            out = server.get_excel_path(inner)
            self.assertEqual(out, resolved)
            self.assertTrue(path_is_allowed(resolved, jail_realpath=None))

    def test_stdio_allowlist_rejects_outside_root(self):
        server.EXCEL_FILES_PATH = None
        with tempfile.TemporaryDirectory() as allowed, tempfile.TemporaryDirectory() as other:
            os.environ["EXCEL_MCP_ALLOWED_PATHS"] = allowed
            outer = os.path.join(other, "nope.xlsx")
            open(outer, "wb").close()
            with self.assertRaises(ValueError):
                server.get_excel_path(outer)

    def test_sse_jail_and_allowlist_intersection_pass(self):
        with tempfile.TemporaryDirectory() as jail:
            os.environ["EXCEL_MCP_ALLOWED_PATHS"] = jail
            server.EXCEL_FILES_PATH = jail
            sub = os.path.join(jail, "sub")
            os.makedirs(sub, exist_ok=True)
            rel = os.path.join("sub", "w.xlsx")
            target = os.path.join(sub, "w.xlsx")
            open(target, "wb").close()
            out = server.get_excel_path(rel)
            self.assertTrue(server._resolved_path_is_within(jail, out))

    def test_sse_inside_jail_outside_allowlist_fails(self):
        with tempfile.TemporaryDirectory() as jail, tempfile.TemporaryDirectory() as other:
            os.environ["EXCEL_MCP_ALLOWED_PATHS"] = other
            server.EXCEL_FILES_PATH = jail
            sub = os.path.join(jail, "sub")
            os.makedirs(sub, exist_ok=True)
            rel = os.path.join("sub", "w.xlsx")
            open(os.path.join(sub, "w.xlsx"), "wb").close()
            with self.assertRaises(ValueError):
                server.get_excel_path(rel)

    def test_sse_allowlist_unset_unchanged_relative_ok(self):
        os.environ.pop("EXCEL_MCP_ALLOWED_PATHS", None)
        with tempfile.TemporaryDirectory() as jail:
            server.EXCEL_FILES_PATH = jail
            sub = os.path.join(jail, "deep")
            os.makedirs(sub, exist_ok=True)
            rel = os.path.join("deep", "f.xlsx")
            open(os.path.join(sub, "f.xlsx"), "wb").close()
            out = server.get_excel_path(rel)
            self.assertTrue(server._resolved_path_is_within(jail, out))

    def test_path_is_allowed_export_sse_intersection(self):
        """``path_is_allowed`` matches jail ∩ allowlist when both apply."""
        with tempfile.TemporaryDirectory() as root:
            jail = os.path.join(root, "jail")
            os.makedirs(os.path.join(jail, "s"), exist_ok=True)
            f = os.path.join(jail, "s", "a.xlsx")
            open(f, "wb").close()
            jail_rp = os.path.realpath(jail)
            root_rp = os.path.realpath(root)
            os.environ["EXCEL_MCP_ALLOWED_PATHS"] = root_rp
            resolved = resolve_target(os.path.join("s", "a.xlsx"), cwd=jail_rp)
            self.assertTrue(path_is_allowed(resolved, jail_realpath=jail_rp))


if __name__ == "__main__":
    unittest.main()
