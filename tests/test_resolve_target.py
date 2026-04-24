import os
import sys
import tempfile
import unittest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.path_resolution import resolve_target  # noqa: E402


class TestResolveTarget(unittest.TestCase):
    def test_absolute_normalizes_with_realpath(self):
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            self.assertEqual(resolve_target(path), os.path.realpath(path))
        finally:
            os.unlink(path)

    def test_relative_prefers_first_existing_file_under_search_roots(self):
        with tempfile.TemporaryDirectory() as d1, tempfile.TemporaryDirectory() as d2:
            only = os.path.join(d2, "book.xlsx")
            with open(only, "wb"):
                pass
            out = resolve_target(
                "book.xlsx",
                cwd=d1,
                search_roots=(d1, d2),
            )
            self.assertEqual(out, os.path.realpath(only))

    def test_relative_falls_back_to_cwd_when_roots_miss(self):
        with tempfile.TemporaryDirectory() as d1, tempfile.TemporaryDirectory() as d2:
            target = os.path.join(d2, "new.xlsx")
            out = resolve_target("new.xlsx", cwd=d2, search_roots=(d1,))
            self.assertEqual(out, os.path.realpath(target))

    def test_rejects_empty(self):
        with self.assertRaises(ValueError):
            resolve_target("")

    def test_rejects_nul(self):
        with self.assertRaises(ValueError):
            resolve_target("a\x00b.xlsx")


if __name__ == "__main__":
    unittest.main()
