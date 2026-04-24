"""Path identity tests for ``resolve_target`` (STORY-2-3, US-4 / PRD AC3).

Windows-only cases use ``skipUnless`` so Linux CI stays green (NFR-6).
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SRC = os.path.join(_REPO_ROOT, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from excel_mcp.path_resolution import resolve_target  # noqa: E402


def _try_symlink(src: str, dst: str, *, target_is_dir: bool = False) -> bool:
    try:
        os.symlink(src, dst, target_is_directory=target_is_dir)
    except OSError:
        return False
    return True


class TestPathEquivalence(unittest.TestCase):
    def test_symlinked_file_resolves_same_as_target(self):
        with tempfile.TemporaryDirectory() as d:
            real = os.path.join(d, "real.xlsx")
            open(real, "wb").close()
            link = os.path.join(d, "link.xlsx")
            if not _try_symlink(real, link):
                self.skipTest("symlinks not supported or not permitted")
            try:
                self.assertEqual(resolve_target(link), resolve_target(real))
            finally:
                try:
                    os.unlink(link)
                except OSError:
                    pass

    @unittest.skipUnless(sys.platform == "win32", "Windows path casing")
    def test_windows_drive_letter_case_equivalence(self):
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            if len(path) < 3 or path[1] != ":":
                self.skipTest("not a traditional drive path")
            flipped = path[0].swapcase() + path[1:]
            self.assertEqual(resolve_target(path), resolve_target(flipped))
        finally:
            try:
                os.unlink(path)
            except OSError:
                pass

    @unittest.skipUnless(sys.platform == "win32", "8.3 short names are Windows-specific")
    def test_windows_short_path_matches_long_path(self):
        import ctypes
        from ctypes import wintypes

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            long_path = os.path.abspath(f.name)
        try:
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            n = ctypes.windll.kernel32.GetShortPathNameW(long_path, buf, len(buf))
            if not n:
                self.skipTest("GetShortPathNameW unavailable for this path")
            short_path = buf.value
            if short_path.lower() == long_path.lower():
                self.skipTest("short name identical to long name on this volume")
            self.assertEqual(resolve_target(short_path), resolve_target(long_path))
        finally:
            try:
                os.unlink(long_path)
            except OSError:
                pass

    @unittest.skipUnless(sys.platform == "win32", "directory junction tests are Windows-specific")
    def test_windows_junction_directory_paths_equivalent(self):
        with tempfile.TemporaryDirectory() as base:
            inner = os.path.join(base, "inner")
            os.makedirs(inner, exist_ok=True)
            real_file = os.path.join(inner, "book.xlsx")
            open(real_file, "wb").close()
            junc = os.path.join(base, "junc")
            if not _try_symlink(inner, junc, target_is_dir=True):
                self.skipTest("directory symlink/junction not permitted")
            try:
                via_junc = os.path.join(junc, "book.xlsx")
                self.assertEqual(resolve_target(via_junc), resolve_target(real_file))
            finally:
                try:
                    os.unlink(via_junc)
                except OSError:
                    pass
                try:
                    os.rmdir(inner)
                except OSError:
                    pass
                try:
                    os.unlink(junc)
                except OSError:
                    pass


if __name__ == "__main__":
    unittest.main()
