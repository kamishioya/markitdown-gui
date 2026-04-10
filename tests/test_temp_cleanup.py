from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from markitdown_gui._temp_cleanup import (
    COPILOT_TEMP_PREFIX,
    LEGACY_TEMP_DIR_STALE_AFTER_SECONDS,
    XLSX_PDF_TEMP_PREFIX,
    build_temp_dir_prefix,
    cleanup_markitdown_temp_dirs,
)


class TempCleanupTests(unittest.TestCase):
    def test_cleanup_removes_stale_legacy_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            stale_dir = temp_root / f"{COPILOT_TEMP_PREFIX}legacy"
            stale_dir.mkdir()
            keep_dir = temp_root / "keep-me"
            keep_dir.mkdir()

            old_time = 1000
            stale_after = LEGACY_TEMP_DIR_STALE_AFTER_SECONDS
            import os

            os.utime(stale_dir, (old_time, old_time))

            removed_paths = cleanup_markitdown_temp_dirs(
                temp_root=temp_root,
                legacy_stale_after_seconds=stale_after,
                current_time=old_time + stale_after + 1,
            )

            self.assertEqual(removed_paths, [stale_dir])
            self.assertFalse(stale_dir.exists())
            self.assertTrue(keep_dir.exists())

    def test_cleanup_removes_dead_pid_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            dead_dir = temp_root / f"{XLSX_PDF_TEMP_PREFIX}pid4242-abandoned"
            dead_dir.mkdir()

            with mock.patch(
                "markitdown_gui._temp_cleanup._is_process_running",
                return_value=False,
            ):
                removed_paths = cleanup_markitdown_temp_dirs(temp_root=temp_root)

            self.assertEqual(removed_paths, [dead_dir])
            self.assertFalse(dead_dir.exists())

    def test_cleanup_keeps_live_pid_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            live_dir = temp_root / f"{COPILOT_TEMP_PREFIX}pid5252-active"
            live_dir.mkdir()

            with mock.patch(
                "markitdown_gui._temp_cleanup._is_process_running",
                return_value=True,
            ):
                removed_paths = cleanup_markitdown_temp_dirs(temp_root=temp_root)

            self.assertEqual(removed_paths, [])
            self.assertTrue(live_dir.exists())

    def test_build_temp_dir_prefix_includes_pid_marker(self) -> None:
        prefix = build_temp_dir_prefix(COPILOT_TEMP_PREFIX)
        self.assertTrue(prefix.startswith(COPILOT_TEMP_PREFIX))
        self.assertIn("pid", prefix)


if __name__ == "__main__":
    unittest.main()