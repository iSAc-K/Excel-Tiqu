from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from excel_update_core import (
    DownloadDisplayText,
    build_update_progress_text,
    build_update_status_text,
    can_cancel_update,
    can_close_update_window,
    format_byte_count,
    format_remaining_time,
    read_version,
)


class FakeProgress:
    def __init__(
        self,
        downloaded_bytes: int,
        total_bytes: int | None,
        average_bytes_per_second: float,
        estimated_seconds_remaining: float | None,
    ) -> None:
        self.downloaded_bytes = downloaded_bytes
        self.total_bytes = total_bytes
        self.average_bytes_per_second = average_bytes_per_second
        self.estimated_seconds_remaining = estimated_seconds_remaining


class ExcelUpdateCoreTests(unittest.TestCase):
    def test_read_version_strips_optional_v_prefix(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "VERSION.txt").write_text("v2.1\nrelease_date: 2026-07-05\n", encoding="utf-8")

            self.assertEqual(read_version(root), "2.1")

    def test_read_version_returns_empty_for_missing_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(read_version(Path(tmp)), "")

    def test_format_byte_count(self) -> None:
        self.assertEqual(format_byte_count(12), "12 B")
        self.assertEqual(format_byte_count(1536), "1.5 KB")
        self.assertEqual(format_byte_count(2 * 1024 * 1024), "2.0 MB")

    def test_format_remaining_time(self) -> None:
        self.assertEqual(format_remaining_time(None), "Calculating")
        self.assertEqual(format_remaining_time(4.2), "About 5 sec")
        self.assertEqual(format_remaining_time(61), "About 2 min")

    def test_build_progress_text_for_known_total(self) -> None:
        text = build_update_progress_text(FakeProgress(5, 10, 2.5, 2.0))

        self.assertEqual(
            text,
            DownloadDisplayText(
                downloaded="5 B / 10 B",
                speed="2 B/s",
                remaining="About 2 sec",
                percent="50%",
                value=0.5,
                indeterminate=False,
            ),
        )

    def test_build_progress_text_for_unknown_total(self) -> None:
        text = build_update_progress_text(FakeProgress(5, None, 2.5, None))

        self.assertEqual(text.downloaded, "5 B")
        self.assertEqual(text.percent, "Downloading")
        self.assertTrue(text.indeterminate)

    def test_status_texts_include_versions_and_errors(self) -> None:
        self.assertIn("Checking", build_update_status_text("checking", "2.1"))
        self.assertIn("latest version", build_update_status_text("latest", "2.1", "2.1"))
        available = build_update_status_text("available", "2.1", "2.2", ["Auto update"])
        self.assertIn("2.2", available)
        self.assertIn("Auto update", available)
        self.assertIn("Network timeout", build_update_status_text("failed", "2.1", error="Network timeout"))

    def test_update_window_state_rules(self) -> None:
        self.assertTrue(can_cancel_update("downloading"))
        self.assertTrue(can_cancel_update("verifying"))
        self.assertFalse(can_cancel_update("preparing_install"))
        self.assertFalse(can_close_update_window("downloading"))
        self.assertTrue(can_close_update_window("latest"))


if __name__ == "__main__":
    unittest.main()
