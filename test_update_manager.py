from __future__ import annotations

import hashlib
import io
import json
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

from update_manager import (
    DownloadProgress,
    UpdateCancelled,
    UpdateInfo,
    download_update,
    fetch_update_info,
    fetch_update_info_with_retry,
    is_newer_version,
    parse_update_manifest,
    verify_sha256,
)


class FakeResponse:
    def __init__(self, payload: bytes, content_length: str | None = None) -> None:
        self.stream = io.BytesIO(payload)
        self.headers: dict[str, str] = {}
        if content_length is not None:
            self.headers["Content-Length"] = content_length

    def read(self, size: int = -1) -> bytes:
        return self.stream.read(size)

    def __enter__(self) -> "FakeResponse":
        return self

    def __exit__(self, *_args: object) -> bool:
        return False


class ControlledClock:
    def __init__(self, *values: float) -> None:
        self.values = iter(values)
        self.last = values[-1]

    def __call__(self) -> float:
        self.last = next(self.values, self.last)
        return self.last


class UpdateManagerTests(unittest.TestCase):
    def test_semantic_version_comparison(self) -> None:
        self.assertTrue(is_newer_version("2.10", "2.9.9"))
        self.assertTrue(is_newer_version("2.1.1", "2.1"))
        self.assertFalse(is_newer_version("2.1.0", "2.1"))
        self.assertFalse(is_newer_version("2.0.9", "2.1"))

    def test_manifest_validation(self) -> None:
        info = parse_update_manifest(
            {
                "version": "2.2",
                "download_url": "https://github.com/iSAc-K/Excel-Tiqu/releases/download/v2.2/Excel-Tiqu-v2.2.zip",
                "sha256": "a" * 64,
                "notes": "Auto update",
            }
        )

        self.assertEqual(info.version, "2.2")
        self.assertEqual(info.notes, ["Auto update"])
        with self.assertRaises(ValueError):
            parse_update_manifest({"version": "2.2", "download_url": "http://example.com/app.zip", "sha256": "a" * 64})
        with self.assertRaises(ValueError):
            parse_update_manifest({"version": "2.2", "download_url": "https://example.com/app.exe", "sha256": "a" * 64})
        with self.assertRaises(ValueError):
            parse_update_manifest({"version": "2.2", "download_url": "https://example.com/app.zip", "sha256": "bad"})

    def test_fetch_update_info_decodes_utf8_sig(self) -> None:
        payload = b"\xef\xbb\xbf" + json.dumps(
            {
                "version": "2.3",
                "download_url": "https://example.com/app.zip",
                "sha256": "a" * 64,
                "notes": [],
            }
        ).encode("utf-8")
        with patch("update_manager._urlopen", return_value=FakeResponse(payload)) as urlopen:
            info = fetch_update_info("https://example.com/update.json", timeout=3.0)

        self.assertEqual(info.version, "2.3")
        self.assertEqual(urlopen.call_args.args[1], 3.0)

    def test_fetch_update_info_retries_transient_errors(self) -> None:
        calls: list[int] = []
        sleeps: list[float] = []
        expected = UpdateInfo("2.4", "https://example.com/app.zip", "a" * 64, [])

        def fetcher() -> UpdateInfo:
            calls.append(1)
            if len(calls) < 3:
                raise TimeoutError("temporary")
            return expected

        result = fetch_update_info_with_retry(
            attempts=3,
            retry_delay=0.25,
            fetcher=fetcher,
            sleeper=sleeps.append,
        )

        self.assertEqual(result, expected)
        self.assertEqual(sleeps, [0.25, 0.25])

    def test_verify_sha256(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "update.zip"
            path.write_bytes(b"payload")

            self.assertTrue(verify_sha256(path, hashlib.sha256(b"payload").hexdigest()))
            self.assertFalse(verify_sha256(path, "0" * 64))

    def test_download_reports_progress_and_verified(self) -> None:
        payload = b"x" * 12
        info = UpdateInfo("2.2", "https://example.com/app.zip", hashlib.sha256(payload).hexdigest(), [])
        events: list[DownloadProgress] = []
        with tempfile.TemporaryDirectory() as tmp:
            download_dir = Path(tmp) / "download"
            download_dir.mkdir()
            with (
                patch("update_manager._urlopen", return_value=FakeResponse(payload, str(len(payload)))),
                patch("update_manager.tempfile.mkdtemp", return_value=str(download_dir)),
            ):
                path = download_update(
                    info,
                    progress_callback=events.append,
                    clock=ControlledClock(10, 11, 12, 14),
                    chunk_size=4,
                )

            self.assertTrue(path.exists())
            self.assertEqual([event.downloaded_bytes for event in events if event.phase == "downloading"], [4, 8, 12])
            self.assertEqual(events[-1].phase, "verified")
            path.unlink()
            download_dir.rmdir()

    def test_download_cancel_deletes_temp_dir(self) -> None:
        payload = b"x" * 20
        info = UpdateInfo("2.2", "https://example.com/app.zip", hashlib.sha256(payload).hexdigest(), [])
        cancel = threading.Event()

        def cancel_after_first_chunk(event: DownloadProgress) -> None:
            if event.phase == "downloading" and event.downloaded_bytes >= 4:
                cancel.set()

        with tempfile.TemporaryDirectory() as tmp:
            download_dir = Path(tmp) / "download"
            download_dir.mkdir()
            with (
                patch("update_manager._urlopen", return_value=FakeResponse(payload, str(len(payload)))),
                patch("update_manager.tempfile.mkdtemp", return_value=str(download_dir)),
            ):
                with self.assertRaises(UpdateCancelled):
                    download_update(info, cancel_event=cancel, progress_callback=cancel_after_first_chunk, chunk_size=4)

            self.assertFalse(download_dir.exists())

    def test_sha256_failure_deletes_temp_dir(self) -> None:
        payload = b"payload"
        info = UpdateInfo("2.2", "https://example.com/app.zip", "0" * 64, [])
        with tempfile.TemporaryDirectory() as tmp:
            download_dir = Path(tmp) / "download"
            download_dir.mkdir()
            with (
                patch("update_manager._urlopen", return_value=FakeResponse(payload, str(len(payload)))),
                patch("update_manager.tempfile.mkdtemp", return_value=str(download_dir)),
            ):
                with self.assertRaises(ValueError):
                    download_update(info)

            self.assertFalse(download_dir.exists())


if __name__ == "__main__":
    unittest.main()
