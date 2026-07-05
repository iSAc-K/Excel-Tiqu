from __future__ import annotations

import hashlib
import io
import json
import tempfile
import threading
import unittest
import urllib.error
from dataclasses import FrozenInstanceError
from pathlib import Path
from unittest.mock import patch

from update_manager import (
    DownloadProgress,
    UPDATE_MANIFEST_URL,
    USER_AGENT,
    UpdateCancelled,
    UpdateInfo,
    download_update,
    fetch_update_info,
    fetch_update_info_with_retry,
    is_newer_version,
    parse_update_manifest,
    verify_sha256,
)
import update_manager


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


def make_http_error(code: int, message: str) -> urllib.error.HTTPError:
    error = urllib.error.HTTPError("https://example.com/update.json", code, message, {}, io.BytesIO())
    error.close()
    return error


class UpdateManagerTests(unittest.TestCase):
    def test_update_endpoint_constants(self) -> None:
        self.assertEqual(
            UPDATE_MANIFEST_URL,
            "https://github.com/iSAc-K/Excel-Tiqu/releases/latest/download/update.json",
        )
        self.assertEqual(USER_AGENT, "Excel-Tiqu-Updater")

    def test_update_info_and_download_progress_are_immutable(self) -> None:
        info = UpdateInfo("2.2", "https://example.com/app.zip", "a" * 64, [])
        progress = DownloadProgress("downloading", 1, 2, 3.0, 4.0, 5.0)

        with self.assertRaises(FrozenInstanceError):
            info.version = "2.3"  # type: ignore[misc]
        with self.assertRaises(FrozenInstanceError):
            progress.phase = "verified"  # type: ignore[misc]

    def test_semantic_version_comparison(self) -> None:
        self.assertTrue(is_newer_version("2.10", "2.9.9"))
        self.assertTrue(is_newer_version("2.1.1", "2.1"))
        self.assertTrue(is_newer_version("v2.2", "2.1.9"))
        self.assertFalse(is_newer_version("2.2", "v2.2.0"))
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
        with self.assertRaises(ValueError):
            parse_update_manifest({"version": "2.2", "download_url": "https://example.com/app.zip", "sha256": "A" * 64})
        with self.assertRaises(ValueError):
            parse_update_manifest(
                {"version": "not-a-version", "download_url": "https://example.com/app.zip", "sha256": "a" * 64}
            )

    def test_urlopen_uses_certifi_context_when_available(self) -> None:
        fake_context = object()
        fake_response = FakeResponse(b"{}")

        with (
            patch("update_manager.certifi.where", return_value="certifi-ca.pem") as certifi_where,
            patch("update_manager.ssl.create_default_context", return_value=fake_context) as create_context,
            patch("update_manager.urllib.request.urlopen", return_value=fake_response) as urlopen,
        ):
            response = update_manager._urlopen("https://example.com/update.json", 7.0)

        request = urlopen.call_args.args[0]
        self.assertIs(response, fake_response)
        certifi_where.assert_called_once_with()
        create_context.assert_called_once_with(cafile="certifi-ca.pem")
        self.assertEqual(request.full_url, "https://example.com/update.json")
        self.assertEqual(request.headers["User-agent"], USER_AGENT)
        self.assertEqual(urlopen.call_args.kwargs["timeout"], 7.0)
        self.assertIs(urlopen.call_args.kwargs["context"], fake_context)

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

    def test_fetch_update_info_retries_transient_error_matrix(self) -> None:
        cases: list[tuple[str, BaseException]] = [
            ("connection", ConnectionError("temporary connection failure")),
            ("url", urllib.error.URLError("temporary url failure")),
            ("http408", make_http_error(408, "Request Timeout")),
            ("http429", make_http_error(429, "Too Many Requests")),
        ]

        for _name, error in cases:
            with self.subTest(error=repr(error)):
                calls: list[int] = []
                sleeps: list[float] = []
                expected = UpdateInfo("2.4", "https://example.com/app.zip", "a" * 64, [])

                def fetcher() -> UpdateInfo:
                    calls.append(1)
                    if len(calls) < 3:
                        raise error
                    return expected

                result = fetch_update_info_with_retry(
                    attempts=3,
                    retry_delay=0.25,
                    fetcher=fetcher,
                    sleeper=sleeps.append,
                )

                self.assertEqual(result, expected)
                self.assertEqual(calls, [1, 1, 1])
                self.assertEqual(sleeps, [0.25, 0.25])

    def test_fetch_update_info_does_not_retry_nontransient_http_error(self) -> None:
        calls: list[int] = []

        def fetcher() -> UpdateInfo:
            calls.append(1)
            raise make_http_error(404, "Not Found")

        with self.assertRaises(urllib.error.HTTPError) as context:
            fetch_update_info_with_retry(
                attempts=3,
                retry_delay=0.25,
                fetcher=fetcher,
                sleeper=lambda _delay: None,
            )

        self.assertEqual(context.exception.code, 404)
        self.assertEqual(calls, [1])

    def test_fetch_update_info_retries_transient_http_error_then_succeeds(self) -> None:
        calls: list[int] = []
        sleeps: list[float] = []
        expected = UpdateInfo("2.4", "https://example.com/app.zip", "a" * 64, [])

        def fetcher() -> UpdateInfo:
            calls.append(1)
            if len(calls) < 3:
                raise make_http_error(503, "Unavailable")
            return expected

        result = fetch_update_info_with_retry(
            attempts=3,
            retry_delay=0.25,
            fetcher=fetcher,
            sleeper=sleeps.append,
        )

        self.assertEqual(result, expected)
        self.assertEqual(calls, [1, 1, 1])
        self.assertEqual(sleeps, [0.25, 0.25])

    def test_fetch_update_info_does_not_retry_manifest_validation_errors(self) -> None:
        calls: list[int] = []

        def fetcher() -> UpdateInfo:
            calls.append(1)
            raise ValueError("bad manifest")

        with self.assertRaisesRegex(ValueError, "bad manifest"):
            fetch_update_info_with_retry(
                attempts=3,
                retry_delay=0.25,
                fetcher=fetcher,
                sleeper=lambda _delay: None,
            )

        self.assertEqual(calls, [1])

    def test_fetch_update_info_does_not_retry_non_transient_errors(self) -> None:
        calls: list[int] = []

        def fetcher() -> UpdateInfo:
            calls.append(1)
            raise TypeError("programming error")

        with self.assertRaisesRegex(TypeError, "programming error"):
            fetch_update_info_with_retry(
                attempts=3,
                retry_delay=0.25,
                fetcher=fetcher,
                sleeper=lambda _delay: None,
            )

        self.assertEqual(calls, [1])

    def test_fetch_update_info_rejects_attempts_less_than_one(self) -> None:
        def fetcher() -> UpdateInfo:
            raise AssertionError("fetcher should not run")

        with self.assertRaisesRegex(ValueError, "attempts must be at least 1"):
            fetch_update_info_with_retry(attempts=0, fetcher=fetcher)

    def test_verify_sha256(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "update.zip"
            path.write_bytes(b"payload")

            self.assertTrue(verify_sha256(path, hashlib.sha256(b"payload").hexdigest()))
            self.assertFalse(verify_sha256(path, "0" * 64))

    def test_verify_sha256_emits_initial_verifying_progress(self) -> None:
        events: list[DownloadProgress] = []
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "update.zip"
            path.write_bytes(b"payload")

            self.assertTrue(
                verify_sha256(
                    path,
                    hashlib.sha256(b"payload").hexdigest(),
                    progress_callback=events.append,
                    clock=ControlledClock(10, 10),
                )
            )

        self.assertEqual(events[0].phase, "verifying")
        self.assertEqual(events[0].downloaded_bytes, 0)

    def test_download_reports_progress_and_verified(self) -> None:
        payload = b"x" * 12
        info = UpdateInfo("2.2", "https://example.com/app.zip", hashlib.sha256(payload).hexdigest(), [])
        events: list[DownloadProgress] = []
        with tempfile.TemporaryDirectory() as tmp:
            download_dir = Path(tmp) / "download"
            download_dir.mkdir()
            with (
                patch("update_manager._urlopen", return_value=FakeResponse(payload, str(len(payload)))),
                patch("update_manager.tempfile.mkdtemp", return_value=str(download_dir)) as mkdtemp,
            ):
                path = download_update(info, progress_callback=events.append, clock=ControlledClock(10, 11, 12, 14), chunk_size=4)

            self.assertTrue(path.exists())
            self.assertEqual(mkdtemp.call_args.kwargs["prefix"], "excel-tiqu-update-")
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

    def test_download_open_failure_deletes_temp_dir(self) -> None:
        info = UpdateInfo("2.2", "https://example.com/app.zip", "a" * 64, [])
        with tempfile.TemporaryDirectory() as tmp:
            download_dir = Path(tmp) / "download"
            download_dir.mkdir()
            with (
                patch("update_manager._urlopen", side_effect=OSError("network unavailable")),
                patch("update_manager.tempfile.mkdtemp", return_value=str(download_dir)),
            ):
                with self.assertRaisesRegex(OSError, "network unavailable"):
                    download_update(info)

            self.assertFalse(download_dir.exists())

    def test_download_progress_callback_failure_deletes_temp_dir(self) -> None:
        payload = b"payload"
        info = UpdateInfo("2.2", "https://example.com/app.zip", hashlib.sha256(payload).hexdigest(), [])

        def failing_callback(_event: DownloadProgress) -> None:
            raise RuntimeError("progress failed")

        with tempfile.TemporaryDirectory() as tmp:
            download_dir = Path(tmp) / "download"
            download_dir.mkdir()
            with (
                patch("update_manager._urlopen", return_value=FakeResponse(payload, str(len(payload)))),
                patch("update_manager.tempfile.mkdtemp", return_value=str(download_dir)),
            ):
                with self.assertRaisesRegex(RuntimeError, "progress failed"):
                    download_update(info, progress_callback=failing_callback, chunk_size=4)

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
