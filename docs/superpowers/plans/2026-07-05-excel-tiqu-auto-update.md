# Excel Tiqu Auto Update Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a GitHub Releases based auto-update flow for the Excel order extraction Windows app, matching the proven updater architecture from the Windows file organizer project.

**Architecture:** Keep update checking/downloading, install-time replacement, and GUI orchestration in separate files. `update_manager.py` handles network, manifest, download, progress, cancellation, and SHA-256 verification; `updater.py` runs as a separate process to replace files after the main EXE exits; `extract_orders_gui.py` only owns UI state and process launching.

**Tech Stack:** Python 3, unittest, urllib, ssl/certifi fallback, hashlib, zipfile, tempfile, shutil, threading, queue, CustomTkinter, PyInstaller, PowerShell, GitHub Releases.

---

## File Structure

- Create `VERSION.txt`: single source for the app version.
- Create `update_manager.py`: manifest parsing, semantic version comparison, HTTPS fetch, ZIP download, progress events, SHA-256 verification, cancellation.
- Create `updater.py`: safe ZIP extraction, protected file filtering, backup, install, rollback, restart, small installer window.
- Create `excel_update_core.py`: UI-independent helpers for reading `VERSION.txt`, formatting update progress text, status text, and deciding which update window states are cancellable/closable.
- Modify `extract_orders_gui.py`: add update state, update window, async update check, download/cancel/install orchestration, update locking with the existing extraction worker.
- Modify `build_exe.ps1`: build both the main EXE and `updater.exe`, copy release metadata and config files, optionally produce a release ZIP and `update.json`.
- Create `test_update_manager.py`: unit coverage for update network/parsing/download behavior.
- Create `test_updater.py`: unit coverage for install preservation, zip-slip rejection, and rollback.
- Create `test_excel_update_core.py`: unit coverage for version reading and GUI status/progress formatting.
- Modify `README.md`: document software update behavior and release packaging.

## Task 1: Version Metadata And Core Formatting Helpers

**Files:**
- Create: `VERSION.txt`
- Create: `excel_update_core.py`
- Create: `test_excel_update_core.py`

- [ ] **Step 1: Create the failing tests**

Create `test_excel_update_core.py`:

```python
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
```

- [ ] **Step 2: Run the tests and verify they fail because the module is missing**

Run:

```powershell
python -m unittest test_excel_update_core -v
```

Expected: `ModuleNotFoundError: No module named 'excel_update_core'`.

- [ ] **Step 3: Add the version file**

Create `VERSION.txt`:

```text
v2.1
release_date: 2026-07-05
name: Excel Tiqu
```

- [ ] **Step 4: Implement `excel_update_core.py`**

Create `excel_update_core.py`:

```python
from __future__ import annotations

from dataclasses import dataclass
import math
from pathlib import Path
import sys
from typing import Literal


UpdateStatus = Literal[
    "checking",
    "latest",
    "available",
    "downloading",
    "verifying",
    "cancelled",
    "failed",
    "preparing_install",
    "updater_started",
]


@dataclass(frozen=True)
class DownloadDisplayText:
    downloaded: str
    speed: str
    remaining: str
    percent: str
    value: float
    indeterminate: bool


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def read_version(base_dir: str | Path | None = None) -> str:
    root = Path(base_dir) if base_dir is not None else app_base_dir()
    version_path = root / "VERSION.txt"
    if not version_path.exists():
        return ""
    for line in version_path.read_text(encoding="utf-8").splitlines():
        value = line.strip()
        if value:
            return value[1:] if value.lower().startswith("v") else value
    return ""


def format_byte_count(byte_count: int | float) -> str:
    value = max(0.0, float(byte_count))
    if value < 1024:
        return f"{int(value)} B"
    for unit in ("KB", "MB", "GB", "TB"):
        value /= 1024
        displayed_value = float(f"{value:.1f}")
        if displayed_value < 1024 or unit == "TB":
            return f"{displayed_value:.1f} {unit}"
    raise AssertionError("unreachable")


def format_download_speed(bytes_per_second: int | float) -> str:
    return f"{format_byte_count(bytes_per_second)}/s"


def format_remaining_time(seconds: int | float | None) -> str:
    if seconds is None:
        return "Calculating"
    rounded_seconds = max(0, math.ceil(seconds))
    if rounded_seconds < 60:
        return f"About {rounded_seconds} sec"
    return f"About {math.ceil(rounded_seconds / 60)} min"


def build_update_progress_text(progress: object) -> DownloadDisplayText:
    downloaded_bytes = int(getattr(progress, "downloaded_bytes"))
    total_bytes = getattr(progress, "total_bytes")
    average_speed = float(getattr(progress, "average_bytes_per_second"))
    remaining_seconds = getattr(progress, "estimated_seconds_remaining")
    downloaded = format_byte_count(downloaded_bytes)
    speed = format_download_speed(average_speed)
    if total_bytes is None or int(total_bytes) <= 0:
        return DownloadDisplayText(
            downloaded=downloaded,
            speed=speed,
            remaining="Calculating",
            percent="Downloading",
            value=0.0,
            indeterminate=True,
        )
    total = int(total_bytes)
    value = min(1.0, max(0.0, downloaded_bytes / total))
    return DownloadDisplayText(
        downloaded=f"{downloaded} / {format_byte_count(total)}",
        speed=speed,
        remaining=format_remaining_time(remaining_seconds),
        percent=f"{round(value * 100)}%",
        value=value,
        indeterminate=False,
    )


def build_update_status_text(
    status: UpdateStatus,
    current_version: str,
    latest_version: str = "",
    notes: list[str] | None = None,
    error: str = "",
) -> str:
    if status == "checking":
        return f"Checking for updates...\n\nCurrent version: {current_version}"
    if status == "latest":
        version = latest_version or current_version
        return f"Already on the latest version.\n\nCurrent version: {current_version}\nOnline version: {version}"
    if status == "available":
        note_text = "\n".join(f"- {note}" for note in (notes or [])) or "- No release notes provided"
        return (
            f"New version available.\n\nCurrent version: {current_version}\n"
            f"Latest version: {latest_version}\n\nRelease notes:\n{note_text}"
        )
    if status == "downloading":
        return (
            f"Downloading update...\n\nCurrent version: {current_version}\n"
            f"Target version: {latest_version}"
        )
    if status == "verifying":
        return (
            f"Verifying update package...\n\nCurrent version: {current_version}\n"
            f"Target version: {latest_version}"
        )
    if status == "cancelled":
        return (
            f"Update stopped. No program files were changed.\n\nCurrent version: {current_version}\n"
            f"Target version: {latest_version}"
        )
    if status == "preparing_install":
        return (
            f"Preparing to install update...\n\nCurrent version: {current_version}\n"
            f"Target version: {latest_version}"
        )
    if status == "updater_started":
        return (
            f"Updater started. The app will close to finish installation.\n\nCurrent version: {current_version}\n"
            f"Target version: {latest_version}"
        )
    return (
        f"Update check failed.\n\nCurrent version: {current_version}\n"
        f"Reason: {error or 'Unknown error'}"
    )


def can_cancel_update(status: UpdateStatus) -> bool:
    return status in {"downloading", "verifying"}


def can_close_update_window(status: UpdateStatus) -> bool:
    return status in {"checking", "available", "latest", "failed", "cancelled"}
```

- [ ] **Step 5: Run the focused tests**

Run:

```powershell
python -m unittest test_excel_update_core -v
```

Expected: all tests pass.

- [ ] **Step 6: Commit**

Run:

```powershell
git add -- VERSION.txt excel_update_core.py test_excel_update_core.py
git commit -m "feat: add update version helpers"
```

Expected: commit succeeds.

## Task 2: Update Manager Download And Manifest Logic

**Files:**
- Create: `update_manager.py`
- Create: `test_update_manager.py`

- [ ] **Step 1: Create failing tests for manifest, versions, download, SHA, and cancellation**

Create `test_update_manager.py`:

```python
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
                path = download_update(info, progress_callback=events.append, clock=ControlledClock(10, 11, 12, 14), chunk_size=4)

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
```

- [ ] **Step 2: Run the tests and verify they fail because `update_manager.py` is missing**

Run:

```powershell
python -m unittest test_update_manager -v
```

Expected: `ModuleNotFoundError: No module named 'update_manager'`.

- [ ] **Step 3: Implement `update_manager.py`**

Create `update_manager.py` by adapting the reference project's `update_manager.py`. Use this app-specific URL and user agent:

```python
UPDATE_MANIFEST_URL = "https://github.com/iSAc-K/Excel-Tiqu/releases/latest/download/update.json"
USER_AGENT = "Excel-Tiqu-Updater"
```

The file must define:

```python
@dataclass(frozen=True)
class UpdateInfo:
    version: str
    download_url: str
    sha256: str
    notes: list[str]


ProgressPhase = Literal["downloading", "verifying", "verified"]


@dataclass(frozen=True)
class DownloadProgress:
    phase: ProgressPhase
    downloaded_bytes: int
    total_bytes: int | None
    elapsed_seconds: float
    average_bytes_per_second: float
    estimated_seconds_remaining: float | None
```

Important implementation requirements:

- `_version_tuple("v2.1")` and `_version_tuple("2.1")` both return `(2, 1)`.
- `parse_update_manifest()` accepts `notes` as either `str` or `list`.
- `fetch_update_info()` decodes JSON with `utf-8-sig`.
- `_urlopen()` uses `certifi.where()` when `certifi` is importable.
- `download_update()` creates `tempfile.mkdtemp(prefix="excel-tiqu-update-")`.
- Any download, verification, cancellation, SHA, or callback exception removes the temp dir.
- `verify_sha256()` reports a first `"verifying"` progress event with `downloaded_bytes=0`.

- [ ] **Step 4: Run the focused tests**

Run:

```powershell
python -m unittest test_update_manager -v
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

Run:

```powershell
git add -- update_manager.py test_update_manager.py
git commit -m "feat: add update manager"
```

Expected: commit succeeds.

## Task 3: Independent Installer And Rollback

**Files:**
- Create: `updater.py`
- Create: `test_updater.py`

- [ ] **Step 1: Create failing installer tests**

Create `test_updater.py`:

```python
from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from updater import InstallProgress, UpdateInstallError, apply_update_package


class UpdaterTests(unittest.TestCase):
    def test_replaces_program_files_and_preserves_user_files_and_dirs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            install = root / "install"
            install.mkdir()
            (install / "program.txt").write_text("old", encoding="utf-8")
            (install / "category_config.json").write_text("mine", encoding="utf-8")
            (install / "app_settings.json").write_text("settings", encoding="utf-8")
            (install / "logs").mkdir()
            (install / "logs" / "run.log").write_text("log", encoding="utf-8")
            (install / "backups").mkdir()
            (install / "backups" / "summary.xlsx").write_text("backup", encoding="utf-8")
            package = root / "update.zip"
            with zipfile.ZipFile(package, "w") as archive:
                archive.writestr("program.txt", "new")
                archive.writestr("category_config.json", "overwrite")
                archive.writestr("app_settings.json", "overwrite")
                archive.writestr("logs/run.log", "overwrite")
                archive.writestr("backups/summary.xlsx", "overwrite")

            apply_update_package(package, install)

            self.assertEqual((install / "program.txt").read_text(encoding="utf-8"), "new")
            self.assertEqual((install / "category_config.json").read_text(encoding="utf-8"), "mine")
            self.assertEqual((install / "app_settings.json").read_text(encoding="utf-8"), "settings")
            self.assertEqual((install / "logs" / "run.log").read_text(encoding="utf-8"), "log")
            self.assertEqual((install / "backups" / "summary.xlsx").read_text(encoding="utf-8"), "backup")

    def test_rejects_zip_slip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            package = root / "update.zip"
            with zipfile.ZipFile(package, "w") as archive:
                archive.writestr("../outside.txt", "bad")

            with self.assertRaises(ValueError):
                apply_update_package(package, root / "install")

            self.assertFalse((root / "outside.txt").exists())

    def test_reports_backup_install_and_complete_progress(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            install = root / "install"
            install.mkdir()
            (install / "program.txt").write_text("old", encoding="utf-8")
            package = root / "update.zip"
            with zipfile.ZipFile(package, "w") as archive:
                archive.writestr("program.txt", "new")
                archive.writestr("extra.txt", "extra")
            events: list[InstallProgress] = []

            apply_update_package(package, install, progress_callback=events.append)

            phases = [event.phase for event in events]
            self.assertIn("backing_up", phases)
            self.assertIn("installing", phases)
            self.assertEqual(phases[-1], "complete")

    def test_restores_current_file_when_copy_fails_after_partial_overwrite(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            install = root / "install"
            install.mkdir()
            target = install / "program.txt"
            target.write_text("old-complete-content", encoding="utf-8")
            package = root / "update.zip"
            with zipfile.ZipFile(package, "w") as archive:
                archive.writestr("program.txt", "new-content")
            real_copy2 = __import__("shutil").copy2
            failed_once = False

            def failing_copy(source: Path, destination: Path, *args: object, **kwargs: object) -> object:
                nonlocal failed_once
                source_path = Path(source)
                destination_path = Path(destination)
                if not failed_once and source_path.name == "program.txt" and destination_path == target:
                    failed_once = True
                    destination_path.write_text("partial", encoding="utf-8")
                    raise OSError("simulated interrupted copy")
                return real_copy2(source, destination, *args, **kwargs)

            with patch("updater.shutil.copy2", side_effect=failing_copy):
                with self.assertRaises(UpdateInstallError) as caught:
                    apply_update_package(package, install)

            self.assertIn("simulated interrupted copy", str(caught.exception.install_error))
            self.assertIsNone(caught.exception.rollback_error)
            self.assertFalse(caught.exception.backup_dir.exists())
            self.assertEqual(target.read_text(encoding="utf-8"), "old-complete-content")


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the tests and verify they fail because `updater.py` is missing**

Run:

```powershell
python -m unittest test_updater -v
```

Expected: `ModuleNotFoundError: No module named 'updater'`.

- [ ] **Step 3: Implement `updater.py`**

Create `updater.py` by adapting the reference project's installer, with this protected list:

```python
PRESERVED_NAMES = {
    "category_config.json",
    "app_settings.json",
    "logs",
    "backups",
}
```

The file must define:

```python
InstallPhase = Literal["waiting", "backing_up", "installing", "rolling_back", "complete"]


@dataclass(frozen=True)
class InstallProgress:
    phase: InstallPhase
    completed_files: int
    total_files: int
    current_file: str = ""
```

Implementation requirements:

- `_safe_members()` rejects absolute paths and any member whose normalized `Path.parts` contains `".."`.
- `apply_update_package()` handles ZIPs that either contain files at root or contain one top-level release directory.
- Build the source file list with:

```python
sources = [
    source
    for source in source_root.rglob("*")
    if source.is_file()
    and source.relative_to(source_root).parts[0] not in PRESERVED_NAMES
]
```

- Back up existing target files before copying new files.
- On install failure, restore backed-up files and remove newly created files that had no backup.
- If rollback succeeds, delete the backup dir; if rollback fails, keep the backup dir and expose it in `UpdateInstallError.backup_dir`.
- `main()` accepts `--package`, `--install-dir`, `--parent-pid`, and `--restart`.
- The updater window title should be `Updating Excel Tiqu`.

- [ ] **Step 4: Run the focused tests**

Run:

```powershell
python -m unittest test_updater -v
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

Run:

```powershell
git add -- updater.py test_updater.py
git commit -m "feat: add standalone updater"
```

Expected: commit succeeds.

## Task 4: GUI Update State And Manual Update Window

**Files:**
- Modify: `extract_orders_gui.py`

- [ ] **Step 1: Add update imports and state fields**

Modify the import section of `extract_orders_gui.py`:

```python
import shutil
import subprocess
```

Add imports after `import customtkinter as ctk`:

```python
from excel_update_core import (
    UpdateStatus,
    build_update_progress_text,
    build_update_status_text,
    can_cancel_update,
    can_close_update_window,
    read_version,
)
from update_manager import DownloadProgress, UpdateCancelled, download_update, fetch_update_info_with_retry, is_newer_version
```

In `ExtractOrdersApp.__init__`, after `self.start_controls: list[Any] = []`, add:

```python
self.current_version = read_version(runtime_base_dir()) or "2.1"
self.update_window: ctk.CTkToplevel | None = None
self.update_status_label: ctk.CTkLabel | None = None
self.update_action_button: ctk.CTkButton | None = None
self.update_progress_bar: ctk.CTkProgressBar | None = None
self.update_percent_label: ctk.CTkLabel | None = None
self.update_downloaded_label: ctk.CTkLabel | None = None
self.update_speed_label: ctk.CTkLabel | None = None
self.update_remaining_label: ctk.CTkLabel | None = None
self.update_status: UpdateStatus = "latest"
self.update_latest_version = ""
self.update_cancel_event: threading.Event | None = None
self.update_queue: queue.Queue[tuple[str, object]] = queue.Queue()
self.update_worker_thread: threading.Thread | None = None
self.update_info: object | None = None
```

- [ ] **Step 2: Add a Software Update button**

In `_build_reports_page()`, add an action button beside the existing report/config actions:

```python
self.update_button = ctk.CTkButton(
    actions,
    text="Software Update",
    width=128,
    fg_color="#2563eb",
    command=self.open_update_window,
)
self.update_button.grid(row=2, column=4, sticky="e", padx=(8, 18), pady=(0, 16))
```

If column `4` conflicts with existing widgets, place it in the next unused column in the same `actions` frame. Do not move extraction controls.

- [ ] **Step 3: Add update window methods**

Add these methods inside `ExtractOrdersApp` before `open_path()`:

```python
def is_extract_running(self) -> bool:
    return bool(self.worker_thread and self.worker_thread.is_alive())


def open_update_window(self) -> None:
    if self.update_window is not None and self.update_window.winfo_exists():
        self.update_window.lift()
        self.update_window.focus_force()
        return
    window = ctk.CTkToplevel(self)
    self.update_window = window
    window.title("Software Update")
    window.geometry("560x440")
    window.resizable(False, False)
    window.transient(self)
    window.protocol("WM_DELETE_WINDOW", self.close_update_window)
    window.grid_columnconfigure(0, weight=1)

    self.update_status_label = ctk.CTkLabel(
        window,
        text=build_update_status_text("latest", self.current_version),
        justify="left",
        anchor="w",
        wraplength=500,
    )
    self.update_status_label.grid(row=0, column=0, sticky="ew", padx=24, pady=(24, 14))

    self.update_progress_bar = ctk.CTkProgressBar(window, width=420)
    self.update_progress_bar.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 8))
    self.update_progress_bar.set(0)

    self.update_percent_label = ctk.CTkLabel(window, text="0%")
    self.update_percent_label.grid(row=2, column=0, sticky="w", padx=24)

    metrics = ctk.CTkFrame(window, fg_color="transparent")
    metrics.grid(row=3, column=0, sticky="ew", padx=24, pady=(8, 16))
    metrics.grid_columnconfigure((0, 1, 2), weight=1)
    self.update_downloaded_label = ctk.CTkLabel(metrics, text="Downloaded: 0 B")
    self.update_downloaded_label.grid(row=0, column=0, sticky="w")
    self.update_speed_label = ctk.CTkLabel(metrics, text="Speed: 0 B/s")
    self.update_speed_label.grid(row=0, column=1, sticky="w")
    self.update_remaining_label = ctk.CTkLabel(metrics, text="Remaining: Calculating")
    self.update_remaining_label.grid(row=0, column=2, sticky="w")

    self.update_action_button = ctk.CTkButton(
        window,
        text="Check Again",
        command=self.start_manual_update_check,
    )
    self.update_action_button.grid(row=4, column=0, sticky="ew", padx=80, pady=(6, 24))
    self.render_update_state(self.update_status, self.update_latest_version)
    self.start_manual_update_check()


def close_update_window(self) -> None:
    if not can_close_update_window(self.update_status):
        if self.update_window is not None and self.update_window.winfo_exists():
            self.update_window.lift()
            self.update_window.focus_force()
        return
    if self.update_window is not None:
        self.update_window.destroy()
    self.update_window = None
```

- [ ] **Step 4: Add render and progress methods**

Add:

```python
def render_update_state(
    self,
    status: UpdateStatus,
    latest_version: str = "",
    notes: list[str] | None = None,
    error: str = "",
) -> None:
    self.update_status = status
    if latest_version:
        self.update_latest_version = latest_version
    if self.update_status_label is not None:
        self.update_status_label.configure(
            text=build_update_status_text(
                status,
                self.current_version,
                self.update_latest_version,
                notes,
                error,
            )
        )
    if self.update_action_button is None:
        return
    if status == "checking":
        self.update_action_button.configure(text="Checking...", state=tk.DISABLED, command=lambda: None)
    elif status == "available":
        self.update_action_button.configure(text="Update Now", state=tk.NORMAL, command=self.start_update_download)
    elif can_cancel_update(status):
        self.update_action_button.configure(text="Stop Update", state=tk.NORMAL, command=self.stop_update_download)
    elif status in {"preparing_install", "updater_started"}:
        self.update_action_button.configure(text="Installing...", state=tk.DISABLED, command=lambda: None)
    else:
        self.update_action_button.configure(text="Check Again", state=tk.NORMAL, command=self.start_manual_update_check)


def render_update_progress(self, progress: DownloadProgress) -> None:
    text = build_update_progress_text(progress)
    if progress.phase == "downloading" and self.update_status != "downloading":
        self.render_update_state("downloading", self.update_latest_version)
    elif progress.phase in {"verifying", "verified"} and self.update_status == "downloading":
        self.render_update_state("verifying", self.update_latest_version)
    if self.update_progress_bar is not None:
        self.update_progress_bar.stop()
        self.update_progress_bar.configure(mode="indeterminate" if text.indeterminate else "determinate")
        if text.indeterminate:
            self.update_progress_bar.start()
        else:
            self.update_progress_bar.set(text.value)
    if self.update_percent_label is not None:
        self.update_percent_label.configure(text=text.percent)
    if self.update_downloaded_label is not None:
        self.update_downloaded_label.configure(text=f"Downloaded: {text.downloaded}")
    if self.update_speed_label is not None:
        self.update_speed_label.configure(text=f"Speed: {text.speed}")
    if self.update_remaining_label is not None:
        self.update_remaining_label.configure(text=f"Remaining: {text.remaining}")
```

- [ ] **Step 5: Run syntax check**

Run:

```powershell
python -m py_compile extract_orders_gui.py
```

Expected: no output and exit code 0.

- [ ] **Step 6: Commit**

Run:

```powershell
git add -- extract_orders_gui.py
git commit -m "feat: add update window shell"
```

Expected: commit succeeds.

## Task 5: GUI Check, Download, Cancel, And Start Installer

**Files:**
- Modify: `extract_orders_gui.py`

- [ ] **Step 1: Add startup async update check**

In `ExtractOrdersApp.__init__`, after `self.switch_page(self.selected_page)`, add:

```python
self.after(1200, self.check_for_updates_async)
```

Add:

```python
def check_for_updates_async(self) -> None:
    threading.Thread(target=self._check_for_updates_worker, daemon=True).start()


def _check_for_updates_worker(self) -> None:
    try:
        info = fetch_update_info_with_retry()
    except Exception as exc:
        self.update_queue.put(("check_failed", str(exc)))
        self.after(100, self.poll_update_queue)
        return
    self.update_queue.put(("check_done", info))
    self.after(100, self.poll_update_queue)
```

- [ ] **Step 2: Add manual check worker**

Add:

```python
def start_manual_update_check(self) -> None:
    if self.update_status in {"downloading", "verifying", "preparing_install", "updater_started"}:
        return
    self.render_update_state("checking")
    self.check_for_updates_async()
```

- [ ] **Step 3: Add queue polling**

Add:

```python
def poll_update_queue(self) -> None:
    while True:
        try:
            kind, payload = self.update_queue.get_nowait()
        except queue.Empty:
            break
        if kind == "check_done":
            info = payload
            latest_version = str(getattr(info, "version"))
            if is_newer_version(latest_version, self.current_version):
                self.update_info = info
                self.render_update_state("available", latest_version, list(getattr(info, "notes", [])))
                if self.update_window is None or not self.update_window.winfo_exists():
                    self.open_update_window()
            elif self.update_window is not None and self.update_window.winfo_exists():
                self.render_update_state("latest", latest_version)
        elif kind == "check_failed":
            if self.update_window is not None and self.update_window.winfo_exists():
                self.render_update_state("failed", error=str(payload))
        elif kind == "progress" and isinstance(payload, DownloadProgress):
            self.render_update_progress(payload)
        elif kind == "downloaded":
            package, latest_version = payload  # type: ignore[misc]
            self.prepare_install_update(Path(package), str(latest_version))
        elif kind == "cancelled":
            self.render_update_state("cancelled", self.update_latest_version)
            self.update_cancel_event = None
        elif kind == "failed":
            self.render_update_state("failed", self.update_latest_version, error=str(payload))
            self.update_cancel_event = None
    if self.update_worker_thread and self.update_worker_thread.is_alive():
        self.after(100, self.poll_update_queue)
```

- [ ] **Step 4: Add download and cancellation**

Add:

```python
def start_update_download(self) -> None:
    if self.is_extract_running():
        messagebox.showwarning("Software Update", "Order extraction is running. Please wait for it to finish before updating.", parent=self)
        return
    info = self.update_info
    if info is None:
        self.start_manual_update_check()
        return
    self.update_cancel_event = threading.Event()
    self.render_update_state("downloading", str(getattr(info, "version")))
    self.set_running_state(True)
    self.update_worker_thread = threading.Thread(target=self._download_update_worker, args=(info,), daemon=True)
    self.update_worker_thread.start()
    self.after(100, self.poll_update_queue)


def stop_update_download(self) -> None:
    if self.update_cancel_event is not None:
        self.update_cancel_event.set()
    if self.update_action_button is not None:
        self.update_action_button.configure(text="Stopping...", state=tk.DISABLED)


def _download_update_worker(self, info: object) -> None:
    try:
        package = download_update(
            info,  # type: ignore[arg-type]
            cancel_event=self.update_cancel_event,
            progress_callback=lambda progress: self.update_queue.put(("progress", progress)),
        )
        self.update_queue.put(("downloaded", (package, str(getattr(info, "version")))))
    except UpdateCancelled:
        self.update_queue.put(("cancelled", None))
    except Exception as exc:
        self.update_queue.put(("failed", str(exc)))
```

- [ ] **Step 5: Add installer launch**

Add:

```python
def prepare_install_update(self, package: Path, latest_version: str) -> None:
    self.render_update_state("preparing_install", latest_version)
    try:
        updater_exe = runtime_base_dir() / "updater.exe"
        updater_script = runtime_base_dir() / "updater.py"
        if updater_exe.exists():
            temporary_updater = package.parent / "updater.exe"
            shutil.copy2(updater_exe, temporary_updater)
            command = [str(temporary_updater)]
        elif updater_script.exists():
            command = [sys.executable, str(updater_script)]
        else:
            raise FileNotFoundError("Could not find updater.exe or updater.py.")
        command.extend(
            [
                "--package",
                str(package),
                "--install-dir",
                str(runtime_base_dir()),
                "--parent-pid",
                str(os.getpid()),
                "--restart",
                str(Path(sys.executable).resolve() if getattr(sys, "frozen", False) else Path(__file__).resolve()),
            ]
        )
        subprocess.Popen(command, cwd=runtime_base_dir())
    except Exception as exc:
        self.set_running_state(False)
        self.update_cancel_event = None
        self.render_update_state("failed", latest_version, error=str(exc))
        return
    self.render_update_state("updater_started", latest_version)
    self.after(300, self.destroy)
```

- [ ] **Step 6: Prevent extraction while update is active**

At the beginning of `start_extract()`, before input validation, add:

```python
if self.update_status in {"downloading", "verifying", "preparing_install", "updater_started"}:
    messagebox.showwarning("Software Update", "The app is updating. Please wait for the update to finish.", parent=self)
    return
```

- [ ] **Step 7: Re-enable controls after cancel or failure**

In `poll_update_queue()`, inside the `cancelled` and `failed` branches, add:

```python
self.set_running_state(False)
```

Do not re-enable controls in the `downloaded` branch because installation is starting and the app will exit.

- [ ] **Step 8: Run syntax check**

Run:

```powershell
python -m py_compile extract_orders_gui.py
```

Expected: no output and exit code 0.

- [ ] **Step 9: Run focused unit tests from earlier tasks**

Run:

```powershell
python -m unittest test_excel_update_core test_update_manager test_updater -v
```

Expected: all tests pass.

- [ ] **Step 10: Commit**

Run:

```powershell
git add -- extract_orders_gui.py
git commit -m "feat: wire update flow into GUI"
```

Expected: commit succeeds.

## Task 6: Build Script, Release ZIP, And Manifest Generation

**Files:**
- Modify: `build_exe.ps1`

- [ ] **Step 1: Update `build_exe.ps1` to read `VERSION.txt`**

Replace the hard-coded `$appName` assignment with:

```powershell
$versionLine = Get-Content -LiteralPath (Join-Path $PSScriptRoot "VERSION.txt") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
if ([string]::IsNullOrWhiteSpace($versionLine)) {
    throw "VERSION.txt is missing a version line"
}
$version = $versionLine.Trim()
if ($version.StartsWith("v", [System.StringComparison]::OrdinalIgnoreCase)) {
    $version = $version.Substring(1)
}
$appName = "Excel" + [char]0x8BA2 + [char]0x5355 + [char]0x6570 + [char]0x636E + [char]0x63D0 + [char]0x53D6 + [char]0x5DE5 + [char]0x5177 + "_v" + $version
```

- [ ] **Step 2: Build `updater.exe` after the main EXE**

After the existing main PyInstaller command, add:

```powershell
$updaterName = "updater"
python -m PyInstaller --noconfirm --clean --windowed --onedir --name $updaterName --collect-all customtkinter updater.py
$updaterDist = Join-Path ".\dist" $updaterName
$updaterExe = Join-Path $updaterDist "updater.exe"
if (-not (Test-Path $updaterExe)) {
    throw "updater.exe was not built: $updaterExe"
}
Copy-Item $updaterExe (Join-Path $distDir "updater.exe") -Force
```

- [ ] **Step 3: Copy `VERSION.txt`**

Add:

```powershell
if (Test-Path ".\VERSION.txt") {
    Copy-Item ".\VERSION.txt" $distDir -Force
}
```

- [ ] **Step 4: Generate release ZIP and `update.json` when requested**

Near the end of `build_exe.ps1`, before opening Explorer, add:

```powershell
if ($env:BUILD_RELEASE_ZIP -eq "1") {
    $zipName = "Excel-Tiqu-v$version.zip"
    $zipPath = Join-Path ".\dist" $zipName
    if (Test-Path $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }
    Compress-Archive -LiteralPath $distDir -DestinationPath $zipPath -Force
    $sha = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $manifest = [ordered]@{
        version = $version
        download_url = "https://github.com/iSAc-K/Excel-Tiqu/releases/download/v$version/$zipName"
        sha256 = $sha
        notes = @("新增自动更新。")
    }
    $manifestPath = Join-Path ".\dist" "update.json"
    $manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    Write-Host ("Release ZIP: " + $zipPath)
    Write-Host ("Manifest: " + $manifestPath)
}
```

- [ ] **Step 5: Run syntax-level script validation**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$null = [scriptblock]::Create((Get-Content -Raw .\build_exe.ps1)); 'ok'"
```

Expected: prints `ok`.

- [ ] **Step 6: Commit**

Run:

```powershell
git add -- build_exe.ps1
git commit -m "build: package updater and release manifest"
```

Expected: commit succeeds.

## Task 7: Documentation And Full Verification

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Update README software update section**

Add a section near the packaging section:

```markdown
## Software Update

The Windows GUI can check GitHub Releases for updates. The app reads the local version from `VERSION.txt` and checks:

```text
https://github.com/iSAc-K/Excel-Tiqu/releases/latest/download/update.json
```

When a newer version exists, the user can choose `Software Update` in the GUI. The app downloads the release ZIP, verifies SHA-256, starts `updater.exe`, exits, and lets the updater replace program files.

The updater preserves:

- `category_config.json`
- `app_settings.json`
- `logs/`
- `backups/`

Release builds can generate the ZIP and `update.json` by running:

```powershell
$env:BUILD_RELEASE_ZIP='1'
$env:CODEX_NO_OPEN_EXPLORER='1'
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_exe.ps1
```
```

- [ ] **Step 2: Run all unit tests**

Run:

```powershell
python -m unittest discover -p "test*.py" -v
```

Expected: all tests pass.

If system Python is missing dependencies, run with the bundled runtime used previously in this repo:

```powershell
C:\Users\kt\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m unittest discover -p "test*.py" -v
```

- [ ] **Step 3: Run syntax checks**

Run:

```powershell
python -m py_compile extract_orders.py extract_orders_gui.py excel_update_core.py update_manager.py updater.py
```

Expected: no output and exit code 0.

- [ ] **Step 4: Run whitespace check**

Run:

```powershell
git diff --check
```

Expected: no output and exit code 0.

- [ ] **Step 5: Build release artifacts**

Run:

```powershell
$env:BUILD_RELEASE_ZIP='1'
$env:CODEX_NO_OPEN_EXPLORER='1'
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_exe.ps1
```

Expected:

- `dist/Excel订单数据提取工具_v2.1/Excel订单数据提取工具_v2.1.exe` exists.
- `dist/Excel订单数据提取工具_v2.1/updater.exe` exists.
- `dist/Excel订单数据提取工具_v2.1/VERSION.txt` exists.
- `dist/Excel-Tiqu-v2.1.zip` exists.
- `dist/update.json` exists.

- [ ] **Step 6: Commit**

Run:

```powershell
git add -- README.md
git commit -m "docs: document Excel Tiqu updates"
```

Expected: commit succeeds.

## Task 8: GitHub Repository And First Release

**Files:**
- No source files required unless release verification finds a packaging issue.

- [ ] **Step 1: Create the GitHub repository**

Run:

```powershell
gh repo create iSAc-K/Excel-Tiqu --public --source . --remote origin --push
```

Expected:

- Repository `https://github.com/iSAc-K/Excel-Tiqu` exists.
- Local remote `origin` points to that URL.
- Current branch is pushed.

If `gh` reports the repo already exists, run:

```powershell
git remote add origin https://github.com/iSAc-K/Excel-Tiqu.git
git push -u origin master
```

- [ ] **Step 2: Create the release**

Run:

```powershell
gh release create v2.1 .\dist\Excel-Tiqu-v2.1.zip .\dist\update.json --repo iSAc-K/Excel-Tiqu --title "Excel Tiqu v2.1" --notes "Initial auto-update enabled release."
```

Expected: GitHub Release `v2.1` exists with both assets.

- [ ] **Step 3: Verify the public manifest URL**

Run:

```powershell
python -c "import json,urllib.request; data=json.loads(urllib.request.urlopen('https://github.com/iSAc-K/Excel-Tiqu/releases/latest/download/update.json', timeout=20).read().decode('utf-8-sig')); print(data['version']); print(data['download_url']); print(data['sha256'])"
```

Expected:

- First line is `2.1`.
- Second line contains `https://github.com/iSAc-K/Excel-Tiqu/releases/download/v2.1/Excel-Tiqu-v2.1.zip`.
- Third line equals the SHA-256 from local `dist/update.json`.

- [ ] **Step 4: Commit any release workflow corrections**

If no source changes were needed, skip this step. If packaging fixes were needed, run:

```powershell
git add -- build_exe.ps1 README.md
git commit -m "fix: correct Excel Tiqu release packaging"
git push
```

Expected: only release workflow corrections are committed.

## Self-Review Notes

- Spec coverage: Tasks cover version metadata, update manager, independent updater, GUI update flow, protected user files, build packaging, release manifest, docs, GitHub repo creation, public manifest verification, and tests.
- Placeholder scan: No forbidden placeholder wording remains. Code-changing steps include concrete snippets and expected commands.
- Type consistency: `UpdateStatus`, `DownloadProgress`, `UpdateInfo`, `InstallProgress`, and helper function names are introduced before they are consumed by GUI tasks.
