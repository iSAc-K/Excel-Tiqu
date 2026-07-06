from __future__ import annotations

from dataclasses import dataclass
import hashlib
import json
from pathlib import Path
import shutil
import ssl
import tempfile
import threading
import time
from typing import Callable
from urllib.request import Request, urlopen


UPDATE_MANIFEST_URL = "https://github.com/iSAc-K/Excel-Tiqu/releases/latest/download/update.json"


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    download_url: str
    sha256: str
    notes: list[str]


@dataclass(frozen=True)
class DownloadProgress:
    phase: str
    downloaded_bytes: int
    total_bytes: int | None
    average_bytes_per_second: float
    estimated_seconds_remaining: float | None


class UpdateCancelled(Exception):
    pass


def _urlopen(url: str, timeout: float):
    request = Request(url, headers={"User-Agent": "Excel-Tiqu-Updater"})
    try:
        import certifi

        context = ssl.create_default_context(cafile=certifi.where())
    except Exception:
        context = ssl.create_default_context()
    return urlopen(request, timeout=timeout, context=context)


def _version_parts(version: str) -> list[int]:
    text = version.strip()
    if text.lower().startswith("v"):
        text = text[1:]
    parts: list[int] = []
    for piece in text.split("."):
        if not piece.isdigit():
            raise ValueError(f"Invalid version: {version}")
        parts.append(int(piece))
    return parts or [0]


def is_newer_version(candidate: str, current: str) -> bool:
    candidate_parts = _version_parts(candidate)
    current_parts = _version_parts(current or "0")
    length = max(len(candidate_parts), len(current_parts))
    candidate_parts.extend([0] * (length - len(candidate_parts)))
    current_parts.extend([0] * (length - len(current_parts)))
    return candidate_parts > current_parts


def parse_update_manifest(payload: dict[str, object]) -> UpdateInfo:
    version = str(payload.get("version") or "").strip()
    download_url = str(payload.get("download_url") or "").strip()
    sha256 = str(payload.get("sha256") or "").strip().lower()
    notes_value = payload.get("notes", [])
    if not version:
        raise ValueError("Manifest is missing version")
    _version_parts(version)
    if not download_url.startswith("https://") or not download_url.lower().endswith(".zip"):
        raise ValueError("Manifest download_url must be an HTTPS ZIP URL")
    if len(sha256) != 64 or any(char not in "0123456789abcdef" for char in sha256):
        raise ValueError("Manifest sha256 must be 64 hex characters")
    if isinstance(notes_value, str):
        notes = [notes_value]
    elif isinstance(notes_value, list):
        notes = [str(item) for item in notes_value]
    else:
        notes = []
    return UpdateInfo(version=version, download_url=download_url, sha256=sha256, notes=notes)


def fetch_update_info(url: str = UPDATE_MANIFEST_URL, timeout: float = 15.0) -> UpdateInfo:
    with _urlopen(url, timeout) as response:
        data = response.read()
    payload = json.loads(data.decode("utf-8-sig"))
    if not isinstance(payload, dict):
        raise ValueError("Manifest root must be an object")
    return parse_update_manifest(payload)


def fetch_update_info_with_retry(
    attempts: int = 2,
    retry_delay: float = 1.0,
    fetcher: Callable[[], UpdateInfo] | None = None,
    sleeper: Callable[[float], object] = time.sleep,
) -> UpdateInfo:
    fetcher = fetcher or fetch_update_info
    last_error: Exception | None = None
    for index in range(max(1, attempts)):
        try:
            return fetcher()
        except Exception as exc:
            last_error = exc
            if index < attempts - 1:
                sleeper(retry_delay)
    assert last_error is not None
    raise last_error


def verify_sha256(path: Path, expected_sha256: str) -> bool:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().lower() == expected_sha256.lower()


def _progress(
    phase: str,
    downloaded: int,
    total: int | None,
    start: float,
    now: float,
) -> DownloadProgress:
    elapsed = max(0.001, now - start)
    speed = downloaded / elapsed
    remaining = None
    if total and speed > 0:
        remaining = max(0.0, (total - downloaded) / speed)
    return DownloadProgress(phase, downloaded, total, speed, remaining)


def _check_cancel(cancel_event: threading.Event | None) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise UpdateCancelled("Update download was cancelled")


def download_update(
    info: UpdateInfo,
    cancel_event: threading.Event | None = None,
    progress_callback: Callable[[DownloadProgress], object] | None = None,
    clock: Callable[[], float] = time.monotonic,
    chunk_size: int = 1024 * 1024,
    timeout: float = 30.0,
) -> Path:
    temp_dir = Path(tempfile.mkdtemp(prefix="excel_tiqu_update_"))
    package_path = temp_dir / "update.zip"
    start = clock()
    downloaded = 0
    total: int | None = None
    try:
        with _urlopen(info.download_url, timeout) as response:
            length = response.headers.get("Content-Length")
            total = int(length) if length and length.isdigit() else None
            with package_path.open("wb") as handle:
                while True:
                    _check_cancel(cancel_event)
                    chunk = response.read(chunk_size)
                    if not chunk:
                        break
                    handle.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback:
                        progress_callback(_progress("downloading", downloaded, total, start, clock()))
        _check_cancel(cancel_event)
        if progress_callback:
            progress_callback(_progress("verifying", downloaded, total, start, clock()))
        if not verify_sha256(package_path, info.sha256):
            raise ValueError("Downloaded update ZIP failed SHA-256 verification")
        if progress_callback:
            progress_callback(_progress("verified", downloaded, total, start, clock()))
        return package_path
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise
