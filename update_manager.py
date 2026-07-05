from __future__ import annotations

import hashlib
import json
import re
import shutil
import ssl
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from threading import Event
from typing import Callable, Literal
from urllib.parse import urlparse
from urllib.request import Request, urlopen


UPDATE_MANIFEST_URL = "https://github.com/iSAc-K/Excel-Tiqu/releases/latest/download/update.json"
USER_AGENT = "Excel-Tiqu-Updater"

ProgressPhase = Literal["downloading", "verifying", "verified"]


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    download_url: str
    sha256: str
    notes: list[str]


@dataclass(frozen=True)
class DownloadProgress:
    phase: ProgressPhase
    downloaded_bytes: int
    total_bytes: int | None
    elapsed_seconds: float
    average_bytes_per_second: float
    estimated_seconds_remaining: float | None


class UpdateCancelled(Exception):
    pass


def _version_tuple(version: str) -> tuple[int, ...]:
    text = version.strip()
    if text[:1].lower() == "v":
        text = text[1:]
    if re.fullmatch(r"\d+(?:\.\d+)*", text) is None:
        raise ValueError(f"invalid version: {version!r}")
    return tuple(int(part) for part in text.split("."))


def is_newer_version(latest_version: str, current_version: str) -> bool:
    latest = _version_tuple(latest_version)
    current = _version_tuple(current_version)
    width = max(len(latest), len(current))
    return latest + (0,) * (width - len(latest)) > current + (0,) * (width - len(current))


def _validate_download_url(download_url: str) -> None:
    parsed = urlparse(download_url)
    if parsed.scheme.lower() != "https":
        raise ValueError("download_url must use HTTPS")
    if not parsed.path.lower().endswith(".zip"):
        raise ValueError("download_url must point to a ZIP file")


def _validate_sha256(sha256: str) -> None:
    if re.fullmatch(r"[0-9a-f]{64}", sha256) is None:
        raise ValueError("sha256 must be 64 lowercase hexadecimal characters")


def parse_update_manifest(manifest: object) -> UpdateInfo:
    if not isinstance(manifest, dict):
        raise ValueError("manifest must be a JSON object")

    version = manifest.get("version")
    download_url = manifest.get("download_url")
    sha256 = manifest.get("sha256")
    notes_value = manifest.get("notes", [])

    if not isinstance(version, str) or not version.strip():
        raise ValueError("version is required")
    if not isinstance(download_url, str) or not download_url.strip():
        raise ValueError("download_url is required")
    if not isinstance(sha256, str):
        raise ValueError("sha256 is required")

    _version_tuple(version)
    _validate_download_url(download_url)
    _validate_sha256(sha256)

    if isinstance(notes_value, str):
        notes = [notes_value]
    elif isinstance(notes_value, list) and all(isinstance(item, str) for item in notes_value):
        notes = list(notes_value)
    else:
        raise ValueError("notes must be a string or list of strings")

    return UpdateInfo(version.strip(), download_url.strip(), sha256, notes)


def _urlopen(url: str, timeout: float):
    request = Request(url, headers={"User-Agent": USER_AGENT})
    context = None
    try:
        import certifi

        context = ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        context = None
    return urlopen(request, timeout=timeout, context=context)


def fetch_update_info(url: str = UPDATE_MANIFEST_URL, timeout: float = 10.0) -> UpdateInfo:
    with _urlopen(url, timeout) as response:
        data = response.read()
    manifest = json.loads(data.decode("utf-8-sig"))
    return parse_update_manifest(manifest)


def fetch_update_info_with_retry(
    attempts: int = 3,
    retry_delay: float = 1.0,
    fetcher: Callable[[], UpdateInfo] | None = None,
    sleeper: Callable[[float], object] = time.sleep,
) -> UpdateInfo:
    if attempts < 1:
        raise ValueError("attempts must be at least 1")

    fetch = fetcher or fetch_update_info
    last_error: BaseException | None = None
    for attempt in range(attempts):
        try:
            return fetch()
        except ValueError:
            raise
        except Exception as error:
            last_error = error
            if attempt == attempts - 1:
                break
            sleeper(retry_delay)

    assert last_error is not None
    raise last_error


def _build_progress(
    phase: ProgressPhase,
    downloaded_bytes: int,
    total_bytes: int | None,
    start_time: float,
    clock: Callable[[], float],
) -> DownloadProgress:
    elapsed = max(0.0, clock() - start_time)
    average = downloaded_bytes / elapsed if elapsed > 0 and downloaded_bytes > 0 else 0.0
    remaining = None
    if total_bytes is not None and total_bytes > downloaded_bytes and average > 0:
        remaining = (total_bytes - downloaded_bytes) / average
    elif total_bytes is not None and total_bytes <= downloaded_bytes:
        remaining = 0.0

    return DownloadProgress(
        phase=phase,
        downloaded_bytes=downloaded_bytes,
        total_bytes=total_bytes,
        elapsed_seconds=elapsed,
        average_bytes_per_second=average,
        estimated_seconds_remaining=remaining,
    )


def _emit_progress(
    progress_callback: Callable[[DownloadProgress], object] | None,
    progress: DownloadProgress,
) -> None:
    if progress_callback is not None:
        progress_callback(progress)


def verify_sha256(
    path: Path,
    expected_sha256: str,
    progress_callback: Callable[[DownloadProgress], object] | None = None,
    clock: Callable[[], float] = time.monotonic,
) -> bool:
    _validate_sha256(expected_sha256)
    start_time = clock()
    total_bytes = path.stat().st_size
    _emit_progress(progress_callback, _build_progress("verifying", 0, total_bytes, start_time, clock))

    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)

    return digest.hexdigest() == expected_sha256


def _read_content_length(headers: object) -> int | None:
    try:
        value = headers.get("Content-Length")  # type: ignore[attr-defined]
    except AttributeError:
        return None
    if value is None:
        return None
    try:
        length = int(value)
    except ValueError:
        return None
    return length if length >= 0 else None


def download_update(
    info: UpdateInfo,
    cancel_event: Event | None = None,
    progress_callback: Callable[[DownloadProgress], object] | None = None,
    clock: Callable[[], float] = time.monotonic,
    chunk_size: int = 1024 * 1024,
    timeout: float = 30.0,
) -> Path:
    _validate_download_url(info.download_url)
    _validate_sha256(info.sha256)
    temp_dir = Path(tempfile.mkdtemp(prefix="excel-tiqu-update-"))
    output_path = temp_dir / Path(urlparse(info.download_url).path).name
    start_time = clock()
    downloaded = 0
    total_bytes: int | None = None

    try:
        with _urlopen(info.download_url, timeout) as response:
            total_bytes = _read_content_length(response.headers)
            with output_path.open("wb") as handle:
                while True:
                    if cancel_event is not None and cancel_event.is_set():
                        raise UpdateCancelled("Update download cancelled")
                    chunk = response.read(chunk_size)
                    if not chunk:
                        break
                    handle.write(chunk)
                    downloaded += len(chunk)
                    _emit_progress(
                        progress_callback,
                        _build_progress("downloading", downloaded, total_bytes, start_time, clock),
                    )

        if cancel_event is not None and cancel_event.is_set():
            raise UpdateCancelled("Update download cancelled")

        if not verify_sha256(output_path, info.sha256, progress_callback, clock):
            raise ValueError("Downloaded update SHA-256 does not match manifest")

        _emit_progress(
            progress_callback,
            _build_progress("verified", downloaded, total_bytes, start_time, clock),
        )
        return output_path
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise
