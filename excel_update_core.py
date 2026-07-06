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
        displayed = float(f"{value:.1f}")
        if displayed < 1024 or unit == "TB":
            return f"{displayed:.1f} {unit}"
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
