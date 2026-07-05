from __future__ import annotations

import math
import sys
from dataclasses import dataclass
from pathlib import Path
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


def read_version(base_dir: Path | None = None) -> str:
    version_file = (base_dir or app_base_dir()) / "VERSION.txt"
    if not version_file.exists():
        return ""

    for line in version_file.read_text(encoding="utf-8").splitlines():
        value = line.strip()
        if value:
            if value[:1].lower() == "v":
                return value[1:]
            return value

    return ""


def format_byte_count(byte_count: int | float) -> str:
    if byte_count < 1024:
        return f"{round(byte_count):.0f} B"

    kb_count = byte_count / 1024
    if kb_count < 1024:
        return f"{kb_count:.1f} KB"

    mb_count = kb_count / 1024
    return f"{mb_count:.1f} MB"


def format_download_speed(bytes_per_second: int | float) -> str:
    return f"{format_byte_count(bytes_per_second)}/s"


def format_remaining_time(seconds: float | None) -> str:
    if seconds is None:
        return "Calculating"

    rounded_seconds = max(0, math.ceil(seconds))
    if rounded_seconds < 60:
        return f"About {rounded_seconds} sec"

    rounded_minutes = math.ceil(rounded_seconds / 60)
    return f"About {rounded_minutes} min"


def build_update_progress_text(progress: object) -> DownloadDisplayText:
    downloaded_bytes = getattr(progress, "downloaded_bytes")
    total_bytes = getattr(progress, "total_bytes")
    average_bytes_per_second = getattr(progress, "average_bytes_per_second")
    estimated_seconds_remaining = getattr(progress, "estimated_seconds_remaining")

    speed = format_download_speed(average_bytes_per_second)
    remaining = format_remaining_time(estimated_seconds_remaining)

    if not total_bytes:
        return DownloadDisplayText(
            downloaded=format_byte_count(downloaded_bytes),
            speed=speed,
            remaining=remaining,
            percent="Downloading",
            value=0.0,
            indeterminate=True,
        )

    value = max(0.0, min(1.0, downloaded_bytes / total_bytes))
    return DownloadDisplayText(
        downloaded=f"{format_byte_count(downloaded_bytes)} / {format_byte_count(total_bytes)}",
        speed=speed,
        remaining=remaining,
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
        return f"Checking for updates. Current version: {current_version}."
    if status == "latest":
        return f"You are using the latest version ({current_version})."
    if status == "available":
        note_text = "\n".join(notes or [])
        base = f"Version {latest_version} is available. Current version: {current_version}."
        return f"{base}\n{note_text}" if note_text else base
    if status == "downloading":
        return f"Downloading version {latest_version or current_version}."
    if status == "verifying":
        return "Verifying downloaded update."
    if status == "cancelled":
        return "Update cancelled."
    if status == "failed":
        return f"Update failed: {error}" if error else "Update failed."
    if status == "preparing_install":
        return "Preparing to install update."
    if status == "updater_started":
        return "Updater started. The app will close to finish installing."

    return "Update status unknown."


def can_cancel_update(status: UpdateStatus) -> bool:
    return status in {"downloading", "verifying"}


def can_close_update_window(status: UpdateStatus) -> bool:
    return status in {"checking", "available", "latest", "failed", "cancelled"}
