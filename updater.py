from __future__ import annotations

from dataclasses import dataclass
import argparse
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile


PROTECTED_NAMES = {"category_config.json", "app_settings.json"}
PROTECTED_DIRS = {"logs", "backups"}


@dataclass(frozen=True)
class InstallProgress:
    phase: str
    message: str


def _safe_members(zip_file: zipfile.ZipFile) -> list[zipfile.ZipInfo]:
    safe: list[zipfile.ZipInfo] = []
    for member in zip_file.infolist():
        normalized = Path(member.filename.replace("\\", "/"))
        if normalized.is_absolute() or ".." in normalized.parts:
            raise ValueError(f"Unsafe ZIP path: {member.filename}")
        safe.append(member)
    return safe


def extract_update_zip(package: Path, destination: Path) -> Path:
    destination.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(package) as archive:
        members = _safe_members(archive)
        archive.extractall(destination, members)
    roots = [item for item in destination.iterdir()]
    dirs = [item for item in roots if item.is_dir()]
    files = [item for item in roots if item.is_file()]
    if len(dirs) == 1 and not files:
        return dirs[0]
    return destination


def _is_protected(relative_path: Path) -> bool:
    if relative_path.parts and relative_path.parts[0] in PROTECTED_DIRS:
        return True
    return relative_path.name in PROTECTED_NAMES


def _copy_tree_filtered(source_root: Path, install_dir: Path, backup_dir: Path) -> list[Path]:
    replaced: list[Path] = []
    for source in source_root.rglob("*"):
        if source.is_dir():
            continue
        relative = source.relative_to(source_root)
        if _is_protected(relative):
            continue
        target = install_dir / relative
        backup = backup_dir / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.exists():
            backup.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(target, backup)
            replaced.append(relative)
        shutil.copy2(source, target)
    return replaced


def _rollback(install_dir: Path, backup_dir: Path, replaced: list[Path]) -> None:
    for relative in reversed(replaced):
        backup = backup_dir / relative
        target = install_dir / relative
        if backup.exists():
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(backup, target)


def install_update(
    package: Path,
    install_dir: Path,
    progress_callback=None,
) -> None:
    temp_root = Path(tempfile.mkdtemp(prefix="excel_tiqu_install_"))
    backup_dir = temp_root / "backup"
    extract_dir = temp_root / "extract"
    replaced: list[Path] = []
    try:
        if progress_callback:
            progress_callback(InstallProgress("extracting", "Extracting update package"))
        source_root = extract_update_zip(package, extract_dir)
        if progress_callback:
            progress_callback(InstallProgress("backing_up", "Backing up replaced files"))
        backup_dir.mkdir(parents=True, exist_ok=True)
        if progress_callback:
            progress_callback(InstallProgress("installing", "Installing update"))
        replaced = _copy_tree_filtered(source_root, install_dir, backup_dir)
        if progress_callback:
            progress_callback(InstallProgress("complete", "Update installed"))
    except Exception:
        if progress_callback:
            progress_callback(InstallProgress("rolling_back", "Rolling back update"))
        _rollback(install_dir, backup_dir, replaced)
        raise
    finally:
        shutil.rmtree(temp_root, ignore_errors=True)


def wait_for_parent(parent_pid: int, timeout: float = 30.0) -> None:
    if parent_pid <= 0:
        return
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            os.kill(parent_pid, 0)
        except OSError:
            return
        time.sleep(0.5)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--package", required=True)
    parser.add_argument("--install-dir", required=True)
    parser.add_argument("--parent-pid", type=int, default=0)
    parser.add_argument("--restart", default="")
    args = parser.parse_args()

    wait_for_parent(args.parent_pid)
    install_update(Path(args.package), Path(args.install_dir))
    if args.restart:
        subprocess.Popen([args.restart], cwd=args.install_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
