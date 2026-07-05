from __future__ import annotations

import argparse
import os
import posixpath
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath, PureWindowsPath
from typing import Callable, Literal


PRESERVED_NAMES = {
    "category_config.json",
    "app_settings.json",
    "logs",
    "backups",
}

InstallPhase = Literal["waiting", "backing_up", "installing", "rolling_back", "complete"]


@dataclass(frozen=True)
class InstallProgress:
    phase: InstallPhase
    completed_files: int
    total_files: int
    current_file: str = ""


class UpdateInstallError(Exception):
    def __init__(
        self,
        install_error: Exception,
        backup_dir: Path,
        rollback_error: Exception | None = None,
    ) -> None:
        super().__init__(str(install_error))
        self.install_error = install_error
        self.backup_dir = backup_dir
        self.rollback_error = rollback_error


def _safe_members(archive: zipfile.ZipFile) -> list[zipfile.ZipInfo]:
    members = archive.infolist()
    for member in members:
        raw_name = member.filename.replace("\\", "/")
        normalized = posixpath.normpath(raw_name)
        path = PurePosixPath(normalized)
        if (
            raw_name.startswith("/")
            or PureWindowsPath(member.filename).is_absolute()
            or normalized in ("", ".")
            or ".." in path.parts
        ):
            raise ValueError(f"update package contains unsafe path: {member.filename}")
    return members


def _source_root(extract_dir: Path) -> Path:
    children = list(extract_dir.iterdir())
    if len(children) == 1 and children[0].is_dir():
        return children[0]
    return extract_dir


def _install_sources(source_root: Path) -> list[Path]:
    return [
        source
        for source in source_root.rglob("*")
        if source.is_file() and source.relative_to(source_root).parts[0] not in PRESERVED_NAMES
    ]


def apply_update_package(
    package: Path,
    install_dir: Path,
    progress_callback: Callable[[InstallProgress], object] | None = None,
) -> None:
    install_dir.mkdir(parents=True, exist_ok=True)
    backup_dir = Path(tempfile.mkdtemp(prefix="excel-tiqu-backup-"))
    extract_dir = Path(tempfile.mkdtemp(prefix="excel-tiqu-extract-"))
    replaced: list[Path] = []
    cleanup_backup = False

    def report(phase: InstallPhase, completed: int, total: int, current_file: Path | str = "") -> None:
        if progress_callback is not None:
            progress_callback(InstallProgress(phase, completed, total, str(current_file)))

    try:
        with zipfile.ZipFile(package) as archive:
            members = _safe_members(archive)
            archive.extractall(extract_dir, members)

        source_root = _source_root(extract_dir)
        sources = _install_sources(source_root)
        total = len(sources)

        for index, source in enumerate(sources, start=1):
            relative = source.relative_to(source_root)
            target = install_dir / relative
            if target.exists():
                backup = backup_dir / relative
                backup.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(target, backup)
            report("backing_up", index, total, relative)

        for index, source in enumerate(sources, start=1):
            relative = source.relative_to(source_root)
            target = install_dir / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            replaced.append(relative)
            shutil.copy2(source, target)
            report("installing", index, total, relative)

        report("complete", total, total)
        cleanup_backup = True
    except ValueError:
        cleanup_backup = True
        raise
    except Exception as install_error:
        rollback_error: Exception | None = None
        rollback_total = len(replaced)
        try:
            for index, relative in enumerate(reversed(replaced), start=1):
                report("rolling_back", index - 1, rollback_total, relative)
                target = install_dir / relative
                backup = backup_dir / relative
                if backup.exists():
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(backup, target)
                elif target.exists():
                    target.unlink()
                report("rolling_back", index, rollback_total, relative)
            cleanup_backup = True
        except Exception as error:
            rollback_error = error
        raise UpdateInstallError(install_error, backup_dir, rollback_error) from install_error
    finally:
        shutil.rmtree(extract_dir, ignore_errors=True)
        if cleanup_backup:
            shutil.rmtree(backup_dir, ignore_errors=True)


def wait_for_process(pid: int, timeout: float = 30.0) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            os.kill(pid, 0)
        except OSError:
            return
        time.sleep(0.25)
    raise TimeoutError("Timed out waiting for parent process to exit")


class UpdaterWindow:
    POLL_MS = 100

    def __init__(
        self,
        root: object,
        package: Path,
        install_dir: Path,
        parent_pid: int | None = None,
        restart: Path | None = None,
    ) -> None:
        try:
            import customtkinter as ctk
        except ImportError as error:  # pragma: no cover - depends on packaged runtime.
            raise RuntimeError("customtkinter is required for the updater UI") from error

        self.ctk = ctk
        self.root = root
        self.package = package
        self.install_dir = install_dir
        self.parent_pid = parent_pid
        self.restart = restart
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.active = True
        ctk.set_appearance_mode("light")
        self.root.title("Updating Excel Tiqu")
        self.root.geometry("460x260")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()
        self.root.after(self.POLL_MS, self._poll_events)
        threading.Thread(target=self._worker, daemon=True).start()

    def _build_ui(self) -> None:
        ctk = self.ctk
        container = ctk.CTkFrame(self.root, fg_color="#F6F7F9", corner_radius=0)
        container.pack(fill="both", expand=True)
        self.stage_label = ctk.CTkLabel(
            container,
            text="Waiting for Excel Tiqu to close",
            text_color="#17202A",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        self.stage_label.pack(pady=(32, 8))
        self.description_label = ctk.CTkLabel(
            container,
            text="Please keep this window open during the update.",
            text_color="#566573",
            font=ctk.CTkFont(size=13),
        )
        self.description_label.pack(pady=(0, 22))
        self.progress_bar = ctk.CTkProgressBar(container, width=360, height=12)
        self.progress_bar.pack()
        self.progress_bar.set(0)
        self.counter_label = ctk.CTkLabel(
            container,
            text="Preparing...",
            text_color="#566573",
            font=ctk.CTkFont(size=12),
        )
        self.counter_label.pack(pady=(10, 16))
        self.close_button = ctk.CTkButton(container, text="Close", width=110, state="disabled", command=self.root.destroy)
        self.close_button.pack()

    def _worker(self) -> None:
        try:
            self.events.put(("progress", InstallProgress("waiting", 0, 0)))
            if self.parent_pid is not None:
                wait_for_process(self.parent_pid)
            apply_update_package(
                self.package,
                self.install_dir,
                progress_callback=lambda event: self.events.put(("progress", event)),
            )
            self.events.put(("success", None))
        except UpdateInstallError as error:
            self.events.put(("install_error", error))
        except Exception as error:
            self.events.put(("error", error))

    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "progress":
                    self._render_progress(payload)
                elif kind == "success":
                    self._render_success()
                elif kind == "install_error":
                    self._render_install_error(payload)
                elif kind == "error":
                    self._render_error(str(payload))
        except queue.Empty:
            pass
        if self.root.winfo_exists():
            self.root.after(self.POLL_MS, self._poll_events)

    def _render_progress(self, payload: object) -> None:
        if not isinstance(payload, InstallProgress):
            return
        titles = {
            "waiting": "Waiting for Excel Tiqu to close",
            "backing_up": "Backing up current files",
            "installing": "Installing update",
            "rolling_back": "Restoring previous version",
            "complete": "Update installed",
        }
        self.stage_label.configure(text=titles[payload.phase])
        value = payload.completed_files / payload.total_files if payload.total_files > 0 else 0
        self.progress_bar.set(min(1.0, max(0.0, value)))
        if payload.total_files > 0:
            self.counter_label.configure(text=f"{payload.completed_files} / {payload.total_files}  {payload.current_file}")
        else:
            self.counter_label.configure(text="Preparing...")

    def _render_success(self) -> None:
        self.active = False
        self.stage_label.configure(text="Update complete")
        self.description_label.configure(text="Excel Tiqu has been updated.")
        self.progress_bar.set(1)
        self.counter_label.configure(text="100%")
        if self.restart is not None:
            self.root.after(1000, self._restart_and_close)
        else:
            self.close_button.configure(state="normal")

    def _restart_and_close(self) -> None:
        try:
            assert self.restart is not None
            subprocess.Popen([str(self.restart)], cwd=self.install_dir)
        except Exception as error:
            self._render_error(f"Update completed, but restart failed: {error}")
            return
        self.root.destroy()

    def _render_install_error(self, payload: object) -> None:
        if not isinstance(payload, UpdateInstallError):
            self._render_error(str(payload))
            return
        if payload.rollback_error is None:
            message = f"Install failed and the previous version was restored.\n{payload.install_error}"
        else:
            message = (
                "Install failed and rollback also failed.\n"
                f"Install error: {payload.install_error}\n"
                f"Rollback error: {payload.rollback_error}\n"
                f"Backup directory: {payload.backup_dir}"
            )
        self._render_error(message)

    def _render_error(self, message: str) -> None:
        self.active = False
        self.stage_label.configure(text="Update failed")
        self.description_label.configure(text=message, wraplength=390)
        self.counter_label.configure(text="")
        self.close_button.configure(state="normal")

    def _on_close(self) -> None:
        if not self.active:
            self.root.destroy()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Excel Tiqu updater")
    parser.add_argument("--package", required=True)
    parser.add_argument("--install-dir", required=True)
    parser.add_argument("--parent-pid", type=int)
    parser.add_argument("--restart")
    args = parser.parse_args(argv)

    try:
        import customtkinter as ctk
    except ImportError as error:  # pragma: no cover - depends on packaged runtime.
        raise RuntimeError("customtkinter is required for the updater UI") from error

    root = ctk.CTk()
    UpdaterWindow(
        root=root,
        package=Path(args.package),
        install_dir=Path(args.install_dir),
        parent_pid=args.parent_pid,
        restart=Path(args.restart) if args.restart else None,
    )
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
