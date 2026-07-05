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
