from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from updater import extract_update_zip, install_update


def make_zip(path: Path, entries: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path, "w") as archive:
        for name, payload in entries.items():
            archive.writestr(name, payload)


class UpdaterTests(unittest.TestCase):
    def test_extract_rejects_zip_slip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            package = root / "bad.zip"
            make_zip(package, {"../outside.txt": b"bad"})

            with self.assertRaises(ValueError):
                extract_update_zip(package, root / "extract")

    def test_install_preserves_user_files_and_dirs(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            install_dir = root / "app"
            install_dir.mkdir()
            (install_dir / "app.exe").write_text("old", encoding="utf-8")
            (install_dir / "category_config.json").write_text("user config", encoding="utf-8")
            (install_dir / "logs").mkdir()
            (install_dir / "logs" / "run.log").write_text("log", encoding="utf-8")

            package = root / "update.zip"
            make_zip(
                package,
                {
                    "Excel/app.exe": b"new",
                    "Excel/category_config.json": b"default config",
                    "Excel/logs/run.log": b"new log",
                    "Excel/VERSION.txt": b"v2.2",
                },
            )

            install_update(package, install_dir)

            self.assertEqual((install_dir / "app.exe").read_text(encoding="utf-8"), "new")
            self.assertEqual((install_dir / "VERSION.txt").read_text(encoding="utf-8"), "v2.2")
            self.assertEqual((install_dir / "category_config.json").read_text(encoding="utf-8"), "user config")
            self.assertEqual((install_dir / "logs" / "run.log").read_text(encoding="utf-8"), "log")

    def test_install_rolls_back_replaced_files_on_failure(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            install_dir = root / "app"
            install_dir.mkdir()
            (install_dir / "app.exe").write_text("old", encoding="utf-8")
            package = root / "update.zip"
            make_zip(package, {"Excel/app.exe": b"new", "Excel/after.txt": b"after"})
            original_copy2 = __import__("shutil").copy2

            def flaky_copy(source: Path, target: Path, *args, **kwargs):
                if str(target).endswith("after.txt"):
                    raise OSError("disk full")
                return original_copy2(source, target, *args, **kwargs)

            with patch("updater.shutil.copy2", side_effect=flaky_copy):
                with self.assertRaises(OSError):
                    install_update(package, install_dir)

            self.assertEqual((install_dir / "app.exe").read_text(encoding="utf-8"), "old")


if __name__ == "__main__":
    unittest.main()
