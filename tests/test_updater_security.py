from __future__ import annotations

import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import updater


class UpdaterSecurityTests(unittest.TestCase):
    def test_build_update_manifest_contains_expected_contract(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            session_dir = Path(tmpdir) / "session"
            session_dir.mkdir(parents=True, exist_ok=True)
            installer_path = session_dir / "setup.exe"
            installer_path.write_bytes(b"installer")
            app_path = Path(tmpdir) / "RECA_INCLUSION_LABORAL.exe"
            app_path.write_bytes(b"app")
            version_path = Path(tmpdir) / "VERSION"
            version_path.write_text("2.0.12\n", encoding="utf-8")

            manifest = updater.build_update_manifest(
                session_dir=session_dir,
                installer_path=installer_path,
                expected_version="2.0.12",
                current_pid=77,
                relaunch_command=[str(app_path)],
                installed_app_path=app_path,
                installed_version_paths=[version_path],
                release_url="https://example.com/release",
            )
            expected_hash = updater.calculate_file_sha256(installer_path)

            self.assertEqual(manifest["target_pid"], 77)
            self.assertEqual(manifest["expected_version"], "2.0.12")
            self.assertEqual(manifest["relaunch_args"], [str(app_path)])
            self.assertEqual(manifest["installed_version_paths"], [str(version_path)])
            self.assertEqual(manifest["release_url"], "https://example.com/release")
            self.assertEqual(manifest["installer_sha256"], expected_hash)

    def test_execute_update_manifest_does_not_relaunch_on_failed_installer(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            session_dir = Path(tmpdir) / "session"
            session_dir.mkdir(parents=True, exist_ok=True)
            installer_path = session_dir / "setup.exe"
            installer_path.write_bytes(b"installer")
            version_path = session_dir / "VERSION"
            version_path.write_text("2.0.11\n", encoding="utf-8")
            manifest = updater.build_update_manifest(
                session_dir=session_dir,
                installer_path=installer_path,
                expected_version="2.0.12",
                current_pid=77,
                relaunch_command=["C:/Program Files/RECA/reca.exe"],
                installed_app_path="C:/Program Files/RECA/reca.exe",
                installed_version_paths=[version_path],
            )
            manifest_path = session_dir / updater.UPDATE_MANIFEST_FILE_NAME
            updater._atomic_write_json(manifest_path, manifest)
            relaunch = Mock()

            result = updater.execute_update_manifest(
                manifest_path,
                wait_for_target=lambda _pid: None,
                installer_runner=lambda _args: 5,
                relaunch_runner=relaunch,
            )

        self.assertEqual(result["state"], "failed")
        self.assertEqual(result["installer_exit_code"], 5)
        relaunch.assert_not_called()

    def test_execute_update_manifest_does_not_relaunch_on_version_mismatch(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            session_dir = Path(tmpdir) / "session"
            session_dir.mkdir(parents=True, exist_ok=True)
            installer_path = session_dir / "setup.exe"
            installer_path.write_bytes(b"installer")
            version_path = session_dir / "VERSION"
            version_path.write_text("2.0.11\n", encoding="utf-8")
            manifest = updater.build_update_manifest(
                session_dir=session_dir,
                installer_path=installer_path,
                expected_version="2.0.12",
                current_pid=77,
                relaunch_command=["C:/Program Files/RECA/reca.exe"],
                installed_app_path="C:/Program Files/RECA/reca.exe",
                installed_version_paths=[version_path],
            )
            manifest_path = session_dir / updater.UPDATE_MANIFEST_FILE_NAME
            updater._atomic_write_json(manifest_path, manifest)
            relaunch = Mock()

            def _runner(_args):
                Path(manifest["installer_log_path"]).write_text("ok", encoding="utf-8")
                return 0

            result = updater.execute_update_manifest(
                manifest_path,
                wait_for_target=lambda _pid: None,
                installer_runner=_runner,
                relaunch_runner=relaunch,
            )

        self.assertEqual(result["state"], "failed")
        self.assertEqual(result["expected_version"], "2.0.12")
        self.assertEqual(result["detected_installed_version"], "2.0.11")
        relaunch.assert_not_called()

    def test_execute_update_manifest_relaunches_only_after_matching_version(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            session_dir = Path(tmpdir) / "session"
            session_dir.mkdir(parents=True, exist_ok=True)
            installer_path = session_dir / "setup.exe"
            installer_path.write_bytes(b"installer")
            version_path = session_dir / "VERSION"
            version_path.write_text("2.0.11\n", encoding="utf-8")
            manifest = updater.build_update_manifest(
                session_dir=session_dir,
                installer_path=installer_path,
                expected_version="2.0.12",
                current_pid=77,
                relaunch_command=["C:/Program Files/RECA/reca.exe"],
                installed_app_path="C:/Program Files/RECA/reca.exe",
                installed_version_paths=[version_path],
            )
            manifest_path = session_dir / updater.UPDATE_MANIFEST_FILE_NAME
            updater._atomic_write_json(manifest_path, manifest)
            events = []

            def _wait(_pid):
                events.append("waited")

            def _runner(_args):
                events.append("install")
                time.sleep(0.01)
                Path(manifest["installer_log_path"]).write_text("ok", encoding="utf-8")
                version_path.write_text("2.0.12\n", encoding="utf-8")
                return 0

            def _relaunch(args):
                events.append(("relaunch", list(args)))
                return None

            result = updater.execute_update_manifest(
                manifest_path,
                wait_for_target=_wait,
                installer_runner=_runner,
                relaunch_runner=_relaunch,
            )

        self.assertEqual(result["state"], "succeeded")
        self.assertEqual(result["detected_installed_version"], "2.0.12")
        self.assertEqual(events[0], "waited")
        self.assertEqual(events[1], "install")
        self.assertEqual(events[2][0], "relaunch")

    def test_inspect_pending_update_reports_interrupted_when_status_missing(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            session_dir = Path(tmpdir) / "session"
            session_dir.mkdir(parents=True, exist_ok=True)
            installer_path = session_dir / "setup.exe"
            installer_path.write_bytes(b"installer")
            version_path = session_dir / "VERSION"
            version_path.write_text("2.0.11\n", encoding="utf-8")
            manifest = updater.build_update_manifest(
                session_dir=session_dir,
                installer_path=installer_path,
                expected_version="2.0.12",
                current_pid=77,
                relaunch_command=["C:/Program Files/RECA/reca.exe"],
                installed_app_path="C:/Program Files/RECA/reca.exe",
                installed_version_paths=[version_path],
            )

            with patch.dict("os.environ", {"LOCALAPPDATA": tmpdir}, clear=False):
                updater._atomic_write_json(Path(manifest["manifest_path"]), manifest)
                updater._atomic_write_json(
                    updater._update_pointer_path(),
                    {
                        "session_dir": manifest["session_dir"],
                        "manifest_path": manifest["manifest_path"],
                        "status_path": manifest["status_path"],
                        "expected_version": manifest["expected_version"],
                    },
                )
                pending = updater.inspect_pending_update(current_version="2.0.11")

        self.assertIsNotNone(pending)
        self.assertEqual(pending["outcome"], "interrupted")
        self.assertEqual(pending["expected_version"], "2.0.12")


if __name__ == "__main__":
    unittest.main()
