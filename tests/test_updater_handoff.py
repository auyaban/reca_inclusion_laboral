from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import updater


class UpdaterHandoffTests(unittest.TestCase):
    def test_start_update_handoff_writes_manifest_and_waits_for_ready_ack(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            session_dir = Path(tmpdir) / "session"
            session_dir.mkdir(parents=True, exist_ok=True)
            installer_path = session_dir / "setup.exe"
            installer_path.write_bytes(b"installer")
            helper_source = Path(tmpdir) / updater.DEFAULT_UPDATE_HELPER_NAME
            helper_source.write_bytes(b"helper")
            app_executable = Path(tmpdir) / "RECA_INCLUSION_LABORAL.exe"
            app_executable.write_bytes(b"app")

            class _Proc:
                pid = 4321

                @staticmethod
                def poll():
                    return None

            def _fake_launch(command):
                manifest = updater._read_json(command[2])
                updater.write_update_status(
                    manifest["status_path"],
                    "ready",
                    "Helper listo.",
                    helper_pid=4321,
                )
                Path(manifest["ack_path"]).write_text("4321\n", encoding="utf-8")
                return _Proc()

            with patch.dict(os.environ, {"LOCALAPPDATA": tmpdir}, clear=False):
                with patch.object(updater, "get_installed_updater_helper_path", return_value=helper_source):
                    with patch.object(updater, "_launch_detached_process", side_effect=_fake_launch):
                        info = updater.start_update_handoff(
                            installer_path=installer_path,
                            expected_version="2.0.12",
                            current_pid=99,
                            relaunch_command=[str(app_executable)],
                            app_executable=app_executable,
                            release_url="https://example.com/release",
                            ack_timeout=1.0,
                        )

            manifest = updater._read_json(info["manifest_path"])
            self.assertEqual(manifest["expected_version"], "2.0.12")
            self.assertEqual(manifest["target_pid"], 99)
            self.assertEqual(manifest["release_url"], "https://example.com/release")
            self.assertTrue(Path(info["ack_path"]).exists())
            self.assertTrue(Path(info["helper_path"]).exists())
            self.assertTrue(Path(info["helper_path"]).name.endswith(".exe"))

    def test_start_update_handoff_clears_pointer_when_helper_never_acks(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            session_dir = Path(tmpdir) / "session"
            session_dir.mkdir(parents=True, exist_ok=True)
            installer_path = session_dir / "setup.exe"
            installer_path.write_bytes(b"installer")
            helper_source = Path(tmpdir) / updater.DEFAULT_UPDATE_HELPER_NAME
            helper_source.write_bytes(b"helper")
            app_executable = Path(tmpdir) / "RECA_INCLUSION_LABORAL.exe"
            app_executable.write_bytes(b"app")

            class _Proc:
                pid = 2222

                @staticmethod
                def poll():
                    return None

            with patch.dict(os.environ, {"LOCALAPPDATA": tmpdir}, clear=False):
                with patch.object(updater, "get_installed_updater_helper_path", return_value=helper_source):
                    with patch.object(updater, "_launch_detached_process", return_value=_Proc()):
                        with self.assertRaisesRegex(RuntimeError, "no confirmó arranque"):
                            updater.start_update_handoff(
                                installer_path=installer_path,
                                expected_version="2.0.12",
                                current_pid=99,
                                relaunch_command=[str(app_executable)],
                                app_executable=app_executable,
                                ack_timeout=0.2,
                            )

                self.assertFalse(updater._update_pointer_path().exists())


if __name__ == "__main__":
    unittest.main()
