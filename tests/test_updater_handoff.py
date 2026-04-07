from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import updater


class UpdaterHandoffTests(unittest.TestCase):
    def test_build_post_exit_installer_cmd_invokes_installer_directly(self) -> None:
        script_path = updater._build_post_exit_installer_cmd(
            Path("C:/tmp/setup.exe"),
            current_pid=1234,
            relaunch_command=["C:/Program Files/RECA/reca.exe", "--flag"],
        )
        self.addCleanup(lambda: script_path.unlink(missing_ok=True))

        content = script_path.read_text(encoding="utf-8")

        self.assertIn('C:\\tmp\\setup.exe /VERYSILENT /CURRENTUSER /SUPPRESSMSGBOXES /NORESTART', content)
        self.assertNotIn('start "" /wait', content.lower())
        self.assertIn('reca_installer_1234.log', content)
        self.assertIn('start "" "C:/Program Files/RECA/reca.exe" --flag', content)

    def test_build_post_exit_installer_cmd_uses_temp_script_path(self) -> None:
        script_path = updater._build_post_exit_installer_cmd(
            Path("C:/tmp/setup.exe"),
            current_pid=77,
            relaunch_command=None,
        )
        self.addCleanup(lambda: script_path.unlink(missing_ok=True))

        self.assertEqual(script_path.parent, Path(tempfile.gettempdir()))
        self.assertEqual(script_path.name, "reca_updater_77.cmd")


if __name__ == "__main__":
    unittest.main()
