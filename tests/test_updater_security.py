from __future__ import annotations

import base64
import unittest
from pathlib import Path
from unittest.mock import patch

import updater


class UpdaterSecurityTests(unittest.TestCase):
    def test_repo_config_ignores_owner_and_repo_from_env(self) -> None:
        with patch.object(
            updater,
            "_load_env_file",
            return_value={
                "GITHUB_REPO_OWNER": "evil'; calc; '",
                "GITHUB_REPO_NAME": "payload`$(whoami)",
            },
        ):
            owner, repo, _token, _installer_asset, _hash_asset = updater._repo_config()

        self.assertEqual(owner, updater.DEFAULT_REPO_OWNER)
        self.assertEqual(repo, updater.DEFAULT_REPO_NAME)

    def test_escape_powershell_single_quoted_escapes_single_quotes_and_backticks(self) -> None:
        escaped = updater._escape_powershell_single_quoted("a'b`c")
        self.assertEqual(escaped, "a''b``c")

    def test_latest_release_via_powershell_escapes_owner_and_repo(self) -> None:
        captured = {}

        class _Completed:
            returncode = 0
            stdout = "v1.2.3"
            stderr = ""

        def _fake_run(command, **kwargs):
            captured["command"] = command
            captured["env"] = kwargs.get("env")
            return _Completed()

        with patch.object(updater.subprocess, "run", side_effect=_fake_run):
            version = updater._latest_release_via_powershell("ow'ner", "re`po")

        self.assertEqual(version, "1.2.3")
        script = captured["command"][-1]
        self.assertIn("ow''ner", script)
        self.assertIn("re``po", script)

    def test_latest_release_via_powershell_passes_token_via_env_not_command(self) -> None:
        captured = {}

        class _Completed:
            returncode = 0
            stdout = "v1.2.3"
            stderr = ""

        def _fake_run(command, **kwargs):
            captured["command"] = command
            captured["env"] = kwargs.get("env")
            return _Completed()

        with patch.object(updater.subprocess, "run", side_effect=_fake_run):
            version = updater._latest_release_via_powershell("owner", "repo", token="ghp_secret_token")

        self.assertEqual(version, "1.2.3")
        script = captured["command"][-1]
        self.assertNotIn("ghp_secret_token", script)
        self.assertIn("$env:RECA_GITHUB_TOKEN", script)
        self.assertEqual(captured["env"]["RECA_GITHUB_TOKEN"], "ghp_secret_token")

    def test_run_installer_raises_on_nonzero_exit(self) -> None:
        class _Completed:
            returncode = 5

        with patch.object(updater.subprocess, "run", return_value=_Completed()):
            with self.assertRaisesRegex(RuntimeError, "código 5"):
                updater.run_installer(Path("C:/tmp/setup.exe"), wait=True)

    def test_build_post_exit_installer_script_waits_and_relaunches(self) -> None:
        script = updater._build_post_exit_installer_script(
            Path("C:/tmp/setup.exe"),
            current_pid=1234,
            relaunch_command=["C:/Program Files/RECA/reca.exe", "--flag"],
        )

        self.assertIn("Wait-Process -Id $pidToWait", script)
        self.assertIn("$pidToWait=1234", script)
        self.assertIn("C:\\tmp\\setup.exe", script)
        self.assertIn("C:/Program Files/RECA/reca.exe", script)
        self.assertIn("--flag", script)

    def test_schedule_installer_after_exit_launches_hidden_powershell(self) -> None:
        captured = {}

        class _Proc:
            pid = 4321

        def _fake_popen(command, **kwargs):
            captured["command"] = command
            captured["kwargs"] = kwargs
            return _Proc()

        with patch.object(updater.subprocess, "Popen", side_effect=_fake_popen):
            updater.schedule_installer_after_exit(
                Path("C:/tmp/setup.exe"),
                current_pid=99,
                relaunch_command=["C:/Program Files/RECA/reca.exe"],
            )

        self.assertEqual(captured["command"][:5], [
            "powershell",
            "-NoProfile",
            "-NonInteractive",
            "-WindowStyle",
            "Hidden",
        ])
        self.assertEqual(captured["command"][5], "-EncodedCommand")
        decoded = base64.b64decode(captured["command"][6]).decode("utf-16le")
        self.assertIn("$pidToWait=99", decoded)
        self.assertIn("C:\\tmp\\setup.exe", decoded)
        self.assertIn("C:/Program Files/RECA/reca.exe", decoded)
        self.assertTrue(captured["kwargs"]["close_fds"])
        self.assertIn("creationflags", captured["kwargs"])


if __name__ == "__main__":
    unittest.main()
