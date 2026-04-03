from __future__ import annotations

import unittest
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


if __name__ == "__main__":
    unittest.main()
