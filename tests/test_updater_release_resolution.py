from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import updater


class UpdaterReleaseResolutionTests(unittest.TestCase):
    def test_get_latest_release_uses_redirect_when_api_fails(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            logs_dir = Path(tmpdir) / "logs"
            logs_dir.mkdir(parents=True, exist_ok=True)
            with patch.object(updater, "appdata_logs_dir", return_value=logs_dir):
                with patch.object(updater, "_http_get_json", side_effect=RuntimeError("api down")):
                    with patch.object(updater, "_latest_release_via_redirect", return_value="2.1.1"):
                        version, assets = updater.get_latest_release_assets()

        self.assertEqual(version, "2.1.1")
        self.assertEqual(assets, updater._assets_for_version("2.1.1"))

    def test_get_latest_release_uses_cached_release_when_live_checks_fail(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            logs_dir = Path(tmpdir) / "logs"
            logs_dir.mkdir(parents=True, exist_ok=True)
            cache_path = Path(tmpdir) / updater.RELEASE_CACHE_FILE_NAME
            cache_path.write_text(
                json.dumps(
                    {
                        "version": "2.1.1",
                        "assets": {"setup.exe": "https://example.test/setup.exe"},
                        "source": "release/latest api",
                        "checked_at": "2026-04-08T17:00:00Z",
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            with patch.object(updater, "appdata_logs_dir", return_value=logs_dir):
                with patch.object(updater, "_http_get_json", side_effect=RuntimeError("api down")):
                    with patch.object(
                        updater,
                        "_latest_release_via_redirect",
                        side_effect=RuntimeError("redirect down"),
                    ):
                        version, assets = updater.get_latest_release_assets()

        self.assertEqual(version, "2.1.1")
        self.assertEqual(assets, {"setup.exe": "https://example.test/setup.exe"})

    def test_get_latest_release_ignores_live_version_older_than_cache(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            logs_dir = Path(tmpdir) / "logs"
            logs_dir.mkdir(parents=True, exist_ok=True)
            cache_path = Path(tmpdir) / updater.RELEASE_CACHE_FILE_NAME
            cache_path.write_text(
                json.dumps(
                    {
                        "version": "2.1.1",
                        "assets": {"setup.exe": "https://example.test/setup.exe"},
                        "source": "release/latest api",
                        "checked_at": "2026-04-08T17:00:00Z",
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            with patch.object(updater, "appdata_logs_dir", return_value=logs_dir):
                with patch.object(
                    updater,
                    "_http_get_json",
                    return_value={
                        "tag_name": "v2.0.6",
                        "assets": [
                            {
                                "name": "setup.exe",
                                "browser_download_url": "https://example.test/old-setup.exe",
                            }
                        ],
                    },
                ):
                    with patch.object(
                        updater,
                        "_latest_release_via_redirect",
                        side_effect=RuntimeError("redirect down"),
                    ):
                        version, assets = updater.get_latest_release_assets()

        self.assertEqual(version, "2.1.1")
        self.assertEqual(assets, {"setup.exe": "https://example.test/setup.exe"})

    def test_get_latest_release_persists_successful_live_lookup(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            logs_dir = Path(tmpdir) / "logs"
            logs_dir.mkdir(parents=True, exist_ok=True)
            with patch.object(updater, "appdata_logs_dir", return_value=logs_dir):
                with patch.object(
                    updater,
                    "_http_get_json",
                    return_value={
                        "tag_name": "v2.1.2",
                        "assets": [
                            {
                                "name": updater.DEFAULT_INSTALLER_ASSET,
                                "browser_download_url": "https://example.test/setup.exe",
                            }
                        ],
                    },
                ):
                    with patch.object(
                        updater,
                        "_latest_release_via_redirect",
                        side_effect=RuntimeError("redirect down"),
                    ):
                        version, assets = updater.get_latest_release_assets()

            cache_path = Path(tmpdir) / updater.RELEASE_CACHE_FILE_NAME
            cache = json.loads(cache_path.read_text(encoding="utf-8"))

        self.assertEqual(version, "2.1.2")
        self.assertEqual(
            assets,
            {updater.DEFAULT_INSTALLER_ASSET: "https://example.test/setup.exe"},
        )
        self.assertEqual(cache["version"], "2.1.2")
        self.assertEqual(
            cache["assets"],
            {updater.DEFAULT_INSTALLER_ASSET: "https://example.test/setup.exe"},
        )


if __name__ == "__main__":
    unittest.main()
