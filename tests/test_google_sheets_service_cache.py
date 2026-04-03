from __future__ import annotations

import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch

import google_sheets_client


class GoogleSheetsServiceCacheTests(unittest.TestCase):
    def test_service_cache_rebuilds_when_credentials_path_changes(self) -> None:
        fake_google = types.ModuleType("google")
        fake_google_oauth2 = types.ModuleType("google.oauth2")
        fake_service_account = types.ModuleType("google.oauth2.service_account")
        fake_googleapiclient = types.ModuleType("googleapiclient")
        fake_discovery = types.ModuleType("googleapiclient.discovery")
        build_calls = []

        class _FakeCredentials:
            @staticmethod
            def from_service_account_file(path, scopes=None):
                return {"path": path, "scopes": scopes}

        def _fake_build(_service, _version, credentials=None, cache_discovery=None):
            build_calls.append(credentials["path"])
            return {"credentials_path": credentials["path"]}

        fake_service_account.Credentials = _FakeCredentials
        fake_discovery.build = _fake_build

        with tempfile.TemporaryDirectory() as tmpdir:
            creds_a = Path(tmpdir) / "a.json"
            creds_b = Path(tmpdir) / "b.json"
            creds_a.write_text("{}", encoding="utf-8")
            creds_b.write_text("{}", encoding="utf-8")

            with patch.dict(
                sys.modules,
                {
                    "google": fake_google,
                    "google.oauth2": fake_google_oauth2,
                    "google.oauth2.service_account": fake_service_account,
                    "googleapiclient": fake_googleapiclient,
                    "googleapiclient.discovery": fake_discovery,
                },
            ):
                google_sheets_client.clear_google_sheets_service_cache()
                with patch.object(google_sheets_client, "_get_credentials_path", side_effect=[str(creds_a), str(creds_a), str(creds_b)]):
                    first = google_sheets_client.get_google_sheets_service()
                    second = google_sheets_client.get_google_sheets_service()
                    third = google_sheets_client.get_google_sheets_service()

        self.assertEqual(first, second)
        self.assertNotEqual(second, third)
        self.assertEqual(build_calls, [str(creds_a), str(creds_b)])


if __name__ == "__main__":
    unittest.main()
