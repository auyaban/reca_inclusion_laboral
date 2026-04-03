from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import drive_upload
import google_sheets_client
from formularios.seguimientos import seguimientos


class RuntimeEnvResolutionTests(unittest.TestCase):
    def test_drive_upload_credentials_path_accepts_google_service_account_from_env_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            creds_path = Path(tmpdir) / "service-account.json"
            creds_path.write_text("{}", encoding="utf-8")

            with patch.object(
                drive_upload,
                "_load_env_file",
                return_value=({"GOOGLE_SERVICE_ACCOUNT_FILE": creds_path.name}, os.fspath(Path(tmpdir) / ".env")),
            ):
                with patch.dict(os.environ, {}, clear=True):
                    with patch.object(drive_upload, "_get_bundle_dir", return_value=tmpdir):
                        resolved = drive_upload._get_credentials_path()

        self.assertEqual(resolved, os.fspath(creds_path))

    def test_google_sheets_default_spreadsheet_id_reads_env_file(self) -> None:
        spreadsheet_url = "https://docs.google.com/spreadsheets/d/demoSheetId-123/edit#gid=0"

        with patch.object(
            google_sheets_client,
            "_load_runtime_env",
            return_value={"GOOGLE_SHEETS_DEFAULT_SPREADSHEET_ID": spreadsheet_url},
        ):
            with patch.dict(os.environ, {}, clear=True):
                with patch.object(google_sheets_client, "_load_config", return_value={}):
                    spreadsheet_id = google_sheets_client.get_default_spreadsheet_id()

        self.assertEqual(spreadsheet_id, "demoSheetId-123")

    def test_google_sheets_credentials_path_accepts_relative_path_next_to_env(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            creds_path = Path(tmpdir) / "service-account.json"
            creds_path.write_text("{}", encoding="utf-8")

            with patch.object(
                google_sheets_client,
                "_load_env_file",
                return_value=({"GOOGLE_SERVICE_ACCOUNT_FILE": creds_path.name}, os.fspath(Path(tmpdir) / ".env")),
            ):
                with patch.dict(os.environ, {}, clear=True):
                    with patch.object(google_sheets_client, "_get_bundle_dir", return_value="C:\\bundle"):
                        resolved = google_sheets_client._get_credentials_path()

        self.assertEqual(resolved, os.fspath(creds_path))

    def test_seguimientos_template_id_uses_master_template_id(self) -> None:
        with patch.object(seguimientos, "get_master_template_id", return_value="masterSheetId-789"):
            template_id = seguimientos.get_seguimientos_template_id()

        self.assertEqual(template_id, "masterSheetId-789")

    def test_seguimientos_template_id_falls_back_to_legacy_env_file(self) -> None:
        spreadsheet_url = "https://docs.google.com/spreadsheets/d/templateSheetId-456/edit"

        with patch.object(seguimientos, "get_master_template_id", side_effect=RuntimeError("missing")):
            with patch.object(
                seguimientos,
                "_load_runtime_env",
                return_value={"GOOGLE_SHEETS_SEGUIMIENTOS_TEMPLATE_ID": spreadsheet_url},
            ):
                with patch.dict(os.environ, {}, clear=True):
                    with patch.object(seguimientos.drive_upload, "_load_config", return_value={}):
                        template_id = seguimientos.get_seguimientos_template_id()

        self.assertEqual(template_id, "templateSheetId-456")

    def test_seguimientos_shared_root_reads_env_file(self) -> None:
        expected = r"D:\SeguimientosCompartidos"

        with patch.object(
            seguimientos,
            "_load_runtime_env",
            return_value={"SEGUIMIENTOS_SHARED_ROOT": expected},
        ):
            with patch.dict(os.environ, {}, clear=True):
                with patch.object(seguimientos.drive_upload, "_load_config", return_value={}):
                    shared_root = seguimientos._get_shared_root()

        self.assertEqual(shared_root, expected)


if __name__ == "__main__":
    unittest.main()
