from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from formularios import common


class CommonPathTests(unittest.TestCase):
    def test_sanitize_filename_handles_reserved_windows_names(self) -> None:
        self.assertEqual(common._sanitize_filename("CON.", default="Empresa"), "CON_")
        self.assertEqual(common._sanitize_filename("aux.txt", default="Empresa"), "aux_.txt")

    def test_build_process_output_path_falls_back_when_desktop_candidate_is_too_long(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            long_desktop = str(Path(tmpdir) / ("desktop_" + ("x" * 240)))
            local_app_data = str(Path(tmpdir) / "localappdata")
            with patch.object(common, "_get_desktop_dir", return_value=long_desktop):
                with patch.dict(os.environ, {"LOCALAPPDATA": local_app_data}, clear=False):
                    output_path = common._build_process_output_path(
                        "Empresa Demo",
                        "Proceso de Seleccion Incluyente",
                    )

        self.assertTrue(str(output_path).startswith(local_app_data))

    def test_ensure_roaming_service_account_restores_bundle_and_updates_env(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            roaming_dir = Path(tmpdir) / "roaming"
            bundle_dir = Path(tmpdir) / "bundle"
            bundle_dir.mkdir(parents=True, exist_ok=True)
            bundled_creds = bundle_dir / common.DEFAULT_SERVICE_ACCOUNT_FILE_NAME
            bundled_creds.write_text('{"type":"service_account"}', encoding="utf-8")

            with patch.object(common, "_get_roaming_app_dir", return_value=str(roaming_dir)):
                with patch.object(common, "_get_bundle_dir", return_value=str(bundle_dir)):
                    with patch.object(common, "_get_project_root", return_value=str(bundle_dir)):
                        restored = common._ensure_roaming_service_account_file()
            restored_path = Path(restored)
            self.assertEqual(restored_path, roaming_dir / common.DEFAULT_SERVICE_ACCOUNT_FILE_NAME)
            self.assertTrue(restored_path.exists())
            self.assertEqual(restored_path.read_text(encoding="utf-8"), '{"type":"service_account"}')
            env_path = roaming_dir / ".env"
            self.assertTrue(env_path.exists())
            self.assertIn(
                "GOOGLE_SERVICE_ACCOUNT_FILE=service-account.json",
                env_path.read_text(encoding="utf-8"),
            )


if __name__ == "__main__":
    unittest.main()
