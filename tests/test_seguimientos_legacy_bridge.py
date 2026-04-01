from __future__ import annotations

import unittest
from unittest.mock import patch

from formularios.seguimientos import seguimientos


class SeguimientosLegacyBridgeTests(unittest.TestCase):
    def test_ensure_case_record_migrates_legacy_local_case_to_google_sheets(self) -> None:
        user_row = {"nombre_usuario": "Persona Demo"}
        legacy_record = {
            "source": "legacy_local",
            "cedula": "123456",
            "local_path": r"C:\tmp\seguimiento.xlsx",
            "max_seguimientos": 6,
        }
        migrated_record = {
            "source": "drive",
            "cedula": "123456",
            "file_id": "sheet-123",
            "webViewLink": "https://docs.google.com/spreadsheets/d/sheet-123/edit",
            "max_seguimientos": 6,
        }
        service = object()

        with patch.object(seguimientos, "find_case_record", return_value=legacy_record):
            with patch.object(seguimientos, "_get_drive_service", return_value=service):
                with patch.object(seguimientos, "_ensure_seguimientos_folder", return_value="seguimientos-root"):
                    with patch.object(
                        seguimientos.drive_upload,
                        "_get_or_create_folder",
                        return_value="case-folder-123",
                    ) as get_folder:
                        with patch.object(
                            seguimientos,
                            "build_case_folder_name",
                            return_value="Persona Demo - 123456",
                        ):
                            with patch.object(
                                seguimientos,
                                "_create_native_case_record",
                                return_value=migrated_record,
                            ) as create_native:
                                result = seguimientos.ensure_case_record("123456", user_row, is_compensar=True)

        get_folder.assert_called_once_with(service, "seguimientos-root", "Persona Demo - 123456")
        create_native.assert_called_once_with(
            service,
            "case-folder-123",
            "Persona Demo - 123456",
            "123456",
            user_row,
            6,
            seed_path=r"C:\tmp\seguimiento.xlsx",
        )
        self.assertFalse(result["created"])
        self.assertEqual(result["max_seguimientos"], 6)
        self.assertIs(result["record"], migrated_record)


if __name__ == "__main__":
    unittest.main()
