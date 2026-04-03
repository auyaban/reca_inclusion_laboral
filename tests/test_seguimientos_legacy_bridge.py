from __future__ import annotations

import inspect
import unittest
from unittest.mock import patch

from formularios.seguimientos import seguimientos


class _FakeResponse(dict):
    def __init__(self, status):
        super().__init__()
        self.status = status
        self.headers = {}


class _FakeGoogleError(Exception):
    def __init__(self, status, message="boom"):
        super().__init__(message)
        self.resp = _FakeResponse(status)


class _FakeQueuedRequest:
    def __init__(self, queue):
        self._queue = queue

    def execute(self):
        if not self._queue:
            raise AssertionError("No fake responses left")
        next_item = self._queue.pop(0)
        if isinstance(next_item, Exception):
            raise next_item
        return next_item


class _FakeFilesResource:
    def __init__(self):
        self.list_responses = []
        self.copy_responses = []
        self.list_calls = []
        self.copy_calls = []

    def list(self, **kwargs):
        self.list_calls.append(kwargs)
        return _FakeQueuedRequest(self.list_responses)

    def copy(self, **kwargs):
        self.copy_calls.append(kwargs)
        return _FakeQueuedRequest(self.copy_responses)


class _FakeDriveService:
    def __init__(self):
        self.files_resource = _FakeFilesResource()

    def files(self):
        return self.files_resource


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

    def test_create_native_case_record_uses_confirmed_copy_after_ambiguous_error(self) -> None:
        service = _FakeDriveService()
        service.files_resource.copy_responses = [_FakeGoogleError(503)]
        service.files_resource.list_responses = [
            {
                "files": [
                    {
                        "id": "sheet-123",
                        "name": "Persona Demo - 123456",
                        "mimeType": seguimientos.GOOGLE_SHEETS_MIME,
                        "webViewLink": "https://docs.google.com/spreadsheets/d/sheet-123/edit",
                        "appProperties": {
                            "kind": "seguimiento_il",
                            "request_id": "req-native-123",
                            "cedula": "123456",
                            "max_seguimientos": "6",
                        },
                    }
                ]
            }
        ]

        with (
            patch.object(seguimientos, "get_seguimientos_template_id", return_value="template-id"),
            patch.object(seguimientos.uuid, "uuid4", return_value=type("Uuid", (), {"hex": "req-native-123"})()),
            patch.object(seguimientos, "_build_base_payload_from_user_row", return_value={"nombre_vinculado": "Persona Demo"}),
            patch.object(
                seguimientos,
                "get_spreadsheet",
                return_value={"sheets": [{"properties": {"title": seguimientos.SHEET_BASE}}]},
            ),
            patch.object(seguimientos, "clear_protected_ranges"),
            patch.object(seguimientos, "batch_write_sheet_updates") as batch_write_mock,
            patch.object(seguimientos, "_set_sheet_visibility") as visibility_mock,
        ):
            record = seguimientos._create_native_case_record(
                service,
                "folder-123",
                "Persona Demo - 123456",
                "123456",
                {"nombre_usuario": "Persona Demo"},
                6,
            )

        self.assertEqual(record["file_id"], "sheet-123")
        self.assertEqual(len(service.files_resource.copy_calls), 1)
        self.assertEqual(
            service.files_resource.copy_calls[0]["body"]["appProperties"]["request_id"],
            "req-native-123",
        )
        self.assertEqual(
            service.files_resource.copy_calls[0]["body"]["appProperties"]["kind"],
            "seguimiento_il",
        )
        batch_write_mock.assert_called_once()
        visibility_mock.assert_called_once_with("sheet-123", 6)

    def test_source_no_longer_references_drive_upload_private_request_helper(self) -> None:
        source = inspect.getsource(seguimientos)
        self.assertNotIn("drive_upload._execute_google_request", source)


if __name__ == "__main__":
    unittest.main()
