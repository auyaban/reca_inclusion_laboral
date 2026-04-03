from __future__ import annotations

import sys
import types
import unittest
from unittest.mock import Mock, patch

import drive_upload
import google_sheets_client
from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.seguimientos import seguimientos


class GoogleSheetsProtectionTests(unittest.TestCase):
    def test_clear_protected_ranges_deletes_all_sheet_protections(self) -> None:
        spreadsheet = {
            "sheets": [
                {
                    "properties": {"sheetId": 1, "title": "Hoja 1"},
                    "protectedRanges": [
                        {"protectedRangeId": 101},
                        {"protectedRangeId": "202"},
                    ],
                },
                {
                    "properties": {"sheetId": 2, "title": "Hoja 2"},
                    "protectedRanges": [
                        {"protectedRangeId": 303},
                        {},
                    ],
                },
            ]
        }
        service = Mock()
        service.spreadsheets.return_value.batchUpdate.return_value.execute.return_value = {}

        with patch.object(google_sheets_client, "get_spreadsheet", return_value=spreadsheet):
            with patch.object(google_sheets_client, "get_google_sheets_service", return_value=service):
                result = google_sheets_client.clear_protected_ranges("spreadsheet-demo")

        self.assertEqual(result["deletedProtectedRangeIds"], [101, 202, 303])
        self.assertEqual(result["deletedProtectedRangeCount"], 3)
        service.spreadsheets.return_value.batchUpdate.assert_called_once_with(
            spreadsheetId="spreadsheet-demo",
            body={
                "requests": [
                    {"deleteProtectedRange": {"protectedRangeId": 101}},
                    {"deleteProtectedRange": {"protectedRangeId": 202}},
                    {"deleteProtectedRange": {"protectedRangeId": 303}},
                ]
            },
        )

    def test_publish_sheet_clears_protections_and_unbolds_copy(self) -> None:
        fake_google = types.ModuleType("google")
        fake_google_oauth2 = types.ModuleType("google.oauth2")
        fake_service_account = types.ModuleType("google.oauth2.service_account")
        fake_googleapiclient = types.ModuleType("googleapiclient")
        fake_discovery = types.ModuleType("googleapiclient.discovery")

        class _FakeCredentials:
            @staticmethod
            def from_service_account_file(*args, **kwargs):
                return object()

        class _FakeFiles:
            def copy(self, **kwargs):
                return Mock(execute=Mock(return_value={"id": "sheet-copy-id", "name": "Demo"}))

        class _FakeDriveService:
            def __init__(self):
                self._files = _FakeFiles()

            def files(self):
                return self._files

        fake_service_account.Credentials = _FakeCredentials
        fake_discovery.build = lambda *args, **kwargs: _FakeDriveService()

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
            with patch.object(drive_upload, "_get_credentials_path", return_value="creds.json"):
                with patch.object(drive_upload, "_get_excel_folder_id", return_value="root-folder"):
                    with patch.object(drive_upload, "_resolve_target_root_id", return_value="resolved-root"):
                        with patch.object(
                            drive_upload,
                            "_get_available_filename",
                            return_value="Evaluacion de Accesibilidad",
                        ):
                            with patch.object(
                                google_sheets_client,
                                "get_evaluacion_accesibilidad_template_id",
                                return_value="template-id",
                            ):
                                with patch.object(
                                    google_sheets_client,
                                    "clear_protected_ranges",
                                ) as clear_protected_ranges:
                                    with patch.object(
                                        google_sheets_client,
                                        "clear_sheet_ranges",
                                    ) as clear_sheet_ranges:
                                        with patch.object(
                                            google_sheets_client,
                                            "batch_write_sheet_updates",
                                        ) as batch_write_sheet_updates:
                                            with patch.object(
                                                google_sheets_client,
                                                "set_sheet_ranges_bold",
                                            ) as set_sheet_ranges_bold:
                                                result = drive_upload.publish_evaluacion_accesibilidad_sheet(
                                                    sheet_writes=[],
                                                    format_ranges=["'Hoja'!A1", "'Hoja'!B2:B3"],
                                                    base_name="Evaluacion de Accesibilidad",
                                                )

        self.assertEqual(result["file_id"], "sheet-copy-id")
        clear_protected_ranges.assert_called_once_with("sheet-copy-id")
        clear_sheet_ranges.assert_not_called()
        batch_write_sheet_updates.assert_called_once()
        self.assertEqual(batch_write_sheet_updates.call_args.args[:2], ("sheet-copy-id", []))
        self.assertIsNone(batch_write_sheet_updates.call_args.kwargs.get("auto_resize_excluded_rows"))
        set_sheet_ranges_bold.assert_called_once_with(
            "sheet-copy-id",
            ["'Hoja'!A1", "'Hoja'!B2:B3"],
            bold=False,
        )

    def test_set_sheet_ranges_bold_builds_repeat_cell_requests(self) -> None:
        sheet_name = "2. EVALUACIÓN DE ACCESIBILIDAD"
        spreadsheet = {
            "sheets": [
                {
                    "properties": {"sheetId": 7, "title": sheet_name},
                }
            ]
        }
        service = Mock()
        service.spreadsheets.return_value.batchUpdate.return_value.execute.return_value = {}

        with patch.object(google_sheets_client, "get_spreadsheet", return_value=spreadsheet):
            with patch.object(google_sheets_client, "get_google_sheets_service", return_value=service):
                result = google_sheets_client.set_sheet_ranges_bold(
                    "spreadsheet-demo",
                    [
                        f"'{sheet_name}'!Q17",
                        f"'{sheet_name}'!C212:C215",
                    ],
                    bold=False,
                )

        self.assertEqual(
            result["updatedRanges"],
            [
                f"'{sheet_name}'!Q17",
                f"'{sheet_name}'!C212:C215",
            ],
        )
        service.spreadsheets.return_value.batchUpdate.assert_called_once_with(
            spreadsheetId="spreadsheet-demo",
            body={
                "requests": [
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": 7,
                                "startRowIndex": 16,
                                "endRowIndex": 17,
                                "startColumnIndex": 16,
                                "endColumnIndex": 17,
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "textFormat": {
                                        "bold": False,
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.textFormat.bold",
                        }
                    },
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": 7,
                                "startRowIndex": 211,
                                "endRowIndex": 215,
                                "startColumnIndex": 2,
                                "endColumnIndex": 3,
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "textFormat": {
                                        "bold": False,
                                    }
                                }
                            },
                            "fields": "userEnteredFormat.textFormat.bold",
                        }
                    },
                ]
            },
        )

    def test_evaluacion_payload_exposes_non_bold_ranges(self) -> None:
        payload = evaluacion_accesibilidad.build_google_sheet_export_payload({})
        sheet_name = evaluacion_accesibilidad.SHEET_NAME
        start_row = evaluacion_accesibilidad.EXCEL_MAPPING["section_8"]["start_row"]
        end_row = start_row + int(evaluacion_accesibilidad.SECTION_8["max_items"]) - 1

        self.assertIn(f"'{sheet_name}'!Q17", payload["format_ranges"])
        self.assertIn(f"'{sheet_name}'!A205", payload["format_ranges"])
        self.assertIn(
            f"'{sheet_name}'!{evaluacion_accesibilidad.EXCEL_MAPPING['section_8']['name_col']}{start_row}:"
            f"{evaluacion_accesibilidad.EXCEL_MAPPING['section_8']['name_col']}{end_row}",
            payload["format_ranges"],
        )
        self.assertEqual(len(payload["format_ranges"]), len(set(payload["format_ranges"])))

    def test_seguimientos_native_copy_clears_protections_before_writing(self) -> None:
        service = Mock()
        service.files.return_value.copy.return_value.execute.return_value = {
            "id": "seguimiento-copy-id",
            "name": "Caso Demo",
            "mimeType": seguimientos.GOOGLE_SHEETS_MIME,
            "webViewLink": "https://docs.google.com/spreadsheets/d/seguimiento-copy-id/edit",
            "appProperties": {},
        }

        with patch.object(seguimientos, "get_seguimientos_template_id", return_value="template-id"):
            with patch.object(seguimientos, "_build_base_payload_from_user_row", return_value={"campo": "valor"}):
                with patch.object(
                    seguimientos,
                    "_build_base_sheet_updates",
                    return_value=[{"range": "A1", "value": "demo"}],
                ) as build_base_sheet_updates:
                    with patch.object(seguimientos, "clear_protected_ranges") as clear_protected_ranges:
                        with patch.object(seguimientos, "batch_write_sheet_updates") as batch_write_sheet_updates:
                            with patch.object(seguimientos, "_set_sheet_visibility") as set_sheet_visibility:
                                with patch.object(
                                    seguimientos,
                                    "get_spreadsheet",
                                    return_value={
                                        "sheets": [
                                            {"properties": {"title": seguimientos.SHEET_BASE}}
                                        ]
                                    },
                                ):
                                    record = seguimientos._create_native_case_record(
                                        service,
                                        "folder-id",
                                        "Caso Demo",
                                        "123",
                                        {"nombre_usuario": "Persona Demo"},
                                        3,
                                    )

        self.assertEqual(record["file_id"], "seguimiento-copy-id")
        build_base_sheet_updates.assert_called_once_with(
            {"campo": "valor"},
            base_sheet_name=seguimientos.SHEET_BASE,
        )
        clear_protected_ranges.assert_called_once_with("seguimiento-copy-id")
        batch_write_sheet_updates.assert_called_once()
        written_updates = batch_write_sheet_updates.call_args.args[1]
        self.assertIn({"range": "A1", "value": "demo"}, written_updates)
        self.assertIn({"range": "'SEGUIMIENTO PROCESO IL 1'!O12", "value": ""}, written_updates)
        self.assertIn({"range": "'SEGUIMIENTO PROCESO IL 6'!N50", "value": ""}, written_updates)
        set_sheet_visibility.assert_called_once_with("seguimiento-copy-id", 3)

    def test_set_sheet_visibility_hides_non_seguimiento_tabs(self) -> None:
        spreadsheet = {
            "sheets": [
                {"properties": {"sheetId": 1, "title": seguimientos.SHEET_BASE}},
                {"properties": {"sheetId": 2, "title": "Caracterización"}},
                {"properties": {"sheetId": 3, "title": "SEGUIMIENTO PROCESO IL 1"}},
                {"properties": {"sheetId": 4, "title": "SEGUIMIENTO PROCESO IL 6"}},
                {"properties": {"sheetId": 5, "title": seguimientos.SHEET_FINAL}},
            ]
        }
        service = Mock()
        request = object()
        service.spreadsheets.return_value.batchUpdate.return_value = request

        with patch.object(seguimientos, "get_google_sheets_service", return_value=service):
            with patch.object(seguimientos, "get_spreadsheet", return_value=spreadsheet):
                with patch.object(seguimientos, "execute_google_request_with_retry", return_value={}) as execute_mock:
                    seguimientos._set_sheet_visibility("spreadsheet-demo", 3)

        service.spreadsheets.return_value.batchUpdate.assert_called_once()
        body = service.spreadsheets.return_value.batchUpdate.call_args.kwargs["body"]
        requests = body["requests"]
        hidden_by_id = {
            item["updateSheetProperties"]["properties"]["sheetId"]: item["updateSheetProperties"]["properties"]["hidden"]
            for item in requests
        }
        self.assertFalse(hidden_by_id[1])
        self.assertTrue(hidden_by_id[2])
        self.assertFalse(hidden_by_id[3])
        self.assertFalse(hidden_by_id[4])
        self.assertFalse(hidden_by_id[5])
        execute_mock.assert_called_once_with(
            request,
            operation_name="seguimientos.set_sheet_visibility",
        )


if __name__ == "__main__":
    unittest.main()
