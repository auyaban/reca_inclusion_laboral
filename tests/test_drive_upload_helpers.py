from __future__ import annotations

from contextlib import ExitStack
import unittest
from unittest import mock

import drive_upload


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
        self.create_responses = []
        self.copy_responses = []
        self.update_responses = []
        self.list_calls = []
        self.create_calls = []
        self.copy_calls = []
        self.update_calls = []

    def list(self, **kwargs):
        self.list_calls.append(kwargs)
        return _FakeQueuedRequest(self.list_responses)

    def create(self, **kwargs):
        self.create_calls.append(kwargs)
        return _FakeQueuedRequest(self.create_responses)

    def copy(self, **kwargs):
        self.copy_calls.append(kwargs)
        return _FakeQueuedRequest(self.copy_responses)

    def update(self, **kwargs):
        self.update_calls.append(kwargs)
        return _FakeQueuedRequest(self.update_responses)


class _FakeDriveService:
    def __init__(self):
        self.files_resource = _FakeFilesResource()

    def files(self):
        return self.files_resource


class DriveUploadHelperTests(unittest.TestCase):
    def test_build_dated_sheet_title_adds_numeric_suffix_when_needed(self) -> None:
        title = drive_upload._build_dated_sheet_title(
            "8. SENSIBILIZACION",
            {
                "8. SENSIBILIZACION",
                "8. SENSIBILIZACION - 2026-03-27",
            },
            current_date="2026-03-27",
        )
        self.assertEqual(title, "8. SENSIBILIZACION - 2026-03-27 (2)")

    def test_rewrite_sheet_payloads_updates_all_sheet_references(self) -> None:
        rewrites = {"5. CONTRATACION": "5. CONTRATACION - 2026-03-27"}
        writes, clear_ranges, checkboxes, unmerge, row_insertions = drive_upload._rewrite_sheet_payloads(
            [{"range": "'5. CONTRATACION'!A1", "value": "demo"}],
            ["'5. CONTRATACION'!A1:B5"],
            [{"range": "'5. CONTRATACION'!C2", "value": True}],
            [{"sheet_name": "5. CONTRATACION", "start_row": 0, "end_row": 10}],
            [{"sheet_name": "5. CONTRATACION", "start_row": 20, "base_rows": 3, "total_rows": 5}],
            rewrites,
        )

        self.assertEqual(writes[0]["range"], "'5. CONTRATACION - 2026-03-27'!A1")
        self.assertEqual(clear_ranges[0], "'5. CONTRATACION - 2026-03-27'!A1:B5")
        self.assertEqual(checkboxes[0]["range"], "'5. CONTRATACION - 2026-03-27'!C2")
        self.assertEqual(unmerge[0]["sheet_name"], "5. CONTRATACION - 2026-03-27")
        self.assertEqual(row_insertions[0]["sheet_name"], "5. CONTRATACION - 2026-03-27")

    def test_all_target_ranges_populated_requires_full_coverage(self) -> None:
        value_ranges = {
            "'3. REVISION'!A1": [["Dato"]],
            "'3. REVISION'!B1": [[""]],
            "'3. REVISION'!C1": [[0]],
        }

        self.assertFalse(
            drive_upload._all_target_ranges_populated(
                value_ranges,
                ["'3. REVISION'!A1", "'3. REVISION'!B1"],
            )
        )
        self.assertTrue(
            drive_upload._all_target_ranges_populated(
                value_ranges,
                ["'3. REVISION'!A1", "'3. REVISION'!C1"],
            )
        )

    def test_publish_sheet_from_template_copies_missing_sheet_on_reuse(self) -> None:
        sheet_writes = [{"range": "'Nueva'!A1", "value": "demo"}]

        with self._publish_common_patches():
            with (
                mock.patch.object(drive_upload, "_find_existing_spreadsheet", return_value="existing-id"),
                mock.patch("google_sheets_client.get_sheet_titles", return_value=["Anterior"]),
                mock.patch("google_sheets_client.copy_sheet_to_spreadsheet", return_value={"sheetId": 321, "title": "Nueva"}) as copy_mock,
                mock.patch("google_sheets_client.clear_protected_ranges") as clear_mock,
                mock.patch("google_sheets_client.batch_write_sheet_updates") as batch_write_mock,
                mock.patch("google_sheets_client.hide_sheets") as hide_mock,
                mock.patch("google_sheets_client.batch_read_sheet_values") as read_mock,
                mock.patch("google_sheets_client.clear_sheet_ranges") as clear_ranges_mock,
                mock.patch("google_sheets_client.insert_template_rows") as insert_rows_mock,
                mock.patch("google_sheets_client.insert_template_block_rows") as insert_block_rows_mock,
                mock.patch("google_sheets_client.set_native_checkboxes") as checkboxes_mock,
                mock.patch("google_sheets_client.unmerge_cells_in_area") as unmerge_mock,
            ):
                result = drive_upload.publish_sheet_from_template(
                    template_id="template-id",
                    sheet_writes=sheet_writes,
                    base_name="Caso",
                )

        self.assertEqual(result["file_id"], "existing-id")
        copy_mock.assert_called_once_with(
            "template-id",
            "Nueva",
            "existing-id",
            new_sheet_name="Nueva",
        )
        clear_mock.assert_called_once_with("existing-id")
        batch_write_mock.assert_called_once_with(
            "existing-id",
            sheet_writes,
            auto_resize_excluded_rows=None,
        )
        self.assertEqual(set(hide_mock.call_args.args[1]), {"Nueva"})
        read_mock.assert_not_called()
        clear_ranges_mock.assert_not_called()
        insert_rows_mock.assert_not_called()
        insert_block_rows_mock.assert_not_called()
        checkboxes_mock.assert_not_called()
        unmerge_mock.assert_not_called()

    def test_publish_sheet_from_template_copies_occupied_sheet_and_rewrites_payloads(self) -> None:
        expected_title = "5. CONTRATACION - 2026-04-01"
        sheet_writes = [{"range": "'5. CONTRATACION'!A1", "value": "demo"}]
        clear_ranges = ["'5. CONTRATACION'!A1:B5"]
        checkbox_cells = [{"range": "'5. CONTRATACION'!C2", "value": True}]
        unmerge_areas = [{"sheet_name": "5. CONTRATACION", "start_row": 0, "end_row": 10}]
        row_insertions = [{"sheet_name": "5. CONTRATACION", "start_row": 20, "base_rows": 3, "total_rows": 5}]

        with self._publish_common_patches():
            with (
                mock.patch.object(drive_upload, "_find_existing_spreadsheet", return_value="existing-id"),
                mock.patch("google_sheets_client.get_sheet_titles", return_value=["5. CONTRATACION"]),
                mock.patch(
                    "google_sheets_client.batch_read_sheet_values",
                    return_value={
                        "'5. CONTRATACION'!A1": [["Dato"]],
                        "'5. CONTRATACION'!C2": [["TRUE"]],
                    },
                ),
                mock.patch.object(drive_upload, "_build_dated_sheet_title", return_value=expected_title),
                mock.patch(
                    "google_sheets_client.copy_sheet_to_spreadsheet",
                    return_value={"sheetId": 222, "title": expected_title},
                ) as copy_mock,
                mock.patch("google_sheets_client.clear_protected_ranges"),
                mock.patch("google_sheets_client.batch_write_sheet_updates") as batch_write_mock,
                mock.patch("google_sheets_client.clear_sheet_ranges") as clear_ranges_mock,
                mock.patch("google_sheets_client.set_native_checkboxes") as checkboxes_mock,
                mock.patch("google_sheets_client.unmerge_cells_in_area") as unmerge_mock,
                mock.patch("google_sheets_client.insert_template_rows") as insert_rows_mock,
                mock.patch("google_sheets_client.insert_template_block_rows") as insert_block_rows_mock,
                mock.patch("google_sheets_client.hide_sheets") as hide_mock,
            ):
                result = drive_upload.publish_sheet_from_template(
                    template_id="template-id",
                    sheet_writes=sheet_writes,
                    base_name="Caso",
                    clear_ranges=clear_ranges,
                    checkbox_cells=checkbox_cells,
                    unmerge_areas=unmerge_areas,
                    row_insertions=row_insertions,
                )

        rewritten_write = [{"range": f"'{expected_title}'!A1", "value": "demo"}]
        self.assertEqual(result["file_id"], "existing-id")
        copy_mock.assert_called_once_with(
            "template-id",
            "5. CONTRATACION",
            "existing-id",
            new_sheet_name=expected_title,
        )
        batch_write_mock.assert_called_once_with(
            "existing-id",
            rewritten_write,
            auto_resize_excluded_rows={},
        )
        clear_ranges_mock.assert_called_once_with("existing-id", [f"'{expected_title}'!A1:B5"])
        checkboxes_mock.assert_called_once_with(
            "existing-id",
            [{"range": f"'{expected_title}'!C2", "value": True}],
        )
        unmerge_mock.assert_called_once_with("existing-id", expected_title, 0, 10, 0, 21)
        insert_rows_mock.assert_called_once_with(
            "existing-id",
            expected_title,
            insert_at_row=23,
            template_row=22,
            count=2,
            paste_type="PASTE_NORMAL",
        )
        insert_block_rows_mock.assert_not_called()
        self.assertEqual(set(hide_mock.call_args.args[1]), {expected_title})

    def test_publish_sheet_from_template_reuses_only_when_target_ranges_are_empty(self) -> None:
        sheet_writes = [{"range": "'5. CONTRATACION'!A1", "value": "demo"}]

        with self._publish_common_patches():
            with (
                mock.patch.object(drive_upload, "_find_existing_spreadsheet", return_value="existing-id"),
                mock.patch("google_sheets_client.get_sheet_titles", return_value=["5. CONTRATACION"]),
                mock.patch(
                    "google_sheets_client.batch_read_sheet_values",
                    return_value={"'5. CONTRATACION'!A1": [[""]]},
                ),
                mock.patch("google_sheets_client.copy_sheet_to_spreadsheet") as copy_mock,
                mock.patch("google_sheets_client.clear_protected_ranges"),
                mock.patch("google_sheets_client.batch_write_sheet_updates") as batch_write_mock,
                mock.patch("google_sheets_client.hide_sheets") as hide_mock,
                mock.patch("google_sheets_client.clear_sheet_ranges"),
                mock.patch("google_sheets_client.insert_template_rows"),
                mock.patch("google_sheets_client.insert_template_block_rows"),
                mock.patch("google_sheets_client.set_native_checkboxes"),
                mock.patch("google_sheets_client.unmerge_cells_in_area"),
            ):
                result = drive_upload.publish_sheet_from_template(
                    template_id="template-id",
                    sheet_writes=sheet_writes,
                    base_name="Caso",
                )

        self.assertEqual(result["file_id"], "existing-id")
        copy_mock.assert_not_called()
        batch_write_mock.assert_called_once_with(
            "existing-id",
            sheet_writes,
            auto_resize_excluded_rows=None,
        )
        self.assertEqual(set(hide_mock.call_args.args[1]), {"5. CONTRATACION"})

    def test_publish_sheet_from_template_copies_dated_sheet_when_any_target_range_is_occupied(self) -> None:
        expected_title = "5. CONTRATACION - 2026-04-03"
        sheet_writes = [{"range": "'5. CONTRATACION'!A1", "value": "demo"}]

        with self._publish_common_patches():
            with (
                mock.patch.object(drive_upload, "_find_existing_spreadsheet", return_value="existing-id"),
                mock.patch("google_sheets_client.get_sheet_titles", return_value=["5. CONTRATACION"]),
                mock.patch(
                    "google_sheets_client.batch_read_sheet_values",
                    return_value={"'5. CONTRATACION'!A1": [["Dato existente"]]},
                ),
                mock.patch.object(drive_upload, "_build_dated_sheet_title", return_value=expected_title),
                mock.patch(
                    "google_sheets_client.copy_sheet_to_spreadsheet",
                    return_value={"sheetId": 999, "title": expected_title},
                ) as copy_mock,
                mock.patch("google_sheets_client.clear_protected_ranges"),
                mock.patch("google_sheets_client.batch_write_sheet_updates") as batch_write_mock,
                mock.patch("google_sheets_client.hide_sheets") as hide_mock,
                mock.patch("google_sheets_client.clear_sheet_ranges"),
                mock.patch("google_sheets_client.insert_template_rows"),
                mock.patch("google_sheets_client.insert_template_block_rows"),
                mock.patch("google_sheets_client.set_native_checkboxes"),
                mock.patch("google_sheets_client.unmerge_cells_in_area"),
            ):
                result = drive_upload.publish_sheet_from_template(
                    template_id="template-id",
                    sheet_writes=sheet_writes,
                    base_name="Caso",
                )

        self.assertEqual(result["file_id"], "existing-id")
        copy_mock.assert_called_once_with(
            "template-id",
            "5. CONTRATACION",
            "existing-id",
            new_sheet_name=expected_title,
        )
        batch_write_mock.assert_called_once_with(
            "existing-id",
            [{"range": f"'{expected_title}'!A1", "value": "demo"}],
            auto_resize_excluded_rows={},
        )
        self.assertEqual(set(hide_mock.call_args.args[1]), {expected_title})

    def test_get_or_create_folder_returns_confirmed_folder_after_ambiguous_create_error(self) -> None:
        service = _FakeDriveService()
        service.files_resource.list_responses = [
            {"files": []},
            {"files": [{"id": "folder-123", "name": "Empresa"}]},
        ]
        service.files_resource.create_responses = [_FakeGoogleError(503)]

        with mock.patch("google_api_requests.time.sleep") as sleep_mock:
            folder_id = drive_upload._get_or_create_folder(service, "root-folder", "Empresa")

        self.assertEqual(folder_id, "folder-123")
        self.assertEqual(len(service.files_resource.create_calls), 1)
        self.assertEqual(len(service.files_resource.list_calls), 2)
        sleep_mock.assert_not_called()

    def test_publish_sheet_from_template_new_copy_uses_confirmed_resource_after_ambiguous_error(self) -> None:
        fake_service = _FakeDriveService()
        fake_service.files_resource.copy_responses = [_FakeGoogleError(503)]
        fake_service.files_resource.list_responses = [
            {
                "files": [
                    {
                        "id": "sheet-123",
                        "name": "Caso",
                        "mimeType": drive_upload.GOOGLE_SHEETS_MIME,
                        "webViewLink": "https://docs.google.com/spreadsheets/d/sheet-123/edit",
                        "appProperties": {"request_id": "req-sheet-123"},
                    }
                ]
            }
        ]

        with (
            mock.patch("google.oauth2.service_account.Credentials.from_service_account_file", return_value=object()),
            mock.patch("googleapiclient.discovery.build", return_value=fake_service),
            mock.patch.object(drive_upload, "_get_credentials_path", return_value="creds.json"),
            mock.patch.object(drive_upload, "_get_excel_folder_id", return_value="root-folder"),
            mock.patch.object(drive_upload, "_resolve_target_root_id", return_value="root-folder"),
            mock.patch.object(drive_upload, "_find_existing_spreadsheet", return_value=None),
            mock.patch.object(drive_upload, "_get_available_filename", return_value="Caso"),
            mock.patch.object(drive_upload.uuid, "uuid4", return_value=mock.Mock(hex="req-sheet-123")),
            mock.patch("google_sheets_client.clear_protected_ranges"),
            mock.patch("google_sheets_client.batch_write_sheet_updates") as batch_write_mock,
            mock.patch("google_sheets_client.hide_sheets") as hide_mock,
            mock.patch("google_sheets_client.clear_sheet_ranges"),
            mock.patch("google_sheets_client.insert_template_rows"),
            mock.patch("google_sheets_client.insert_template_block_rows"),
            mock.patch("google_sheets_client.set_native_checkboxes"),
            mock.patch("google_sheets_client.unmerge_cells_in_area"),
            mock.patch("google_sheets_client.batch_read_sheet_values"),
            mock.patch("google_sheets_client.get_sheet_titles"),
            mock.patch("google_sheets_client.copy_sheet_to_spreadsheet"),
        ):
            result = drive_upload.publish_sheet_from_template(
                template_id="template-id",
                sheet_writes=[{"range": "'Nueva'!A1", "value": "demo"}],
                base_name="Caso",
            )

        self.assertEqual(result["file_id"], "sheet-123")
        self.assertEqual(len(fake_service.files_resource.copy_calls), 1)
        self.assertEqual(
            fake_service.files_resource.copy_calls[0]["body"]["appProperties"]["request_id"],
            "req-sheet-123",
        )
        self.assertEqual(
            fake_service.files_resource.copy_calls[0]["body"]["appProperties"]["kind"],
            "google_sheet_publish",
        )
        batch_write_mock.assert_called_once_with(
            "sheet-123",
            [{"range": "'Nueva'!A1", "value": "demo"}],
            auto_resize_excluded_rows=None,
        )
        self.assertEqual(set(hide_mock.call_args.args[1]), {"Nueva"})

    def test_upload_excel_to_drive_uses_confirmed_resource_after_ambiguous_error(self) -> None:
        fake_service = _FakeDriveService()
        fake_service.files_resource.create_responses = [_FakeGoogleError(503)]
        fake_service.files_resource.list_responses = [
            {
                "files": [
                    {
                        "id": "file-123",
                        "name": "archivo.xlsx",
                        "mimeType": drive_upload.XLSX_MIME,
                        "webViewLink": "https://drive.google.com/file/d/file-123/view",
                        "appProperties": {"request_id": "req-file-123"},
                    }
                ]
            }
        ]

        with (
            mock.patch.object(drive_upload.os.path, "exists", return_value=True),
            mock.patch("google.oauth2.service_account.Credentials.from_service_account_file", return_value=object()),
            mock.patch("googleapiclient.discovery.build", return_value=fake_service),
            mock.patch("googleapiclient.http.MediaFileUpload", return_value=object()),
            mock.patch.object(drive_upload, "_get_credentials_path", return_value="creds.json"),
            mock.patch.object(drive_upload, "_get_excel_folder_id", return_value="root-folder"),
            mock.patch.object(drive_upload, "_resolve_target_root_id", return_value="root-folder"),
            mock.patch.object(drive_upload, "_get_available_filename", return_value="archivo.xlsx"),
            mock.patch.object(drive_upload.uuid, "uuid4", return_value=mock.Mock(hex="req-file-123")),
        ):
            result = drive_upload.upload_excel_to_drive(r"C:\tmp\archivo.xlsx")

        self.assertEqual(result["file_id"], "file-123")
        self.assertEqual(len(fake_service.files_resource.create_calls), 1)
        self.assertEqual(
            fake_service.files_resource.create_calls[0]["body"]["appProperties"]["request_id"],
            "req-file-123",
        )
        self.assertEqual(
            fake_service.files_resource.create_calls[0]["body"]["appProperties"]["kind"],
            "excel_upload",
        )

    def _publish_common_patches(self):
        return _PublishPatches()


class _PublishPatches:
    def __enter__(self):
        self._stack = ExitStack()
        self._stack.enter_context(
            mock.patch(
                "google.oauth2.service_account.Credentials.from_service_account_file",
                return_value=object(),
            )
        )
        self._stack.enter_context(
            mock.patch("googleapiclient.discovery.build", return_value=object())
        )
        self._stack.enter_context(
            mock.patch.object(drive_upload, "_get_credentials_path", return_value="creds.json")
        )
        self._stack.enter_context(
            mock.patch.object(drive_upload, "_get_excel_folder_id", return_value="root-folder")
        )
        self._stack.enter_context(
            mock.patch.object(drive_upload, "_resolve_target_root_id", return_value="root-folder")
        )
        return self

    def __exit__(self, exc_type, exc, tb):
        return self._stack.__exit__(exc_type, exc, tb)


if __name__ == "__main__":
    unittest.main()
