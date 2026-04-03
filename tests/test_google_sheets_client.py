from __future__ import annotations

import unittest
from unittest.mock import patch

import google_sheets_client as sheets_client


class _FakeExecute:
    def __init__(self, payload):
        self._payload = payload

    def execute(self):
        return self._payload


class _FakeValuesResource:
    def __init__(self):
        self.batch_bodies = []
        self.update_calls = []

    def batchUpdate(self, **kwargs):
        self.batch_bodies.append(kwargs.get("body") or {})
        return _FakeExecute({"totalUpdatedCells": 1, "totalUpdatedRows": 1, "responses": []})

    def update(self, **kwargs):
        self.update_calls.append(kwargs)
        return _FakeExecute({"updatedRange": kwargs.get("range")})


class _FakeSpreadsheetsResource:
    def __init__(self, meta_payload):
        self._meta_payload = meta_payload
        self.batch_bodies = []
        self._values = _FakeValuesResource()

    def get(self, **_kwargs):
        return _FakeExecute(self._meta_payload)

    def batchUpdate(self, **kwargs):
        self.batch_bodies.append(kwargs.get("body") or {})
        return _FakeExecute({})

    def values(self):
        return self._values


class _FakeService:
    def __init__(self, meta_payload):
        self._spreadsheets = _FakeSpreadsheetsResource(meta_payload)

    def spreadsheets(self):
        return self._spreadsheets


class GoogleSheetsClientTests(unittest.TestCase):
    def test_filter_auto_resize_ranges_skips_excluded_rows(self) -> None:
        filtered = sheets_client._filter_auto_resize_ranges(
            [
                "'5. CONTRATACIÓN INCLUYENTE'!A17",
                "'5. CONTRATACIÓN INCLUYENTE'!M66",
                "'5. CONTRATACIÓN INCLUYENTE'!M67",
                "'5. CONTRATACIÓN INCLUYENTE'!C20",
                "'5. CONTRATACIÓN INCLUYENTE - 2026-04-03'!A17",
            ],
            excluded_rows_by_sheet={
                "5. CONTRATACIÓN INCLUYENTE": [17, 66, 67],
                "5. CONTRATACIÓN INCLUYENTE - 2026-04-03": [17],
            },
        )

        self.assertEqual(filtered, ["'5. CONTRATACIÓN INCLUYENTE'!C20"])

    def test_insert_template_block_rows_unmerges_inserted_area_before_copy(self) -> None:
        meta = {"sheets": [{"properties": {"sheetId": 9, "title": "5. CONTRATACIÓN INCLUYENTE"}}]}
        fake_service = _FakeService(meta)

        with (
            patch.object(sheets_client, "extract_spreadsheet_id", return_value="sheet-id"),
            patch.object(sheets_client, "get_google_sheets_service", return_value=fake_service),
            patch.object(sheets_client, "unmerge_cells_in_area") as unmerge_mock,
            patch.object(
                sheets_client,
                "execute_google_request_with_retry",
                side_effect=lambda request, **_kwargs: request.execute(),
            ),
        ):
            result = sheets_client.insert_template_block_rows(
                "sheet-id",
                "5. CONTRATACIÓN INCLUYENTE",
                insert_at_row=68,
                template_start_row=16,
                template_end_row=67,
                repeat_count=2,
            )

        self.assertEqual(result["insertedRows"], 104)
        self.assertEqual(result["insertedBlocks"], 2)
        self.assertEqual(len(fake_service._spreadsheets.batch_bodies), 2)

        insert_requests = fake_service._spreadsheets.batch_bodies[0]["requests"]
        self.assertEqual(len(insert_requests), 1)
        self.assertEqual(
            insert_requests[0]["insertDimension"]["range"],
            {
                "sheetId": 9,
                "dimension": "ROWS",
                "startIndex": 67,
                "endIndex": 171,
            },
        )

        unmerge_mock.assert_called_once_with(
            "sheet-id",
            "5. CONTRATACIÓN INCLUYENTE",
            67,
            171,
        )

        copy_requests = fake_service._spreadsheets.batch_bodies[1]["requests"]
        self.assertEqual(len(copy_requests), 2)
        self.assertEqual(copy_requests[0]["copyPaste"]["source"]["startRowIndex"], 15)
        self.assertEqual(copy_requests[0]["copyPaste"]["source"]["endRowIndex"], 67)
        self.assertEqual(copy_requests[0]["copyPaste"]["destination"]["startRowIndex"], 67)
        self.assertEqual(copy_requests[0]["copyPaste"]["destination"]["endRowIndex"], 119)
        self.assertEqual(copy_requests[1]["copyPaste"]["destination"]["startRowIndex"], 119)
        self.assertEqual(copy_requests[1]["copyPaste"]["destination"]["endRowIndex"], 171)

    def test_auto_resize_rows_for_ranges_merges_overlapping_segments(self) -> None:
        meta = {
            "sheets": [
                {"properties": {"sheetId": 7, "title": "Base"}},
                {"properties": {"sheetId": 8, "title": "Seguimiento"}},
            ]
        }
        fake_service = _FakeService(meta)

        with (
            patch.object(sheets_client, "extract_spreadsheet_id", return_value="sheet-id"),
            patch.object(sheets_client, "get_google_sheets_service", return_value=fake_service),
            patch.object(
                sheets_client,
                "execute_google_request_with_retry",
                side_effect=lambda request, **_kwargs: request.execute(),
            ),
        ):
            sheets_client.auto_resize_rows_for_ranges(
                "sheet-id",
                [
                    "'Base'!A2",
                    "'Base'!C2:D3",
                    "'Base'!B5",
                    "'Seguimiento'!A10:A12",
                ],
            )

        body = fake_service._spreadsheets.batch_bodies[0]
        requests = body["requests"]
        self.assertEqual(
            requests,
            [
                {
                    "autoResizeDimensions": {
                        "dimensions": {
                            "sheetId": 7,
                            "dimension": "ROWS",
                            "startIndex": 1,
                            "endIndex": 3,
                        }
                    }
                },
                {
                    "autoResizeDimensions": {
                        "dimensions": {
                            "sheetId": 7,
                            "dimension": "ROWS",
                            "startIndex": 4,
                            "endIndex": 5,
                        }
                    }
                },
                {
                    "autoResizeDimensions": {
                        "dimensions": {
                            "sheetId": 8,
                            "dimension": "ROWS",
                            "startIndex": 9,
                            "endIndex": 12,
                        }
                    }
                },
            ],
        )

    def test_batch_write_sheet_updates_resizes_written_rows(self) -> None:
        meta = {
            "sheets": [
                {"properties": {"sheetId": 3, "title": "Base"}},
            ]
        }
        fake_service = _FakeService(meta)

        with (
            patch.object(sheets_client, "extract_spreadsheet_id", return_value="sheet-id"),
            patch.object(sheets_client, "get_google_sheets_service", return_value=fake_service),
            patch.object(
                sheets_client,
                "execute_google_request_with_retry",
                side_effect=lambda request, **_kwargs: request.execute(),
            ),
        ):
            sheets_client.batch_write_sheet_updates(
                "sheet-id",
                [
                    {"range": "'Base'!A2", "value": "Texto largo"},
                    {"range": "'Base'!C7", "value": "Otro valor"},
                ],
            )

        values_body = fake_service._spreadsheets._values.batch_bodies[0]
        self.assertEqual(
            values_body["data"],
            [
                {
                    "range": "'Base'!A2",
                    "majorDimension": "ROWS",
                    "values": [["Texto largo"]],
                },
                {
                    "range": "'Base'!C7",
                    "majorDimension": "ROWS",
                    "values": [["Otro valor"]],
                },
            ],
        )
        self.assertEqual(values_body["valueInputOption"], "RAW")
        resize_body = fake_service._spreadsheets.batch_bodies[0]
        self.assertEqual(
            resize_body["requests"],
            [
                {
                    "autoResizeDimensions": {
                        "dimensions": {
                            "sheetId": 3,
                            "dimension": "ROWS",
                            "startIndex": 1,
                            "endIndex": 2,
                        }
                    }
                },
                {
                    "autoResizeDimensions": {
                        "dimensions": {
                            "sheetId": 3,
                            "dimension": "ROWS",
                            "startIndex": 6,
                            "endIndex": 7,
                        }
                    }
                },
            ],
        )

    def test_batch_write_sheet_updates_skips_auto_resize_for_excluded_rows(self) -> None:
        meta = {
            "sheets": [
                {"properties": {"sheetId": 3, "title": "5. CONTRATACIÓN INCLUYENTE"}},
            ]
        }
        fake_service = _FakeService(meta)

        with (
            patch.object(sheets_client, "extract_spreadsheet_id", return_value="sheet-id"),
            patch.object(sheets_client, "get_google_sheets_service", return_value=fake_service),
            patch.object(
                sheets_client,
                "execute_google_request_with_retry",
                side_effect=lambda request, **_kwargs: request.execute(),
            ),
        ):
            sheets_client.batch_write_sheet_updates(
                "sheet-id",
                [
                    {"range": "'5. CONTRATACIÓN INCLUYENTE'!A17", "value": "Titulo"},
                    {"range": "'5. CONTRATACIÓN INCLUYENTE'!C20", "value": "Dato"},
                    {"range": "'5. CONTRATACIÓN INCLUYENTE'!M66", "value": "Nota"},
                ],
                auto_resize_excluded_rows={"5. CONTRATACIÓN INCLUYENTE": [17, 66, 67]},
            )

        resize_body = fake_service._spreadsheets.batch_bodies[0]
        self.assertEqual(
            resize_body["requests"],
            [
                {
                    "autoResizeDimensions": {
                        "dimensions": {
                            "sheetId": 3,
                            "dimension": "ROWS",
                            "startIndex": 19,
                            "endIndex": 20,
                        }
                    }
                }
            ],
        )

    def test_hide_sheets_unhides_target_before_hiding_previous_visible_sheet(self) -> None:
        meta = {
            "sheets": [
                {"properties": {"sheetId": 1, "title": "Anterior", "hidden": False}},
                {"properties": {"sheetId": 2, "title": "Actual", "hidden": True}},
            ]
        }
        fake_service = _FakeService(meta)

        with patch.object(sheets_client, "extract_spreadsheet_id", return_value="sheet-id"):
            with patch.object(sheets_client, "get_google_sheets_service", return_value=fake_service):
                sheets_client.hide_sheets("sheet-id", ["Actual"])

        body = fake_service._spreadsheets.batch_bodies[0]
        requests = body["requests"]
        self.assertEqual(len(requests), 2)
        self.assertFalse(requests[0]["updateSheetProperties"]["properties"]["hidden"])
        self.assertEqual(requests[0]["updateSheetProperties"]["properties"]["sheetId"], 2)
        self.assertTrue(requests[1]["updateSheetProperties"]["properties"]["hidden"])
        self.assertEqual(requests[1]["updateSheetProperties"]["properties"]["sheetId"], 1)

    def test_hide_sheets_raises_when_keep_sheet_does_not_exist(self) -> None:
        meta = {
            "sheets": [
                {"properties": {"sheetId": 1, "title": "Anterior", "hidden": False}},
                {"properties": {"sheetId": 2, "title": "Archivada", "hidden": True}},
            ]
        }
        fake_service = _FakeService(meta)

        with patch.object(sheets_client, "extract_spreadsheet_id", return_value="sheet-id"):
            with patch.object(sheets_client, "get_google_sheets_service", return_value=fake_service):
                with self.assertRaisesRegex(RuntimeError, "No existe ninguna hoja con el nombre solicitado"):
                    sheets_client.hide_sheets("sheet-id", ["Actual"])

        self.assertEqual(fake_service._spreadsheets.batch_bodies, [])

    def test_batch_write_sheet_updates_keeps_formula_like_values_literal_in_raw_mode(self) -> None:
        meta = {"sheets": [{"properties": {"sheetId": 3, "title": "Base"}}]}
        fake_service = _FakeService(meta)

        with (
            patch.object(sheets_client, "extract_spreadsheet_id", return_value="sheet-id"),
            patch.object(sheets_client, "get_google_sheets_service", return_value=fake_service),
            patch.object(
                sheets_client,
                "execute_google_request_with_retry",
                side_effect=lambda request, **_kwargs: request.execute(),
            ),
        ):
            sheets_client.batch_write_sheet_updates(
                "sheet-id",
                [{"range": "'Base'!A2", "value": "=HYPERLINK(\"https://evil.test\")"}],
            )

        body = fake_service._spreadsheets._values.batch_bodies[0]
        self.assertEqual(body["data"][0]["values"], [["=HYPERLINK(\"https://evil.test\")"]])
        self.assertEqual(body["valueInputOption"], "RAW")
    
    def test_batch_write_sheet_updates_keeps_phone_and_negative_values_in_raw_mode(self) -> None:
        meta = {"sheets": [{"properties": {"sheetId": 3, "title": "Base"}}]}
        fake_service = _FakeService(meta)

        with (
            patch.object(sheets_client, "extract_spreadsheet_id", return_value="sheet-id"),
            patch.object(sheets_client, "get_google_sheets_service", return_value=fake_service),
            patch.object(
                sheets_client,
                "execute_google_request_with_retry",
                side_effect=lambda request, **_kwargs: request.execute(),
            ),
        ):
            sheets_client.batch_write_sheet_updates(
                "sheet-id",
                [
                    {"range": "'Base'!A2", "value": "+57 300 123 4567"},
                    {"range": "'Base'!A3", "value": "-500000"},
                ],
            )

        body = fake_service._spreadsheets._values.batch_bodies[0]
        self.assertEqual(body["data"][0]["values"], [["+57 300 123 4567"]])
        self.assertEqual(body["data"][1]["values"], [["-500000"]])

    def test_write_sheet_values_uses_raw_and_sanitizes_formula_prefixes(self) -> None:
        fake_service = _FakeService({"sheets": []})

        with (
            patch.object(sheets_client, "extract_spreadsheet_id", return_value="sheet-id"),
            patch.object(sheets_client, "get_google_sheets_service", return_value=fake_service),
            patch.object(
                sheets_client,
                "execute_google_request_with_retry",
                side_effect=lambda request, **_kwargs: request.execute(),
            ),
        ):
            sheets_client.write_sheet_values(
                "sheet-id",
                "'Base'!A1:B1",
                [["=SUM(1,2)", "texto"]],
            )

        update_call = fake_service._spreadsheets._values.update_calls[0]
        self.assertEqual(update_call["valueInputOption"], "RAW")
        self.assertEqual(update_call["body"]["values"], [["=SUM(1,2)", "texto"]])

    def test_write_sheet_values_escapes_formula_prefixes_only_for_user_entered(self) -> None:
        fake_service = _FakeService({"sheets": []})

        with (
            patch.object(sheets_client, "extract_spreadsheet_id", return_value="sheet-id"),
            patch.object(sheets_client, "get_google_sheets_service", return_value=fake_service),
            patch.object(
                sheets_client,
                "execute_google_request_with_retry",
                side_effect=lambda request, **_kwargs: request.execute(),
            ),
        ):
            sheets_client.write_sheet_values(
                "sheet-id",
                "'Base'!A1",
                [["=SUM(1,2)"]],
                value_input_option="USER_ENTERED",
            )

        update_call = fake_service._spreadsheets._values.update_calls[0]
        self.assertEqual(update_call["valueInputOption"], "USER_ENTERED")
        self.assertEqual(update_call["body"]["values"], [["'=SUM(1,2)"]])


if __name__ == "__main__":
    unittest.main()
