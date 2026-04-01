from __future__ import annotations

import unittest
from unittest.mock import patch

import google_sheets_client as sheets_client


class _FakeExecute:
    def __init__(self, payload):
        self._payload = payload

    def execute(self):
        return self._payload


class _FakeSpreadsheetsResource:
    def __init__(self, meta_payload):
        self._meta_payload = meta_payload
        self.batch_bodies = []

    def get(self, **_kwargs):
        return _FakeExecute(self._meta_payload)

    def batchUpdate(self, **kwargs):
        self.batch_bodies.append(kwargs.get("body") or {})
        return _FakeExecute({})


class _FakeService:
    def __init__(self, meta_payload):
        self._spreadsheets = _FakeSpreadsheetsResource(meta_payload)

    def spreadsheets(self):
        return self._spreadsheets


class GoogleSheetsClientTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
