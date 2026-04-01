from __future__ import annotations

import unittest

import drive_upload


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


if __name__ == "__main__":
    unittest.main()
