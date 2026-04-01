from __future__ import annotations

import unittest

from formularios.presentacion_programa import presentacion_programa


class PresentacionProgramaRowInsertionsTests(unittest.TestCase):
    def test_build_section_5_row_insertions_only_when_attendees_exceed_base_rows(self) -> None:
        payload = [
            {"nombre": "Uno", "cargo": "Cargo 1"},
            {"nombre": "Dos", "cargo": "Cargo 2"},
            {"nombre": "Tres", "cargo": "Cargo 3"},
            {"nombre": "Cuatro", "cargo": "Cargo 4"},
        ]

        row_insertions = presentacion_programa._build_section_5_row_insertions(
            "1. PRESENTACIÓN DEL PROGRAMA IL",
            payload,
        )

        self.assertEqual(
            row_insertions,
            [
                {
                    "sheet_name": "1. PRESENTACIÓN DEL PROGRAMA IL",
                    "start_row": 75,
                    "base_rows": 3,
                    "total_rows": 4,
                }
            ],
        )

    def test_build_section_5_row_insertions_skips_base_case(self) -> None:
        payload = [
            {"nombre": "Uno", "cargo": "Cargo 1"},
            {"nombre": "Dos", "cargo": "Cargo 2"},
            {"nombre": "Tres", "cargo": "Cargo 3"},
        ]

        row_insertions = presentacion_programa._build_section_5_row_insertions(
            "1. PRESENTACIÓN DEL PROGRAMA IL",
            payload,
        )

        self.assertEqual(row_insertions, [])


if __name__ == "__main__":
    unittest.main()
