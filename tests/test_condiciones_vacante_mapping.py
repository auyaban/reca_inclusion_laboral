from __future__ import annotations

import unittest

from formularios.condiciones_vacante import condiciones_vacante as vacante


class CondicionesVacanteMappingTests(unittest.TestCase):
    def test_section_8_attendees_keep_expected_columns(self) -> None:
        mapping = vacante.EXCEL_MAPPING["section_8"]
        self.assertEqual(mapping["name_col"], "E")
        self.assertEqual(mapping["cargo_col"], "L")
        self.assertEqual(mapping["start_row"], 158)

    def test_section_7_and_8_shift_after_extra_disability_rows(self) -> None:
        original_cache = dict(vacante.FORM_CACHE)
        try:
            vacante.FORM_CACHE.clear()
            vacante.FORM_CACHE["section_6"] = [
                {"discapacidad": f"Discapacidad {index}", "descripcion": f"Descripcion {index}"}
                for index in range(7)
            ]

            section_7_writes = vacante._build_section_writes(
                "section_7",
                {"observaciones_recomendaciones": "Pendiente"},
            )
            section_8_writes = vacante._build_section_writes(
                "section_8",
                [{"nombre": "Leidy Novoa", "cargo": "Profesional"}],
            )

            self.assertEqual(
                section_7_writes[0]["range"],
                f"'{vacante.SHEET_NAME}'!A159",
            )
            self.assertEqual(
                section_8_writes[0]["range"],
                f"'{vacante.SHEET_NAME}'!E161",
            )
            self.assertEqual(
                section_8_writes[1]["range"],
                f"'{vacante.SHEET_NAME}'!L161",
            )
        finally:
            vacante.FORM_CACHE.clear()
            vacante.FORM_CACHE.update(original_cache)

    def test_row_insertions_shift_section_8_after_extra_disability_rows(self) -> None:
        cache = {
            "section_6": [
                {"discapacidad": f"Discapacidad {index}", "descripcion": f"Descripcion {index}"}
                for index in range(8)
            ],
            "section_8": [
                {"nombre": "Uno", "cargo": "Cargo 1"},
                {"nombre": "Dos", "cargo": "Cargo 2"},
                {"nombre": "Tres", "cargo": "Cargo 3"},
                {"nombre": "Cuatro", "cargo": "Cargo 4"},
                {"nombre": "Cinco", "cargo": "Cargo 5"},
            ],
        }

        row_insertions = vacante._build_row_insertions(cache)

        self.assertEqual(
            row_insertions,
            [
                {
                    "sheet_name": vacante.SHEET_NAME,
                    "start_row": 150,
                    "base_rows": 4,
                    "total_rows": 8,
                },
                {
                    "sheet_name": vacante.SHEET_NAME,
                    "start_row": 162,
                    "base_rows": 3,
                    "total_rows": 5,
                },
            ],
        )


if __name__ == "__main__":
    unittest.main()
