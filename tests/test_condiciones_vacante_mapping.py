from __future__ import annotations

import unittest

from formularios.condiciones_vacante import condiciones_vacante as vacante


class CondicionesVacanteMappingTests(unittest.TestCase):
    def test_section_8_attendees_keep_expected_columns(self) -> None:
        mapping = vacante.EXCEL_MAPPING["section_8"]
        self.assertEqual(mapping["name_col"], "E")
        self.assertEqual(mapping["cargo_col"], "L")
        self.assertEqual(mapping["start_row"], 161)

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
                f"'{vacante.SHEET_NAME}'!A162",
            )
            self.assertEqual(
                section_8_writes[0]["range"],
                f"'{vacante.SHEET_NAME}'!E164",
            )
            self.assertEqual(
                section_8_writes[1]["range"],
                f"'{vacante.SHEET_NAME}'!L164",
            )
        finally:
            vacante.FORM_CACHE.clear()
            vacante.FORM_CACHE.update(original_cache)


if __name__ == "__main__":
    unittest.main()
