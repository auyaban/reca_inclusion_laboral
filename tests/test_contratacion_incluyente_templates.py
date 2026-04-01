from __future__ import annotations

import unittest

from formularios.contratacion_incluyente import contratacion_incluyente as contratacion


class ContratacionIncluyenteUnifiedTests(unittest.TestCase):
    def test_sheet_name_is_unified(self) -> None:
        self.assertEqual(contratacion.SHEET_NAME, "5. CONTRATACIÓN INCLUYENTE")

    def test_vinculado_cell_map_has_expected_keys(self) -> None:
        required_keys = {"numero", "nombre_oferente", "cedula", "discapacidad", "telefono_oferente"}
        self.assertTrue(required_keys.issubset(contratacion.VINCULADO_CELL_MAP.keys()))

    def test_section_1_cell_map_has_expected_keys(self) -> None:
        required_keys = {"fecha_visita", "nombre_empresa", "modalidad"}
        self.assertTrue(required_keys.issubset(contratacion.SECTION_1_CELL_MAP.keys()))

    def test_block_height_is_positive(self) -> None:
        self.assertGreater(contratacion.VINCULADO_BLOCK_HEIGHT, 0)

    def test_normalize_dropdown_returns_known_option(self) -> None:
        result = contratacion.normalize_excel_dropdown_value("modalidad", "Presencial")
        self.assertEqual(result, "Presencial")

    def test_normalize_dropdown_returns_raw_for_unknown_field(self) -> None:
        result = contratacion.normalize_excel_dropdown_value("unknown_field", "test")
        self.assertEqual(result, "test")

    def test_contratacion_template_uses_four_base_attendee_rows(self) -> None:
        self.assertEqual(contratacion.EXCEL_MAPPING["section_7"]["rows"], 4)

    def test_section_2_group_row_insertions_clone_full_block(self) -> None:
        payload = [{"numero": str(idx + 1)} for idx in range(4)]
        insertions = contratacion._build_section_2_row_insertions(payload)

        self.assertEqual(len(insertions), 1)
        self.assertEqual(insertions[0]["sheet_name"], contratacion.SHEET_NAME)
        self.assertEqual(insertions[0]["insert_at_row"], contratacion.VINCULADO_SECOND_BLOCK_START_ROW)
        self.assertEqual(insertions[0]["template_start_row"], contratacion.VINCULADO_FIRST_BLOCK_START_ROW)
        self.assertEqual(
            insertions[0]["template_end_row"],
            contratacion.VINCULADO_FIRST_BLOCK_START_ROW + contratacion.VINCULADO_BLOCK_HEIGHT - 1,
        )
        self.assertEqual(insertions[0]["repeat_count"], 3)

    def test_section_2_group_writes_include_export_title_and_offsets(self) -> None:
        payload = [
            {"numero": "1", "nombre_oferente": "Vinculado 1"},
            {"numero": "2", "nombre_oferente": "Vinculado 2"},
        ]
        writes = contratacion._build_section_2_writes(payload)
        mapping = {item["range"]: item["value"] for item in writes}

        self.assertEqual(
            mapping[f"'{contratacion.SHEET_NAME}'!{contratacion.GROUP_EXPORT_TITLE_CELL}"],
            "PROCESO CONTRATACION INCLUYENTE GRUPAL - 2 A 4 VINCULADOS",
        )
        self.assertEqual(
            mapping[f"'{contratacion.SHEET_NAME}'!C23"],
            "Vinculado 1",
        )
        self.assertEqual(
            mapping[f"'{contratacion.SHEET_NAME}'!C75"],
            "Vinculado 2",
        )


if __name__ == "__main__":
    unittest.main()
