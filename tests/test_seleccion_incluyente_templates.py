from __future__ import annotations

import unittest

from formularios.seleccion_incluyente import seleccion_incluyente as si


class SeleccionIncluyenteUnifiedTests(unittest.TestCase):
    def test_sheet_name_is_unified(self) -> None:
        self.assertEqual(si.SHEET_NAME, "4. SELECCIÓN INCLUYENTE")

    def test_oferente_cell_map_has_expected_keys(self) -> None:
        required_keys = {"numero", "nombre_oferente", "cedula", "discapacidad", "telefono_oferente"}
        self.assertTrue(required_keys.issubset(si.OFERENTE_CELL_MAP.keys()))

    def test_section_1_cell_map_has_expected_keys(self) -> None:
        required_keys = {"fecha_visita", "nombre_empresa", "modalidad"}
        self.assertTrue(required_keys.issubset(si.SECTION_1_CELL_MAP.keys()))

    def test_block_height_is_positive(self) -> None:
        self.assertGreater(si.OFERENTE_BLOCK_HEIGHT, 0)

    def test_section_6_row_insertions_use_section_constant(self) -> None:
        payload = [{"nombre": f"Asistente {idx}", "cargo": "Profesional"} for idx in range(5)]
        insertions = si._build_section_6_row_insertions(payload, num_oferentes=2)

        self.assertEqual(len(insertions), 1)
        self.assertEqual(insertions[0]["sheet_name"], si.SHEET_NAME)
        self.assertEqual(
            insertions[0]["start_row"],
            si.SECTION_6_BASE_START_ROW + si.OFERENTE_BLOCK_HEIGHT,
        )
        self.assertEqual(insertions[0]["base_rows"], si.SECTION_6["rows"])
        self.assertEqual(insertions[0]["total_rows"], 5)

    def test_selection_template_uses_two_base_attendee_rows(self) -> None:
        self.assertEqual(si.SECTION_6["rows"], 2)

    def test_section_2_group_row_insertions_clone_full_block(self) -> None:
        payload = [{"numero": str(idx + 1)} for idx in range(3)]
        insertions = si._build_section_2_row_insertions(payload)

        self.assertEqual(len(insertions), 1)
        self.assertEqual(insertions[0]["sheet_name"], si.SHEET_NAME)
        self.assertEqual(insertions[0]["insert_at_row"], si.OFERENTE_SECOND_BLOCK_START_ROW)
        self.assertEqual(insertions[0]["template_start_row"], si.OFERENTE_FIRST_BLOCK_START_ROW)
        self.assertEqual(
            insertions[0]["template_end_row"],
            si.OFERENTE_FIRST_BLOCK_START_ROW + si.OFERENTE_BLOCK_HEIGHT - 1,
        )
        self.assertEqual(insertions[0]["repeat_count"], 2)

    def test_section_2_group_writes_include_titles_and_export_title(self) -> None:
        payload = [
            {"numero": "1", "nombre_oferente": "Oferente 1"},
            {"numero": "2", "nombre_oferente": "Oferente 2"},
        ]
        writes = si._build_section_2_writes(payload)
        mapping = {item["range"]: item["value"] for item in writes}

        self.assertEqual(
            mapping[f"'{si.SHEET_NAME}'!{si.GROUP_EXPORT_TITLE_CELL}"],
            "PROCESO DE SELECCION INCLUYENTE GRUPAL - 2 A 4 OFERENTES",
        )
        self.assertEqual(
            mapping[f"'{si.SHEET_NAME}'!{si.OFERENTE_TITLE_COL}{si.OFERENTE_FIRST_BLOCK_START_ROW}"],
            "OFERENTE 1",
        )
        self.assertEqual(
            mapping[f"'{si.SHEET_NAME}'!{si.OFERENTE_TITLE_COL}{si.OFERENTE_SECOND_BLOCK_START_ROW}"],
            "OFERENTE 2",
        )
        self.assertEqual(
            mapping[f"'{si.SHEET_NAME}'!C19"],
            "Oferente 1",
        )
        self.assertEqual(
            mapping[f"'{si.SHEET_NAME}'!C80"],
            "Oferente 2",
        )

    def test_comunicacion_escrita_fields_use_expected_columns(self) -> None:
        self.assertEqual(si.OFERENTE_CELL_MAP["comunicacion_escrita_nivel_apoyo"], ("I", 51))
        self.assertEqual(si.OFERENTE_CELL_MAP["comunicacion_escrita_apoyo"], ("N", 51))
        self.assertEqual(si.OFERENTE_CELL_MAP["comunicacion_escrita_nota"], ("O", 52))

    def test_comunicacion_escrita_note_honors_oferente_offsets(self) -> None:
        payload = [
            {"numero": "1", "comunicacion_escrita_nota": "Nota oferente 1"},
            {"numero": "2", "comunicacion_escrita_nota": "Nota oferente 2"},
        ]

        writes = si._build_section_2_writes(payload)
        mapping = {item["range"]: item["value"] for item in writes}

        self.assertEqual(
            mapping[f"'{si.SHEET_NAME}'!O52"],
            "Nota oferente 1",
        )
        self.assertEqual(
            mapping[f"'{si.SHEET_NAME}'!O{52 + si.OFERENTE_BLOCK_HEIGHT}"],
            "Nota oferente 2",
        )

    def test_section_6_group_overflow_shifts_with_oferentes_and_keeps_simple_row_contract(self) -> None:
        payload = [{"nombre": f"Asistente {idx}", "cargo": "Profesional"} for idx in range(1, 7)]

        insertions = si._build_section_6_row_insertions(payload, num_oferentes=4)
        writes = si._build_section_6_writes(payload, num_oferentes=4)
        start_row = si.SECTION_6_BASE_START_ROW + (3 * si.OFERENTE_BLOCK_HEIGHT)

        self.assertEqual(len(insertions), 1)
        self.assertEqual(insertions[0]["sheet_name"], si.SHEET_NAME)
        self.assertEqual(insertions[0]["start_row"], start_row)
        self.assertEqual(insertions[0]["base_rows"], si.SECTION_6["rows"])
        self.assertEqual(insertions[0]["total_rows"], 6)
        self.assertNotIn("template_start_row", insertions[0])
        self.assertNotIn("repeat_count", insertions[0])
        self.assertEqual(
            writes[0]["range"],
            f"'{si.SHEET_NAME}'!{si.SECTION_6_NOMBRE_COL}{start_row}",
        )
        self.assertEqual(
            writes[-1]["range"],
            f"'{si.SHEET_NAME}'!{si.SECTION_6_CARGO_COL}{start_row + 5}",
        )


if __name__ == "__main__":
    unittest.main()
