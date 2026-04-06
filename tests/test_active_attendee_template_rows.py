from __future__ import annotations

import unittest

from formularios.condiciones_vacante import condiciones_vacante
from formularios.contratacion_incluyente import contratacion_incluyente
from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.induccion_operativa import induccion_operativa
from formularios.induccion_organizacional import induccion_organizacional
from formularios.presentacion_programa import presentacion_programa
from formularios.seleccion_incluyente import seleccion_incluyente
from formularios.sensibilizacion import sensibilizacion


class ActiveAttendeeTemplateRowsTests(unittest.TestCase):
    def test_active_forms_match_master_template_attendee_base_rows(self) -> None:
        self.assertEqual(sensibilizacion.SECTION_5["rows"], 4)
        self.assertEqual(len(presentacion_programa.SECTION_5["items"]), 3)
        self.assertEqual(evaluacion_accesibilidad.EXCEL_MAPPING["section_8"]["base_rows"], 4)
        self.assertEqual(condiciones_vacante.EXCEL_MAPPING["section_8"]["rows"], 3)
        self.assertEqual(seleccion_incluyente.SECTION_6["rows"], 2)
        self.assertEqual(contratacion_incluyente.EXCEL_MAPPING["section_7"]["rows"], 4)
        self.assertEqual(induccion_organizacional.SECTION_6_BASE_ROWS, 4)
        self.assertEqual(induccion_operativa.SECTION_9["rows"], 4)

    def test_induccion_organizacional_insertions_use_four_base_rows(self) -> None:
        # Template maestro tiene 4 filas base (filas 71-74); inserción solo ocurre a partir del 5.º asistente
        payload_4 = [{"nombre": f"Asistente {i}", "cargo": "Profesional"} for i in range(4)]
        self.assertEqual(induccion_organizacional._build_section_6_row_insertions(payload_4), [])

        payload_5 = [{"nombre": f"Asistente {i}", "cargo": "Profesional"} for i in range(5)]
        insertions = induccion_organizacional._build_section_6_row_insertions(payload_5)
        self.assertEqual(len(insertions), 1)
        self.assertEqual(insertions[0]["sheet_name"], induccion_organizacional.SHEET_NAME)
        self.assertEqual(insertions[0]["start_row"], induccion_organizacional.SECTION_6_START_ROW)
        self.assertEqual(insertions[0]["base_rows"], 4)
        self.assertEqual(insertions[0]["total_rows"], 5)


if __name__ == "__main__":
    unittest.main()
