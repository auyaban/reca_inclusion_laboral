from __future__ import annotations

import unittest

import completion_payloads


class CompletionPayloadTests(unittest.TestCase):
    def test_builds_program_presentation_payload(self) -> None:
        result = completion_payloads.build_completion_payload(
            "presentacion_programa",
            "Presentacion/Reactivacion del programa de inclusion laboral",
            {
                "section_1": {
                    "tipo_visita": "Presentación",
                    "fecha_visita": "2026-03-17",
                    "nombre_empresa": "Empresa Demo",
                    "nit_empresa": "900123456",
                    "modalidad": "Virtual",
                    "profesional_asignado": "Leidy Novoa",
                    "ciudad_empresa": "Bogota",
                    "sede_empresa": "Centro",
                    "caja_compensacion": "Compensar",
                    "asesor": "Asesor Demo",
                },
                "section_5": [{"nombre": "Leidy Novoa", "cargo": "Profesional"}],
            },
            output_path=r"C:\tmp\presentacion.xlsx",
            session_id="session-1",
            app_version="2.0.0",
        )

        payload = result["payload_normalized"]
        self.assertEqual(payload["attachment"]["document_kind"], "program_presentation")
        self.assertEqual(payload["parsed_raw"]["nombre_profesional"], "Leidy Novoa")
        self.assertEqual(payload["parsed_raw"]["asistentes"], ["Leidy Novoa"])
        self.assertEqual(result["source_item_key"], "presentacion_programa:session-1")

    def test_builds_vacancy_review_payload(self) -> None:
        result = completion_payloads.build_completion_payload(
            "condiciones_vacante",
            "Condiciones de Vacante",
            {
                "section_1": {
                    "fecha_visita": "2026-03-17",
                    "nombre_empresa": "Empresa Demo",
                    "nit_empresa": "900123456",
                    "modalidad": "Presencial",
                    "profesional_asignado": "Leidy Novoa",
                },
                "section_2": {
                    "nombre_vacante": "Operario logistico",
                    "numero_vacantes": "3",
                },
                "section_8": [{"nombre": "Leidy Novoa", "cargo": "Profesional"}],
            },
            output_path=r"C:\tmp\vacante.xlsx",
            session_id="session-2",
            app_version="2.0.0",
        )

        parsed = result["payload_normalized"]["parsed_raw"]
        self.assertEqual(result["payload_normalized"]["attachment"]["document_kind"], "vacancy_review")
        self.assertEqual(parsed["cargo_objetivo"], "Operario logistico")
        self.assertEqual(parsed["total_vacantes"], "3")

    def test_builds_vacancy_labs_payload_with_same_normalizer(self) -> None:
        result = completion_payloads.build_completion_payload(
            "condiciones_vacante_labs",
            "Condiciones de Vacante Labs",
            {
                "section_1": {
                    "fecha_visita": "2026-03-17",
                    "nombre_empresa": "Empresa Demo",
                    "nit_empresa": "900123456",
                    "modalidad": "Presencial",
                    "profesional_asignado": "Leidy Novoa",
                },
                "section_2": {
                    "nombre_vacante": "Operario logistico",
                    "numero_vacantes": "2",
                },
                "section_8": [{"nombre": "Leidy Novoa", "cargo": "Profesional"}],
            },
            output_path=r"C:\tmp\vacante_labs.xlsx",
            session_id="session-vac-labs",
            app_version="2.0.0",
        )

        payload = result["payload_normalized"]
        self.assertEqual(payload["form_id"], "condiciones_vacante_labs")
        self.assertEqual(payload["attachment"]["document_kind"], "vacancy_review")
        self.assertEqual(payload["parsed_raw"]["cargo_objetivo"], "Operario logistico")
        self.assertEqual(payload["parsed_raw"]["total_vacantes"], "2")
        self.assertEqual(result["source_item_key"], "condiciones_vacante_labs:session-vac-labs")

    def test_builds_selection_payload_with_participants_and_unique_cargo(self) -> None:
        result = completion_payloads.build_completion_payload(
            "seleccion_incluyente",
            "Proceso de Seleccion Incluyente",
            {
                "section_1": {
                    "fecha_visita": "2026-03-17",
                    "nombre_empresa": "Empresa Demo",
                    "nit_empresa": "900123456",
                    "modalidad": "Virtual",
                    "profesional_asignado": "Leidy Novoa",
                },
                "section_2": [
                    {
                        "nombre_oferente": "Ana Perez",
                        "cedula": "123",
                        "discapacidad": "Auditiva",
                        "genero": "Femenino",
                        "cargo_oferente": "Auxiliar",
                    },
                    {
                        "nombre_oferente": "Luis Gomez",
                        "cedula": "456",
                        "discapacidad": "Visual",
                        "genero": "Masculino",
                        "cargo_oferente": "Auxiliar",
                    },
                ],
                "section_6": [{"nombre": "Leidy Novoa", "cargo": "Profesional"}],
            },
            output_path=r"C:\tmp\seleccion.xlsx",
            session_id="session-3",
            app_version="2.0.0",
        )

        parsed = result["payload_normalized"]["parsed_raw"]
        self.assertEqual(result["payload_normalized"]["attachment"]["document_kind"], "inclusive_selection")
        self.assertEqual(parsed["cargo_objetivo"], "Auxiliar")
        self.assertEqual(len(parsed["participantes"]), 2)
        self.assertEqual(parsed["participantes"][0]["cedula_usuario"], "123")

    def test_builds_selection_labs_payload_with_same_normalizer(self) -> None:
        result = completion_payloads.build_completion_payload(
            "seleccion_incluyente_labs",
            "Proceso de Seleccion Incluyente Labs",
            {
                "section_1": {
                    "fecha_visita": "2026-03-17",
                    "nombre_empresa": "Empresa Demo",
                    "nit_empresa": "900123456",
                    "modalidad": "Virtual",
                    "profesional_asignado": "Leidy Novoa",
                },
                "section_2": [
                    {
                        "nombre_oferente": "Ana Perez",
                        "cedula": "123",
                        "discapacidad": "Auditiva",
                        "cargo_oferente": "Auxiliar",
                    }
                ],
                "section_6": [{"nombre": "Leidy Novoa", "cargo": "Profesional"}],
            },
            output_path=r"C:\tmp\seleccion_labs.xlsx",
            session_id="session-labs",
            app_version="2.0.0",
        )

        payload = result["payload_normalized"]
        self.assertEqual(payload["form_id"], "seleccion_incluyente_labs")
        self.assertEqual(payload["attachment"]["document_kind"], "inclusive_selection")
        self.assertEqual(payload["parsed_raw"]["cargo_objetivo"], "Auxiliar")
        self.assertEqual(result["source_item_key"], "seleccion_incluyente_labs:session-labs")

    def test_builds_sensibilizacion_payload_without_professional_field(self) -> None:
        result = completion_payloads.build_completion_payload(
            "sensibilizacion",
            "Sensibilizacion",
            {
                "section_1": {
                    "fecha_visita": "2026-03-17",
                    "nombre_empresa": "Empresa Demo",
                    "nit_empresa": "900123456",
                    "modalidad": "Mixta",
                },
                "section_5": [{"nombre": "Leidy Novoa", "cargo": "Profesional"}],
            },
            output_path=r"C:\tmp\sensibilizacion.xlsx",
            session_id="session-4",
            app_version="2.0.0",
        )

        parsed = result["payload_normalized"]["parsed_raw"]
        self.assertEqual(result["payload_normalized"]["attachment"]["document_kind"], "sensibilizacion")
        self.assertEqual(parsed["nombre_profesional"], "")
        self.assertEqual(parsed["candidatos_profesional"], ["Leidy Novoa"])

    def test_builds_followup_payload(self) -> None:
        result = completion_payloads.build_followup_completion_payload(
            case_ref=r"C:\tmp\seguimiento.xlsx",
            followup_index=2,
            base_payload={
                "fecha_visita": "2026-03-17",
                "nombre_empresa": "Empresa Demo",
                "nit_empresa": "900123456",
                "modalidad": "Presencial",
                "profesional_asignado": "Leidy Novoa",
                "nombre_vinculado": "Juan Perez",
                "cedula": "999",
                "discapacidad": "Fisica",
                "cargo_vinculado": "Auxiliar logistico",
            },
            followup_payload={
                "seguimiento_numero": "2",
                "tipo_apoyo": "Seguimiento",
                "situacion_encontrada": "Sin novedades",
                "estrategias_ajustes": "Acompanamiento",
                "asistentes": [{"nombre": "Leidy Novoa", "cargo": "Profesional"}],
            },
            form_name="Seguimiento al Proceso de Inclusion Laboral #2",
            session_id="session-5",
            app_version="2.0.0",
            extra_context={
                "local_path": r"C:\tmp\seguimiento.xlsx",
                "remote_url": "https://drive.google.com/demo",
                "drive_file_id": "drive-123",
                "case_meta": {"cedula": "999"},
            },
        )

        payload = result["payload_normalized"]
        self.assertEqual(payload["attachment"]["document_kind"], "follow_up")
        self.assertEqual(payload["parsed_raw"]["numero_seguimiento"], "2")
        self.assertEqual(payload["parsed_raw"]["participantes"][0]["cedula_usuario"], "999")
        self.assertIn("seguimientos:", result["source_item_key"])
        self.assertTrue(result["source_item_key"].endswith(":followup:2"))

    def test_strips_internal_cache_metadata_from_raw_payload(self) -> None:
        result = completion_payloads.build_completion_payload(
            "induccion_organizacional",
            "Induccion Organizacional",
            {
                "section_1": {
                    "fecha_visita": "2026-03-17",
                    "nombre_empresa": "Empresa Demo",
                    "nit_empresa": "900123456",
                    "modalidad": "Presencial",
                    "profesional_asignado": "Leidy Novoa",
                },
                "section_2": [{"nombre_oferente": "Ana", "cargo_oferente": "Auxiliar"}],
                "section_6": [{"nombre": "Leidy Novoa", "cargo": "Profesional"}],
                "_section_history": {"section_3": [{"payload": {"foo": "bar"}}]},
                "_last_saved_at": "2026-03-27 10:00:00",
                "_last_saved_section": "section_3",
                "_last_saved_source": "autosave",
            },
            output_path=r"C:\tmp\induccion.xlsx",
            session_id="session-6",
            app_version="2.0.0",
        )

        raw_cache = result["payload_raw"]["cache_snapshot"]
        self.assertNotIn("_section_history", raw_cache)
        self.assertNotIn("_last_saved_at", raw_cache)
        self.assertNotIn("_last_saved_section", raw_cache)
        self.assertNotIn("_last_saved_source", raw_cache)


if __name__ == "__main__":
    unittest.main()
