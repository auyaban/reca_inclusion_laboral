import unittest

from formularios.condiciones_vacante.voice_section2 import (
    FIELD_ORDER_BY_SUBSECTION,
    get_subsection_spec,
    postprocess_extraction_payload,
    summarize_candidate_updates,
)


def _build_payload(subsection_key: str):
    if subsection_key == "section_2_vacancy":
        section_id = "section_2"
        semantic = {
            "section_2_vacancy": {
                "vacancy": {
                    "vacancy_name": None,
                    "openings_count": None,
                    "position_level": None,
                    "gender_preference": None,
                    "age_requirement_text": None,
                    "work_modality_text": None,
                    "work_location_text": None,
                    "salary_text": None,
                    "contract_signing_text": None,
                    "tests_text": None,
                    "contract_type": None,
                    "additional_benefits_text": None,
                    "gender_flexibility_text": None,
                    "women_benefits_text": None,
                    "certificate_requirement": None,
                    "certificate_notes": None,
                }
            },
            "section_2_1_schedule_experience": {
                "schedule_experience": {
                    "schedule_type": None,
                    "entry_time_text": None,
                    "exit_time_text": None,
                    "lunch_duration": None,
                    "break_duration": None,
                    "workdays_text": None,
                    "flexible_days_text": None,
                    "schedule_notes_text": None,
                    "experience_requirement": None,
                    "main_functions_text": None,
                    "tools_and_equipment_text": None,
                }
            },
        }
    else:
        section_id = "section_2_1"
        semantic = {
            "section_2_vacancy": {
                "vacancy": {
                    "vacancy_name": None,
                    "openings_count": None,
                    "position_level": None,
                    "gender_preference": None,
                    "age_requirement_text": None,
                    "work_modality_text": None,
                    "work_location_text": None,
                    "salary_text": None,
                    "contract_signing_text": None,
                    "tests_text": None,
                    "contract_type": None,
                    "additional_benefits_text": None,
                    "gender_flexibility_text": None,
                    "women_benefits_text": None,
                    "certificate_requirement": None,
                    "certificate_notes": None,
                }
            },
            "section_2_1_schedule_experience": {
                "schedule_experience": {
                    "schedule_type": None,
                    "entry_time_text": None,
                    "exit_time_text": None,
                    "lunch_duration": None,
                    "break_duration": None,
                    "workdays_text": None,
                    "flexible_days_text": None,
                    "schedule_notes_text": None,
                    "experience_requirement": None,
                    "main_functions_text": None,
                    "tools_and_equipment_text": None,
                }
            },
        }
    return {
        "schema_version": 1,
        "form_id": "condiciones_vacante",
        "section_id": section_id,
        "subsection_key": subsection_key,
        "audio_unit": "single_section",
        "transcription_summary": "Resumen",
        "warnings": [],
        "semantic": semantic,
        "candidate": {field_id: None for field_id in FIELD_ORDER_BY_SUBSECTION[subsection_key]},
    }


class CondicionesVacanteVoiceTests(unittest.TestCase):
    def test_semantic_dropdowns_map_to_exact_options(self):
        payload = _build_payload("section_2_vacancy")
        payload["semantic"]["section_2_vacancy"]["vacancy"].update(
            {
                "position_level": "operational",
                "gender_preference": "indifferent",
                "contract_type": "fixed_term",
                "certificate_requirement": "in_process",
                "openings_count": "2",
            }
        )
        payload["candidate"].update(
            {
                "nombre_vacante": "Auxiliar logistico",
                "salario_asignado": "Salario minimo con prestaciones",
                "firma_contrato": "Contratacion inmediata",
            }
        )
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_2_vacancy",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["nivel_cargo"], "Operativo.")
        self.assertEqual(processed["candidate"]["genero"], "Indiferente")
        self.assertEqual(processed["candidate"]["tipo_contrato"], "Término Fijo.")
        self.assertEqual(processed["candidate"]["requiere_certificado"], "En Trámite")
        self.assertEqual(processed["candidate"]["numero_vacantes"], "2")

    def test_numero_vacantes_text_is_normalized_to_digits(self):
        payload = _build_payload("section_2_vacancy")
        payload["candidate"]["numero_vacantes"] = "dos vacantes"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_2_vacancy",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["numero_vacantes"], "2")

    def test_candidate_fallback_normalizes_dropdown_aliases(self):
        payload = _build_payload("section_2_vacancy")
        payload["candidate"].update(
            {
                "nivel_cargo": "operativo",
                "genero": "sin preferencia",
                "tipo_contrato": "termino indefinido",
                "requiere_certificado": "si",
            }
        )
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_2_vacancy",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["nivel_cargo"], "Operativo.")
        self.assertEqual(processed["candidate"]["genero"], "Indiferente")
        self.assertEqual(processed["candidate"]["tipo_contrato"], "Término Indefinido.")
        self.assertEqual(processed["candidate"]["requiere_certificado"], "Sí")

    def test_invalid_dropdown_value_is_dropped_with_warning(self):
        payload = _build_payload("section_2_vacancy")
        payload["candidate"]["nivel_cargo"] = "super senior"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_2_vacancy",
            candidate_index=1,
        )
        self.assertNotIn("nivel_cargo", processed["candidate"])
        self.assertTrue(any("Nivel del cargo" in item for item in processed["warnings"]))

    def test_section_2_1_semantic_values_map_correctly(self):
        payload = _build_payload("section_2_1_schedule_experience")
        payload["semantic"]["section_2_1_schedule_experience"]["schedule_experience"].update(
            {
                "schedule_type": "rotating",
                "lunch_duration": "1h",
                "break_duration": "15m",
                "experience_requirement": "one_year",
                "entry_time_text": "6:00 a.m.",
                "exit_time_text": "2:00 p.m.",
                "main_functions_text": "Operar maquinas y reportar novedades.",
            }
        )
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_2_1_schedule_experience",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["horarios_asignados"], "Horarios Rotativos.")
        self.assertEqual(processed["candidate"]["tiempo_almuerzo"], "1 hora.")
        self.assertEqual(processed["candidate"]["break_descanso"], "15 minutos")
        self.assertEqual(processed["candidate"]["experiencia_meses"], "Un año.")
        self.assertEqual(processed["candidate"]["hora_ingreso"], "6:00 a.m.")
        self.assertEqual(
            processed["candidate"]["funciones_tareas"],
            "Operar maquinas y reportar novedades.",
        )

    def test_spec_and_summary_are_human_readable(self):
        spec = get_subsection_spec("section_2_vacancy")
        self.assertIn("lenguaje natural", str(spec.get("script") or "").lower())
        self.assertEqual(len(spec.get("examples") or []), 1)
        spec_21 = get_subsection_spec("section_2_1_schedule_experience")
        self.assertIn("horario", " ".join(spec_21.get("questions") or []).lower())
        summary = summarize_candidate_updates(
            {
                "nombre_vacante": "Auxiliar de produccion",
                "salario_asignado": "$1.423.500 con prestaciones",
            },
            subsection_key="section_2_vacancy",
        )
        self.assertEqual(
            summary,
            [
                "Nombre de la vacante: Auxiliar de produccion",
                "Salario asignado: $1.423.500 con prestaciones",
            ],
        )


if __name__ == "__main__":
    unittest.main()
