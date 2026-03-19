import unittest
from datetime import date

from formularios.seleccion_incluyente_labs.voice_section2 import (
    AUDIO_UNIT,
    FORM_ID,
    SECTION_ID,
    build_empty_extraction_payload,
    derive_age_from_birthdate,
    get_subsection_spec,
    load_subsection_specs,
    merge_non_null_fields,
    postprocess_extraction_payload,
    validate_extraction_payload,
)
from logging_utils import get_log_file_path


class SelectionSection2VoiceTests(unittest.TestCase):
    def test_valid_empty_payload_passes_validation(self):
        payload = build_empty_extraction_payload("section_4_1_salud")
        self.assertEqual(validate_extraction_payload(payload), [])

    def test_invalid_dropdown_value_fails_validation(self):
        payload = build_empty_extraction_payload("section_2_fields")
        payload["candidate"]["resultado_certificado"] = 123
        errors = validate_extraction_payload(payload)
        self.assertTrue(any("resultado_certificado" in error for error in errors))

    def test_postprocess_derives_age_and_sets_tipo_pension_no_aplica(self):
        payload = build_empty_extraction_payload("section_2_fields")
        payload["transcription_summary"] = "Resumen"
        payload["candidate"]["nombre_oferente"] = "Juan Perez"
        payload["candidate"]["fecha_nacimiento"] = "12051998"
        payload["candidate"]["cuenta_pension"] = "No"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_2_fields",
            candidate_index=3,
        )
        expected_age = derive_age_from_birthdate("12/05/1998")
        self.assertEqual(processed["candidate"]["numero"], "3")
        self.assertEqual(processed["candidate"]["fecha_nacimiento"], "12/05/1998")
        self.assertEqual(processed["candidate"]["edad"], expected_age)
        self.assertEqual(processed["candidate"]["tipo_pension"], "No aplica")

    def test_section_2_semantic_payload_derives_birthdate_and_age(self):
        payload = build_empty_extraction_payload("section_2_fields")
        payload["semantic"]["section_2_identity"] = {
            "identity": {
                "document_number": "10203040",
                "applicant_phone": "3001234567",
                "emergency_phone": "3105556677",
                "birthdate_iso": "1998-05-12",
            }
        }
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_2_fields",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["cedula"], "10203040")
        self.assertEqual(processed["candidate"]["telefono_oferente"], "3001234567")
        self.assertEqual(processed["candidate"]["telefono_emergencia"], "3105556677")
        self.assertEqual(processed["candidate"]["fecha_nacimiento"], "12/05/1998")
        self.assertEqual(processed["candidate"]["edad"], derive_age_from_birthdate("12/05/1998"))

    def test_postprocess_keeps_tipo_pension_empty_when_por_confirmar(self):
        payload = build_empty_extraction_payload("section_2_fields")
        payload["candidate"]["cuenta_pension"] = "Por Confirmar"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_2_fields",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["numero"], "1")
        self.assertEqual(processed["candidate"]["cuenta_pension"], "Por Confirmar")
        self.assertNotIn("tipo_pension", processed["candidate"])

    def test_non_target_subsection_fields_are_ignored(self):
        payload = build_empty_extraction_payload("section_4_1_salud")
        payload["candidate"]["nombre_oferente"] = "No deberia pasar"
        payload["candidate"]["medicamentos_nivel_apoyo"] = "0. No requiere apoyo."
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_4_1_salud",
            candidate_index=2,
        )
        self.assertEqual(processed["candidate"]["numero"], "2")
        self.assertEqual(
            processed["candidate"]["medicamentos_nivel_apoyo"],
            "0. No requiere apoyo.",
        )
        self.assertNotIn("nombre_oferente", processed["candidate"])

    def test_merge_non_null_fields_does_not_wipe_existing_values(self):
        merged = merge_non_null_fields(
            {"nombre_oferente": "Valor manual", "telefono_oferente": "3001112233"},
            {"nombre_oferente": None, "telefono_oferente": "", "cargo_oferente": "Operario"},
        )
        self.assertEqual(merged["nombre_oferente"], "Valor manual")
        self.assertEqual(merged["telefono_oferente"], "3001112233")
        self.assertEqual(merged["cargo_oferente"], "Operario")

    def test_specs_include_expected_guidance_phrases(self):
        specs = load_subsection_specs()
        self.assertIn("warning", specs)
        block_b = get_subsection_spec("section_4_2_b_actividades")
        examples = " ".join(block_b.get("examples") or [])
        questions = " ".join(block_b.get("questions") or [])
        self.assertIn("lo demas no", examples.lower())
        self.assertIn("solo", examples.lower())
        self.assertIn("requiere apoyo", questions.lower())
        self.assertIn("grupo por grupo", str(block_b.get("script") or "").lower())
        block_fields = get_subsection_spec("section_2_fields")
        prompt_fragment = str(block_fields.get("prompt_fragment") or "").lower()
        self.assertIn("tipo_pension", prompt_fragment)
        self.assertIn("por confirmar", " ".join(block_fields.get("examples") or []).lower())
        self.assertGreaterEqual(len(block_fields.get("questions") or []), 3)
        block_3 = get_subsection_spec("section_3_desarrollo")
        self.assertIn("orden narrativo", str(block_3.get("prompt_fragment") or "").lower())
        block_4a = get_subsection_spec("section_4_2_a_habilidades")
        self.assertIn("nivel_apoyo", str(block_4a.get("prompt_fragment") or "").lower())
        self.assertIn("desplaza", " ".join(block_4a.get("questions") or []).lower())

    def test_core_constants_match_contract(self):
        payload = build_empty_extraction_payload("section_3_desarrollo")
        self.assertEqual(payload["form_id"], FORM_ID)
        self.assertEqual(payload["section_id"], SECTION_ID)
        self.assertEqual(payload["audio_unit"], AUDIO_UNIT)

    def test_labs_logger_path_exists(self):
        path = get_log_file_path("labs")
        self.assertTrue(path.endswith("labs.log"))

    def test_dropdown_aliases_are_normalized(self):
        payload = build_empty_extraction_payload("section_2_fields")
        payload["candidate"]["discapacidad"] = "Discapacidad fisica"
        payload["candidate"]["resultado_certificado"] = "aprobado"
        payload["candidate"]["pendiente_otros_oferentes"] = "por confirmar"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_2_fields",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["discapacidad"], "Discapacidad física")
        self.assertEqual(processed["candidate"]["resultado_certificado"], "Aprobado")
        self.assertEqual(processed["candidate"]["pendiente_otros_oferentes"], "Por Confirmar")

    def test_invalid_birthdate_and_phone_are_dropped_with_warning(self):
        payload = build_empty_extraction_payload("section_2_fields")
        payload["candidate"]["fecha_nacimiento"] = "21/99/6"
        payload["candidate"]["telefono_emergencia"] = "338-49-8830-22"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_2_fields",
            candidate_index=1,
        )
        self.assertNotIn("fecha_nacimiento", processed["candidate"])
        self.assertNotIn("telefono_emergencia", processed["candidate"])
        warnings = " ".join(processed["warnings"]).lower()
        self.assertIn("fecha de nacimiento", warnings)
        self.assertIn("telefono de emergencia", warnings)

    def test_section_2_long_document_and_short_phone_are_dropped(self):
        payload = build_empty_extraction_payload("section_2_fields")
        payload["semantic"]["section_2_identity"] = {
            "identity": {
                "document_number": "12345678901",
                "applicant_phone": "300123456",
                "emergency_phone": "310555667788",
                "birthdate_iso": "1998-05-12",
            }
        }
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_2_fields",
            candidate_index=1,
        )
        self.assertNotIn("cedula", processed["candidate"])
        self.assertNotIn("telefono_oferente", processed["candidate"])
        self.assertNotIn("telefono_emergencia", processed["candidate"])
        self.assertEqual(processed["candidate"]["fecha_nacimiento"], "12/05/1998")
        warnings = " ".join(processed["warnings"]).lower()
        self.assertIn("cedula", warnings)
        self.assertIn("telefono del oferente", warnings)
        self.assertIn("telefono de emergencia", warnings)

    def test_section_2_textual_birthdate_is_normalized(self):
        payload = build_empty_extraction_payload("section_2_fields")
        payload["candidate"]["fecha_nacimiento"] = "12 de mayo de 1998"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_2_fields",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["fecha_nacimiento"], "12/05/1998")
        self.assertEqual(processed["candidate"]["edad"], derive_age_from_birthdate("12/05/1998"))

    def test_contract_date_can_be_por_confirmar(self):
        payload = build_empty_extraction_payload("section_2_fields")
        payload["candidate"]["fecha_firma_contrato"] = "pendiente por confirmar"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_2_fields",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["fecha_firma_contrato"], "Por Confirmar")

    def test_section_4_1_natural_language_variants_are_normalized(self):
        payload = build_empty_extraction_payload("section_4_1_salud")
        payload["candidate"]["medicamentos_conocimiento"] = "un tercero conoce los medicamentos"
        payload["candidate"]["medicamentos_horarios"] = "no conoce los horarios de los medicamentos"
        payload["candidate"]["alergias_tipo"] = "presenta alergias y sabe darles manejo"
        payload["candidate"]["controles_asistencia"] = "asiste a controles medicos y conoce el manejo"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_4_1_salud",
            candidate_index=1,
        )
        self.assertEqual(
            processed["candidate"]["medicamentos_conocimiento"],
            "2. Un tercero es quien conoce los medicamentos que consume.",
        )
        self.assertEqual(
            processed["candidate"]["medicamentos_horarios"],
            "3. No conoce los horarios de toma de medicamentos que consume.",
        )
        self.assertEqual(
            processed["candidate"]["alergias_tipo"],
            "1. Presenta alergias y sabe darle manejo.",
        )
        self.assertEqual(
            processed["candidate"]["controles_asistencia"],
            "1. Asiste a controles medicos con especialista y conoce el manejo.",
        )

    def test_section_4_1_realistic_phrases_map_without_warnings(self):
        payload = build_empty_extraction_payload("section_4_1_salud")
        payload["candidate"]["medicamentos_nivel_apoyo"] = "no requiere apoyo"
        payload["candidate"]["medicamentos_conocimiento"] = "la persona conoce los medicamentos que tiene que tomar"
        payload["candidate"]["medicamentos_horarios"] = "conoce los horarios"
        payload["candidate"]["alergias_nivel_apoyo"] = "no requiere apoyo"
        payload["candidate"]["alergias_tipo"] = "no tiene alergias, no refiere alergias"
        payload["candidate"]["restriccion_nivel_apoyo"] = "nivel bajo"
        payload["candidate"]["restriccion_conocimiento"] = "si tiene restricciones medicas y la persona conoce el manejo de esas restricciones"
        payload["candidate"]["controles_nivel_apoyo"] = "nivel bajo"
        payload["candidate"]["controles_asistencia"] = "asiste a controles medicos con especialista y conoce el manejo de esto"
        payload["candidate"]["controles_frecuencia"] = "cada tres meses"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_4_1_salud",
            candidate_index=1,
        )
        self.assertEqual(
            processed["candidate"]["medicamentos_conocimiento"],
            "1. Conoce los medicamentos que consume.",
        )
        self.assertEqual(
            processed["candidate"]["medicamentos_horarios"],
            "1. Conoce los horarios de toma de medicamentos que consume.",
        )
        self.assertEqual(
            processed["candidate"]["alergias_tipo"],
            "0. No presenta alergias.",
        )
        self.assertEqual(
            processed["candidate"]["restriccion_conocimiento"],
            "1. Tiene restricciones medicas y conoce su manejo.",
        )
        self.assertEqual(
            processed["candidate"]["controles_asistencia"],
            "1. Asiste a controles medicos con especialista y conoce el manejo.",
        )
        self.assertEqual(processed["candidate"]["controles_frecuencia"], "Trimestral")
        self.assertEqual(processed["warnings"], [])

    def test_section_4_1_semantic_payload_maps_to_dropdowns(self):
        payload = build_empty_extraction_payload("section_4_1_salud")
        payload["semantic"]["section_4_1_health"] = {
            "medications": {
                "support_level": "not applicable",
                "status": "not taking",
                "schedule_status": "not applicable",
                "details": "No toma medicamentos actualmente.",
            },
            "allergies": {
                "support_level": "none",
                "status": "none reported",
                "details": "No refiere alergias.",
            },
            "restrictions": {
                "support_level": "low",
                "status": "self managed",
                "details": "No puede cargar peso y sabe manejarlo.",
            },
            "specialist_controls": {
                "support_level": "low",
                "attendance": "attends",
                "frequency": "monthly",
                "details": "Consulta con ortopedia una vez al mes.",
            },
        }
        self.assertEqual(validate_extraction_payload(payload), [])
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_4_1_salud",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["medicamentos_nivel_apoyo"], "No aplica.")
        self.assertEqual(processed["candidate"]["medicamentos_conocimiento"], "No aplica.")
        self.assertEqual(processed["candidate"]["medicamentos_horarios"], "No aplica.")
        self.assertEqual(processed["candidate"]["alergias_nivel_apoyo"], "0. No requiere apoyo.")
        self.assertEqual(processed["candidate"]["alergias_tipo"], "0. No presenta alergias.")
        self.assertEqual(
            processed["candidate"]["restriccion_conocimiento"],
            "1. Tiene restricciones medicas y conoce su manejo.",
        )
        self.assertEqual(
            processed["candidate"]["controles_asistencia"],
            "2. Si asiste a controles medicos con especialista.",
        )
        self.assertEqual(processed["candidate"]["controles_frecuencia"], "Mensual")
        self.assertEqual(
            processed["candidate"]["controles_nota"],
            "Consulta con ortopedia una vez al mes.",
        )
        self.assertIn("semantic", processed)

    def test_section_4_1_defaults_do_not_emit_warnings_for_no_aplica_case(self):
        payload = build_empty_extraction_payload("section_4_1_salud")
        payload["candidate"]["medicamentos_nivel_apoyo"] = "no aplica"
        payload["candidate"]["medicamentos_conocimiento"] = "no toma ningun medicamento"
        payload["candidate"]["medicamentos_horarios"] = "no toma ningun medicamento"
        payload["candidate"]["alergias_nivel_apoyo"] = "no aplica"
        payload["candidate"]["alergias_tipo"] = "tampoco tiene alergias"
        payload["candidate"]["restriccion_nivel_apoyo"] = "nivel bajo"
        payload["candidate"]["restriccion_conocimiento"] = "el conoce que es lo que necesita y como puede manejarlo"
        payload["candidate"]["controles_nivel_apoyo"] = "nivel bajo"
        payload["candidate"]["controles_asistencia"] = "tiene cita con medico ortopedista una vez cada mes, el sabe lo que tiene que hacer"
        payload["candidate"]["controles_frecuencia"] = "una vez cada mes"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_4_1_salud",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["medicamentos_conocimiento"], "No aplica.")
        self.assertEqual(processed["candidate"]["medicamentos_horarios"], "No aplica.")
        self.assertEqual(processed["candidate"]["alergias_tipo"], "0. No presenta alergias.")
        self.assertEqual(
            processed["candidate"]["restriccion_conocimiento"],
            "1. Tiene restricciones medicas y conoce su manejo.",
        )
        self.assertEqual(
            processed["candidate"]["controles_asistencia"],
            "2. Si asiste a controles medicos con especialista.",
        )
        self.assertEqual(processed["candidate"]["controles_frecuencia"], "Mensual")
        self.assertEqual(processed["warnings"], [])

    def test_natural_language_dropdowns_do_not_fail_validation(self):
        payload = build_empty_extraction_payload("section_4_2_a_habilidades")
        payload["candidate"]["desplazamiento_modo"] = "se desplaza independiente con apoyo temporal"
        payload["candidate"]["ubicacion_ciudad"] = "se ubica usando maps"
        self.assertEqual(validate_extraction_payload(payload), [])

    def test_unresolved_dropdown_values_are_dropped_with_warning(self):
        payload = build_empty_extraction_payload("section_4_1_salud")
        payload["candidate"]["restriccion_conocimiento"] = "respuesta rara que no coincide"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_4_1_salud",
            candidate_index=1,
        )
        self.assertNotIn("restriccion_conocimiento", processed["candidate"])
        warnings = " ".join(processed["warnings"]).lower()
        self.assertIn("opcion valida", warnings)

    def test_section_4_2_a_derives_levels_from_main_dropdowns(self):
        payload = build_empty_extraction_payload("section_4_2_a_habilidades")
        payload["candidate"]["desplazamiento_modo"] = "independiente con apoyo temporal"
        payload["candidate"]["desplazamiento_transporte"] = "usa transmilenio"
        payload["candidate"]["ubicacion_ciudad"] = "se ubica usando maps"
        payload["candidate"]["dinero_reconocimiento"] = "con apoyo familiar"
        payload["candidate"]["dinero_manejo"] = "maneja el dinero pero en ocasiones requiere apoyo"
        payload["candidate"]["dinero_medios"] = "dinero fisico y tarjeta"
        payload["candidate"]["presentacion_personal"] = "acorde al contexto"
        payload["candidate"]["toma_decisiones"] = "a veces consulta a la madre"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_4_2_a_habilidades",
            candidate_index=1,
        )
        self.assertEqual(
            processed["candidate"]["desplazamiento_modo"],
            "1. Se desplaza de forma independiente con un apoyo temporal (ortesis, baston, silla de ruedas entre otros).",
        )
        self.assertEqual(processed["candidate"]["desplazamiento_nivel_apoyo"], "1. Nivel de apoyo Bajo.")
        self.assertEqual(processed["candidate"]["desplazamiento_transporte"], "Transmilenio, Sitp.")
        self.assertEqual(
            processed["candidate"]["ubicacion_ciudad"],
            "1. Sabe ubicarse en la ciudad pero haciendo uso de aplicaciones (Maps, Waze, entre otros).",
        )
        self.assertEqual(processed["candidate"]["ubicacion_nivel_apoyo"], "1. Nivel de apoyo Bajo.")
        self.assertEqual(processed["candidate"]["dinero_reconocimiento"], "Con apoyo familiar.")
        self.assertEqual(
            processed["candidate"]["dinero_manejo"],
            "1. Reconoce y maneja el dinero pero en ocasiones requiere apoyo.",
        )
        self.assertEqual(processed["candidate"]["dinero_nivel_apoyo"], "1. Nivel de apoyo Bajo.")
        self.assertEqual(processed["candidate"]["dinero_medios"], "Dinero fisico y plastico.")
        self.assertEqual(
            processed["candidate"]["presentacion_personal"],
            "0. Su codigo de vestuario es acorde al contexto.",
        )
        self.assertEqual(processed["candidate"]["presentacion_nivel_apoyo"], "0. No requiere apoyo.")
        self.assertEqual(
            processed["candidate"]["toma_decisiones"],
            "1. Toma decisiones pero en ocasiones requiere el apoyo de un tercero.",
        )
        self.assertEqual(processed["candidate"]["decisiones_nivel_apoyo"], "1. Nivel de apoyo Bajo.")

    def test_section_4_2_a_semantic_payload_maps_to_dropdowns(self):
        payload = build_empty_extraction_payload("section_4_2_a_habilidades")
        payload["semantic"]["section_4_2_a_skills"] = {
            "mobility": {
                "support_level": "low",
                "mode": "temporary support",
                "transport": "mass transit",
                "details": "Usa baston.",
            },
            "orientation": {
                "support_level": "low",
                "city_status": "apps",
                "references_status": "references",
                "details": "Usa Maps cuando lo necesita.",
            },
            "money": {
                "support_level": "low",
                "recognition": "family support",
                "management": "occasional support",
                "mediums": ["cash", "card"],
                "details": "Consulta a la madre para pagos grandes.",
            },
            "presentation": {
                "support_level": "none",
                "dress_code": "appropriate",
                "details": None,
            },
            "written_communication": {
                "support_level": "medium",
                "support_status": "knows not uses",
                "details": None,
            },
            "verbal_communication": {
                "support_level": "not applicable",
                "support_status": "not applicable",
                "details": None,
            },
            "decision_making": {
                "support_level": "low",
                "decision_status": "occasional support",
                "details": "Consulta a la madre.",
            },
        }
        self.assertEqual(validate_extraction_payload(payload), [])
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_4_2_a_habilidades",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["desplazamiento_nivel_apoyo"], "1. Nivel de apoyo Bajo.")
        self.assertEqual(
            processed["candidate"]["desplazamiento_modo"],
            "1. Se desplaza de forma independiente con un apoyo temporal (ortesis, baston, silla de ruedas entre otros).",
        )
        self.assertEqual(processed["candidate"]["desplazamiento_transporte"], "Transmilenio, Sitp.")
        self.assertEqual(processed["candidate"]["ubicacion_ciudad"], "1. Sabe ubicarse en la ciudad pero haciendo uso de aplicaciones (Maps, Waze, entre otros).")
        self.assertEqual(processed["candidate"]["ubicacion_aplicaciones"], "Se ubica por puntos de referencia y direcciones.")
        self.assertEqual(processed["candidate"]["dinero_medios"], "Dinero fisico y plastico.")
        self.assertEqual(processed["candidate"]["presentacion_personal"], "0. Su codigo de vestuario es acorde al contexto.")
        self.assertEqual(processed["candidate"]["comunicacion_escrita_apoyo"], "2. Conoce pero no maneja apoyos.")
        self.assertEqual(processed["candidate"]["comunicacion_verbal_apoyo"], "No aplica.")
        self.assertEqual(processed["candidate"]["toma_decisiones"], "1. Toma decisiones pero en ocasiones requiere el apoyo de un tercero.")

    def test_section_4_2_b_main_negative_answers_fill_missing_subitems(self):
        payload = build_empty_extraction_payload("section_4_2_b_actividades")
        payload["candidate"]["alimentacion"] = "no requiere apoyo en sus actividades de la vida diaria"
        payload["candidate"]["instrumentales_actividades"] = "no aplica"
        payload["candidate"]["actividades_apoyo"] = "no requiere apoyo en sus actividades laborales"
        payload["candidate"]["discriminacion"] = "no ha sufrido discriminacion"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_4_2_b_actividades",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["aseo_nivel_apoyo"], "0. No requiere apoyo.")
        self.assertEqual(processed["candidate"]["aseo_alimentacion"], "No")
        self.assertEqual(processed["candidate"]["aseo_higiene_aseo"], "No")
        self.assertEqual(processed["candidate"]["instrumentales_nivel_apoyo"], "No aplica.")
        self.assertEqual(processed["candidate"]["instrumentales_finanzas"], "No aplica")
        self.assertEqual(processed["candidate"]["instrumentales_salud_cuenta_apoyo"], "No aplica")
        self.assertEqual(processed["candidate"]["actividades_nivel_apoyo"], "0. No requiere apoyo.")
        self.assertEqual(processed["candidate"]["actividades_complementarios_apoyo"], "No")
        self.assertEqual(processed["candidate"]["discriminacion_nivel_apoyo"], "0. No requiere apoyo.")
        self.assertEqual(processed["candidate"]["discriminacion_violencia_apoyo"], "No")
        self.assertEqual(processed["candidate"]["discriminacion_vulneracion_cuenta_apoyo"], "No")

    def test_section_4_2_b_semantic_payload_maps_to_dropdowns(self):
        payload = build_empty_extraction_payload("section_4_2_b_actividades")
        payload["semantic"]["section_4_2_b_support"] = {
            "daily_living": {
                "support_level": "low",
                "scope": "some",
                "child_care": "not applicable",
                "communication_systems": "no",
                "assistive_devices": "no",
                "feeding": "yes",
                "functional_mobility": "no",
                "hygiene": "no",
                "details": "Solo requiere apoyo para alimentacion.",
            },
            "instrumental": {
                "support_level": "low",
                "scope": "some",
                "child_care": "not applicable",
                "communication_systems": "no",
                "community_mobility": "no",
                "finances": "yes",
                "cooking_cleaning": "yes",
                "household": "no",
                "health_support": "no",
                "details": None,
            },
            "work_activities": {
                "support_level": "low",
                "scope": "some",
                "family_recreation_requires_support": "no",
                "family_recreation_has_support": "no",
                "medical_followup_requires_support": "yes",
                "medical_followup_has_support": "no",
                "children_subsidies_has_support": "not applicable",
                "details": None,
            },
            "discrimination": {
                "support_level": "low",
                "scope": "some contexts",
                "physical_violence_requires_support": "no",
                "physical_violence_has_support": "no",
                "rights_violation_requires_support": "no",
                "rights_violation_has_support": "no",
                "details": "Refiere discriminacion laboral previa.",
            },
        }
        self.assertEqual(validate_extraction_payload(payload), [])
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_4_2_b_actividades",
            candidate_index=1,
        )
        self.assertEqual(
            processed["candidate"]["alimentacion"],
            "1. Requiere apoyo en algunas actividades de la vida diaria.",
        )
        self.assertEqual(processed["candidate"]["aseo_alimentacion"], "Si")
        self.assertEqual(processed["candidate"]["aseo_criar_apoyo"], "No aplica")
        self.assertEqual(processed["candidate"]["instrumentales_finanzas"], "Si")
        self.assertEqual(processed["candidate"]["instrumentales_cocina_limpieza"], "Si")
        self.assertEqual(processed["candidate"]["actividades_complementarios_apoyo"], "Si")
        self.assertEqual(processed["candidate"]["actividades_subsidios_cuenta_apoyo"], "No aplica")
        self.assertEqual(
            processed["candidate"]["discriminacion"],
            "1. Ha sufrido de discriminacion en algunos contextos.",
        )
        self.assertEqual(processed["candidate"]["discriminacion_violencia_apoyo"], "No")

    def test_section_4_2_b_binary_phrases_are_normalized_without_closing_rest(self):
        payload = build_empty_extraction_payload("section_4_2_b_actividades")
        payload["candidate"]["instrumentales_actividades"] = "requiere apoyo en algunas actividades instrumentales"
        payload["candidate"]["instrumentales_finanzas"] = "si necesita apoyo"
        payload["candidate"]["instrumentales_cocina_limpieza"] = "no"
        processed = postprocess_extraction_payload(
            payload,
            subsection_key="section_4_2_b_actividades",
            candidate_index=1,
        )
        self.assertEqual(processed["candidate"]["instrumentales_nivel_apoyo"], "1. Nivel de apoyo Bajo.")
        self.assertEqual(
            processed["candidate"]["instrumentales_actividades"],
            "1. Requiere apoyo en algunas actividades instrumentales de la vida diaria.",
        )
        self.assertEqual(processed["candidate"]["instrumentales_finanzas"], "Si")
        self.assertEqual(processed["candidate"]["instrumentales_cocina_limpieza"], "No")
        self.assertNotIn("instrumentales_comunicacion_apoyo", processed["candidate"])


if __name__ == "__main__":
    unittest.main()
