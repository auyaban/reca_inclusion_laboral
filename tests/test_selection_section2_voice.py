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
        payload["candidate"]["resultado_certificado"] = "Tal vez"
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
        self.assertIn("lo demas no", examples.lower())
        self.assertIn("solo", examples.lower())
        block_fields = get_subsection_spec("section_2_fields")
        prompt_fragment = str(block_fields.get("prompt_fragment") or "").lower()
        self.assertIn("tipo_pension", prompt_fragment)
        self.assertIn("por confirmar", " ".join(block_fields.get("examples") or []).lower())

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


if __name__ == "__main__":
    unittest.main()
