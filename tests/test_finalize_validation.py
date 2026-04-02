from __future__ import annotations

import unittest
from unittest.mock import patch

import app
from formularios.condiciones_vacante import condiciones_vacante as vacante
from formularios.evaluacion_programa import evaluacion_accesibilidad as accesibilidad
from formularios.finalize_validation import (
    ValidationIssue,
    format_issues_for_message,
    is_meaningful,
    validate_dynamic_rows,
)
from formularios.sensibilizacion import sensibilizacion


def _filled_fields(fields):
    payload = {}
    for field in fields or []:
        if not isinstance(field, dict):
            continue
        field_id = str(field.get("id") or "").strip()
        if not field_id:
            continue
        payload[field_id] = f"valor-{field_id}"
    return payload


def _filled_items(items):
    payload = {}
    for item in items or []:
        if not isinstance(item, dict):
            continue
        item_id = str(item.get("id") or "").strip()
        if not item_id:
            continue
        payload[item_id] = f"valor-{item_id}"
        payload[f"{item_id}_nota"] = f"Nota: valor-{item_id}"
    return payload


class FinalizeValidationHelperTests(unittest.TestCase):
    def test_is_meaningful_handles_empty_and_valid_values(self) -> None:
        self.assertFalse(is_meaningful(""))
        self.assertFalse(is_meaningful("   "))
        self.assertFalse(is_meaningful("Nota: "))
        self.assertTrue(is_meaningful("No aplica"))
        self.assertTrue(is_meaningful(False))
        self.assertTrue(is_meaningful(0))

    def test_validate_dynamic_rows_ignores_blank_row_and_flags_partial_row(self) -> None:
        issues = []
        validate_dynamic_rows(
            issues,
            "section_8",
            [
                {"nombre": "", "cargo": ""},
                {"nombre": "Sandra", "cargo": ""},
            ],
            [("nombre", "Nombre"), ("cargo", "Cargo")],
        )
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0].section_id, "section_8")
        self.assertEqual(issues[0].field_id, "cargo")
        self.assertEqual(issues[0].row_index, 2)

    def test_format_issues_for_message_lists_section_and_field(self) -> None:
        message = format_issues_for_message(
            [
                ValidationIssue(
                    section_id="section_5",
                    field_id="postura_trabajo",
                    label="Postura de trabajo",
                    message="Campo obligatorio sin diligenciar.",
                )
            ]
        )
        self.assertIn("No se puede finalizar el formato.", message)
        self.assertIn("Seccion 5: Postura de trabajo", message)


class ModuleFinalizeValidationTests(unittest.TestCase):
    def setUp(self) -> None:
        vacante.clear_form_cache()
        accesibilidad.clear_form_cache()
        sensibilizacion.clear_form_cache()
        self.addCleanup(vacante.clear_form_cache)
        self.addCleanup(accesibilidad.clear_form_cache)
        self.addCleanup(sensibilizacion.clear_form_cache)

    def test_condiciones_vacante_validation_flags_section_5_ergonomic_fields(self) -> None:
        issues = vacante.validate_before_finalize(
            {
                "section_5": {
                    "ruido": "Bajo.",
                    "iluminacion": "Bajo.",
                }
            }
        )
        missing = {(issue.section_id, issue.field_id) for issue in issues}
        self.assertIn(("section_5", "postura_trabajo"), missing)
        self.assertIn(("section_5", "puesto_trabajo"), missing)
        self.assertIn(("section_5", "movimientos_repetitivos"), missing)
        self.assertIn(("section_5", "observaciones_peligros"), missing)

    def test_sensibilizacion_validation_accepts_only_required_sections(self) -> None:
        cache = {
            "section_1": _filled_fields(sensibilizacion.SECTION_1["fields"]),
            "section_3": {"observaciones": "Observaciones de prueba"},
            "section_5": [{"nombre": "Sandra", "cargo": "Coordinacion"}],
        }
        issues = sensibilizacion.validate_before_finalize(cache)
        self.assertEqual(issues, [])

    def test_evaluacion_accesibilidad_validation_accepts_section_8_without_tamano_empresa(self) -> None:
        cache = {
            "section_1": _filled_fields(accesibilidad.SECTION_1["fields"]),
            "section_4": {
                "nivel_accesibilidad": "Medio",
                "descripcion": "Descripcion de prueba",
            },
            "section_5": _filled_items(accesibilidad.SECTION_5.get("items")),
            "section_6": _filled_fields(accesibilidad.SECTION_6.get("fields")),
            "section_7": _filled_fields(accesibilidad.SECTION_7.get("fields")),
            "section_8": [
                {
                    "nombre": "Sandra",
                    "cargo": "Coordinacion",
                    "tamano_empresa": "",
                }
            ],
        }

        with patch.object(accesibilidad, "_validate_question_section", return_value=None):
            issues = accesibilidad.validate_before_finalize(cache)

        self.assertEqual(issues, [])

    def test_condiciones_vacante_export_stops_before_publish_when_validation_fails(self) -> None:
        with patch.object(vacante, "cache_file_exists", return_value=False):
            with patch("drive_upload.publish_sheet_from_template") as publish_sheet:
                with self.assertRaises(RuntimeError):
                    vacante.export_to_excel()
        publish_sheet.assert_not_called()


class FinalizePreflightGuardTests(unittest.TestCase):
    def test_start_background_finalization_blocks_on_validation_errors(self) -> None:
        closed = {"value": False}
        shown_sections = []

        class _DummyLoading:
            def close(self):
                closed["value"] = True

        class _DummyWindow:
            _form_id = "condiciones_vacante"

            def _show_section_5(self):
                shown_sections.append("section_5")

        worker_fn_called = {"value": False}

        with patch.object(app, "_guard_form_action", return_value=False):
            with patch.object(app.condiciones_vacante, "save_cache_to_file", return_value=None):
                with patch.object(app.condiciones_vacante, "get_form_cache", return_value={}):
                    with patch.object(
                        app.condiciones_vacante,
                        "validate_before_finalize",
                        return_value=[
                            ValidationIssue(
                                section_id="section_5",
                                field_id="postura_trabajo",
                                label="Postura de trabajo",
                                message="Campo obligatorio sin diligenciar.",
                            )
                        ],
                    ):
                        with patch.object(app, "_original_start_background_finalization") as original_start:
                            with patch.object(app.messagebox, "showerror", return_value=None) as showerror:
                                app._start_background_finalization(
                                    _DummyWindow(),
                                    _DummyLoading(),
                                    form_name="Revision Condicion",
                                    company_name="Empresa Demo",
                                    form_id="condiciones_vacante",
                                    worker_fn=lambda: worker_fn_called.update(value=True),
                                )

        self.assertTrue(closed["value"])
        self.assertEqual(shown_sections, ["section_5"])
        self.assertFalse(worker_fn_called["value"])
        original_start.assert_not_called()
        self.assertIn("Seccion 5: Postura de trabajo", showerror.call_args.args[1])


if __name__ == "__main__":
    unittest.main()
