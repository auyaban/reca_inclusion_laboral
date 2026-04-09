from __future__ import annotations

import unittest
from types import SimpleNamespace
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
        with patch("drive_upload.publish_sheet_from_template") as publish_sheet:
            with self.assertRaises(RuntimeError):
                vacante.export_to_excel()
        publish_sheet.assert_not_called()

    def test_condiciones_vacante_export_prefers_explicit_cache_over_form_cache(self) -> None:
        vacante.FORM_CACHE["section_1"] = {"nombre_empresa": "Empresa Incorrecta"}
        explicit_cache = {
            "section_1": {
                "nombre_empresa": "Empresa Correcta",
                "nit_empresa": "900123456",
                "fecha_visita": "2026-04-09",
            },
            "section_2": {"nombre_vacante": "Analista"},
        }
        captured_payloads = {}

        def _fake_build_section_writes(section_id, payload):
            captured_payloads[section_id] = payload
            return []

        with patch.object(vacante, "validate_before_finalize", return_value=[]):
            with patch.object(vacante, "_build_section_writes", side_effect=_fake_build_section_writes):
                with patch.object(vacante, "_build_row_insertions", return_value=[]):
                    with patch("google_sheets_client.get_master_template_id", return_value="template-123"):
                        with patch(
                            "drive_upload.publish_sheet_from_template",
                            return_value={"webViewLink": "https://drive.test/file", "file_id": "file-123"},
                        ) as publish_sheet:
                            result = vacante.export_to_excel(cache=explicit_cache)

        self.assertEqual(captured_payloads["section_1"]["nombre_empresa"], "Empresa Correcta")
        self.assertEqual(captured_payloads["section_2"]["nombre_vacante"], "Analista")
        self.assertEqual(result["output_path"], "https://drive.test/file")
        self.assertEqual(result["drive_file_id"], "file-123")
        self.assertEqual(result["acta_metadata"]["nombre_empresa"], "Empresa Correcta")
        self.assertEqual(
            publish_sheet.call_args.kwargs["folder_name"],
            vacante._sanitize_filename("Empresa Correcta"),
        )


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


class FinalizeDraftSessionTests(unittest.TestCase):
    def tearDown(self) -> None:
        vacante.clear_form_cache()

    def test_original_finalization_uses_window_draft_snapshot_without_mutating_session(self) -> None:
        vacante.FORM_CACHE["section_1"] = {"nombre_empresa": "Empresa Incorrecta"}
        window = SimpleNamespace(
            _form_id="condiciones_vacante",
            _draft_cache={"section_1": {"nombre_empresa": "Empresa Correcta", "fecha_visita": "2026-04-09"}},
            _finalize_in_progress=False,
            destroy=lambda: None,
        )
        loading = SimpleNamespace()
        worker_inputs = []
        completion_calls = {}

        def _worker(cache_snapshot=None):
            worker_inputs.append(cache_snapshot)
            cache_snapshot["section_1"]["nombre_empresa"] = "Mutado en worker"
            return {
                "output_path": "https://drive.test/file",
                "already_in_drive": True,
                "drive_file_id": "file-123",
                "tipo_acta": "sin_pdf",
                "fecha_servicio": "2026-04-09",
                "acta_metadata": {},
            }

        hub = SimpleNamespace(
            current_session_id="session-123",
            finalize_form_delivery=lambda *args, **kwargs: {"status": "synced", "remote_url": "https://drive.test/file"},
        )

        def _fake_completion_payload(form_id, form_name, cache_snapshot, **kwargs):
            completion_calls["form_id"] = form_id
            completion_calls["cache_snapshot"] = cache_snapshot
            completion_calls["extra_context"] = kwargs.get("extra_context") or {}
            return {"ok": True}

        class _ImmediateThread:
            def __init__(self, target=None, daemon=None):
                self._target = target

            def start(self):
                if self._target is not None:
                    self._target()

        review_result = SimpleNamespace(status="skipped", reviewed_count=0, elapsed_ms=0, reason="test")

        with patch.object(app, "_resolve_hub_window", return_value=hub):
            with patch.object(app.text_review, "review_export_cache", return_value=review_result):
                with patch.object(app, "_safe_widget_after", side_effect=lambda _window, fn: fn()):
                    with patch.object(app, "_close_loading_async", return_value=None):
                        with patch.object(app, "_update_loading_async", return_value=None):
                            with patch.object(app, "_finalize_export_flow", return_value=None):
                                with patch.object(app, "_return_to_hub", return_value=None):
                                    with patch.object(app.completion_payloads, "build_completion_payload", side_effect=_fake_completion_payload):
                                        with patch.object(app.threading, "Thread", _ImmediateThread):
                                            result = app._original_start_background_finalization(
                                                window,
                                                loading,
                                                form_name="Revision Condicion",
                                                company_name="Empresa Correcta",
                                                form_id="condiciones_vacante",
                                                worker_fn=_worker,
                                                show_completion_ui=False,
                                                return_to_hub_on_success=False,
                                                close_window_on_success=False,
                                            )

        self.assertTrue(result)
        self.assertEqual(worker_inputs[0]["section_1"]["nombre_empresa"], "Mutado en worker")
        self.assertEqual(window._draft_cache["section_1"]["nombre_empresa"], "Empresa Correcta")
        self.assertEqual(completion_calls["form_id"], "condiciones_vacante")
        self.assertEqual(
            completion_calls["cache_snapshot"]["section_1"]["nombre_empresa"],
            "Empresa Correcta",
        )
        self.assertEqual(completion_calls["extra_context"].get("payload_source"), "draft_session")


if __name__ == "__main__":
    unittest.main()
