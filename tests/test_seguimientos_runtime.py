from __future__ import annotations

import tkinter as tk
import os
import tempfile
import unittest
from unittest.mock import ANY, Mock, patch

import app
from formularios.seguimientos import seguimientos
from tests.tk_test_utils import TkTestCase, destroy_widget


class _TkTestCase(TkTestCase):
    def setUp(self) -> None:
        super().setUp()
        patchers = [
            patch.object(app, "_maximize_window", lambda _window: None),
            patch.object(app.messagebox, "showerror", return_value=None),
            patch.object(app.messagebox, "showinfo", return_value=None),
            patch.object(app.messagebox, "showwarning", return_value=None),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)


class SeguimientosRuntimeTests(_TkTestCase):
    @staticmethod
    def _run_loading_job_immediately(**kwargs):
        result = kwargs["worker"](lambda *args, **kw: None)
        kwargs["on_success"](result)

    @staticmethod
    def _run_async_task_immediately(window, **kwargs):
        try:
            result = kwargs["worker"]()
        except Exception as exc:
            kwargs["on_error"](exc)
            return True
        kwargs["on_success"](result)
        return True

    def _make_editor(
        self,
        *,
        workflow,
        case_record,
        base_payload=None,
        followup_payload=None,
        save_base_payload=None,
        save_followup_payload=None,
    ):
        patchers = [
            patch.object(app, "attach_dictation", return_value=None),
            patch.object(
                app,
                "_get_asistentes_profesionales_catalog",
                return_value={"nombres": [], "cargos": [], "name_to_cargo": {}},
            ),
            patch.object(
                app.seguimientos,
                "get_case_meta",
                return_value={
                    "max_seguimientos": workflow.get("max_seguimientos", 3),
                    "base_sheet_name": workflow.get("base_sheet_name", seguimientos.SHEET_BASE),
                },
            ),
            patch.object(app.seguimientos, "get_workflow_state", return_value=workflow),
            patch.object(
                app.seguimientos,
                "suggest_next_step",
                return_value={
                    "sheet": workflow.get("suggested_sheet") or workflow.get("base_sheet_name", seguimientos.SHEET_BASE),
                    "message": workflow.get("message", ""),
                },
            ),
            patch.object(app.seguimientos, "sync_case_record_from_local", return_value=None),
        ]
        if base_payload is not None:
            patchers.append(patch.object(app.seguimientos, "get_base_payload", return_value=base_payload))
        if followup_payload is not None:
            patchers.append(patch.object(app.seguimientos, "get_followup_payload", return_value=followup_payload))
        if save_base_payload is not None:
            patchers.append(patch.object(app.seguimientos, "save_base_payload", save_base_payload))
        if save_followup_payload is not None:
            patchers.append(patch.object(app.seguimientos, "save_followup_payload", save_followup_payload))
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        editor = app.SeguimientoEditorWindow(self.root, case_path="", case_record=case_record)
        self.addCleanup(destroy_widget, editor)
        return editor

    def _install_followup_local_draft_store(self):
        tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(tempdir.cleanup)
        path = os.path.join(tempdir.name, "seguimientos_local_drafts.json")
        patcher = patch.object(app, "_get_followup_local_drafts_path", return_value=path)
        patcher.start()
        self.addCleanup(patcher.stop)
        return path

    def _install_empty_hub_draft_store(self):
        tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(tempdir.cleanup)
        path = os.path.join(tempdir.name, "form_drafts_il.json")
        patcher = patch.object(app, "_get_drafts_path", return_value=path)
        patcher.start()
        self.addCleanup(patcher.stop)
        return path

    def test_register_form_disables_generic_drafts(self) -> None:
        meta = seguimientos.register_form()

        self.assertEqual(meta["id"], "seguimientos")
        self.assertIs(meta["supports_drafts"], False)

    def test_get_draft_save_command_returns_none_for_seguimientos(self) -> None:
        window = type(
            "DummyWindow",
            (),
            {
                "_form_id": "seguimientos",
                "_save_draft_command": None,
                "master": object(),
            },
        )()

        self.assertIsNone(app._get_draft_save_command(window))

    def test_bind_form_runtime_skips_generic_autosave_for_seguimientos(self) -> None:
        with patch.object(seguimientos, "get_usuarios_reca_cedulas", return_value=[]):
            window = app.SeguimientosWindow(self.root)
        self.addCleanup(destroy_widget, window)

        class HubStub:
            _empresa_names_cache = []

        hub = HubStub()
        with patch.object(app, "_attach_dictation_for_section", return_value=None):
            app.HubWindow._bind_form_runtime(hub, window, seguimientos.register_form())

        self.assertIsNone(window._save_draft_command)
        self.assertIsNone(window._draft_autosave_after_id)
        self.assertFalse(hasattr(window.cedula_combo, "_draft_autosave_bound"))

    def test_window_disables_search_and_shows_retry_hint_when_cedulas_fail(self) -> None:
        with patch.object(
            seguimientos,
            "get_usuarios_reca_cedulas",
            side_effect=RuntimeError("Supabase no esta disponible (HTTP 401): HTTP Error 401: Unauthorized"),
        ):
            window = app.SeguimientosWindow(self.root)
        self.addCleanup(destroy_widget, window)

        self.assertEqual(str(window.buscar_vinculado_btn.cget("state")), "disabled")
        self.assertEqual(tuple(window.cedula_combo.cget("values") or ()), ())
        self.assertIn("sesión o los permisos", str(window.cedula_hint_label.cget("text") or "").lower())

    def test_reload_cedulas_restores_search_after_successful_retry(self) -> None:
        with patch.object(
            seguimientos,
            "get_usuarios_reca_cedulas",
            side_effect=RuntimeError("Supabase no esta disponible (HTTP 401): HTTP Error 401: Unauthorized"),
        ):
            window = app.SeguimientosWindow(self.root)
        self.addCleanup(destroy_widget, window)

        with patch.object(
            app,
            "_run_async_ui_task",
            side_effect=lambda *args, **kwargs: self._run_async_task_immediately(*args, **kwargs),
        ):
            with patch.object(
                seguimientos,
                "get_usuarios_reca_cedulas",
                return_value=["10101010", "20202020"],
            ):
                window._reload_cedulas()

        self.assertEqual(str(window.buscar_vinculado_btn.cget("state")), "normal")
        self.assertIn("10101010", tuple(window.cedula_combo.cget("values") or ()))
        self.assertIn("recargada", str(window.cedula_hint_label.cget("text") or "").lower())

    def test_editor_smoke_renders_base_followup_and_final_and_saves(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        followup_sheet = f"{seguimientos.SHEET_PREFIX}1"
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-123",
            "webViewLink": "https://example.com/sheet-123",
        }
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, followup_sheet, seguimientos.SHEET_FINAL],
            "editable_sheet": followup_sheet,
            "suggested_sheet": base_sheet,
            "next_followup": 1,
            "message": "Editando la ficha inicial del proceso.",
        }
        base_payload = {
            "fecha_visita": "2026-03-26",
            "modalidad": "Presencial",
            "nombre_empresa": "Empresa Demo",
            "ciudad_empresa": "Bogotá",
            "direccion_empresa": "Calle 1",
            "nit_empresa": "900123456",
            "correo_1": "empresa@example.com",
            "telefono_empresa": "1234567",
            "contacto_empresa": "Ana",
            "cargo": "Líder",
            "asesor": "Asesor",
            "sede_empresa": "Centro",
            "caja_compensacion": "Compensar",
            "profesional_asignado": "Profesional RECA",
            "nombre_vinculado": "Persona Demo",
            "cedula": "10101010",
            "telefono_vinculado": "3000000000",
            "correo_vinculado": "persona@example.com",
            "contacto_emergencia": "Contacto",
            "parentesco": "Familiar",
            "telefono_emergencia": "3111111111",
            "cargo_vinculado": "Analista",
            "certificado_discapacidad": "Si",
            "certificado_porcentaje": "25%",
            "discapacidad": "Auditiva",
            "tipo_contrato": "Indefinido",
            "fecha_firma_contrato": "2026-03-01",
            "fecha_inicio_contrato": "2026-03-05",
            "fecha_fin_contrato": "",
            "apoyos_ajustes": "Ajustes mínimos",
            "funciones_1_5": ["F1", "F2", "F3", "F4", "F5"],
            "funciones_6_10": ["F6", "F7", "F8", "F9", "F10"],
            "seguimiento_fechas_1_3": ["2026-03-26", "", ""],
            "seguimiento_fechas_4_6": ["", "", ""],
        }
        followup_payload = {
            "modalidad": "Mixta",
            "tipo_apoyo": seguimientos.TIPO_APOYO_OPTIONS[0],
            "item_labels": ["Asistencia", "Puntualidad"],
            "item_observaciones": ["Sin novedades", "Cumple"],
            "item_autoevaluacion": ["Excelente", "Bien"],
            "item_eval_empresa": ["Bien", "Excelente"],
            "empresa_item_labels": ["Comunicación", "Productividad"],
            "empresa_eval": ["Bien", "Excelente"],
            "empresa_observacion": ["Estable", "Destacado"],
            "situacion_encontrada": "Sin hallazgos críticos",
            "estrategias_ajustes": "Seguimiento quincenal",
            "asistentes": [
                {"nombre": "Profesional RECA", "cargo": "Profesional"},
                {"nombre": "Jefe directo", "cargo": "Coordinador"},
            ],
        }
        save_base_payload = Mock()
        save_followup_payload = Mock()

        patchers = [
            patch.object(app, "attach_dictation", return_value=None),
            patch.object(
                app,
                "_get_asistentes_profesionales_catalog",
                return_value={"nombres": [], "cargos": [], "name_to_cargo": {}},
            ),
            patch.object(app.seguimientos, "get_case_meta", return_value={"max_seguimientos": 3, "base_sheet_name": base_sheet}),
            patch.object(app.seguimientos, "get_workflow_state", return_value=workflow),
            patch.object(
                app.seguimientos,
                "suggest_next_step",
                return_value={"sheet": base_sheet, "message": "Editar ficha inicial del proceso"},
            ),
            patch.object(app.seguimientos, "get_base_payload", return_value=base_payload),
            patch.object(app.seguimientos, "get_followup_payload", return_value=followup_payload),
            patch.object(app.seguimientos, "save_base_payload", save_base_payload),
            patch.object(app.seguimientos, "save_followup_payload", save_followup_payload),
            patch.object(app.seguimientos, "sync_case_record_from_local", return_value=None),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        editor = app.SeguimientoEditorWindow(self.root, case_path="", case_record=case_record)
        self.addCleanup(destroy_widget, editor)
        editor._run_loading_job = self._run_loading_job_immediately

        self.assertEqual(editor.sheet_var.get(), base_sheet)
        self.assertEqual(editor.base_vars["nombre_empresa"].get(), "Empresa Demo")

        editor._save_current_sheet()
        save_base_payload.assert_called_once_with(editor.case_target, ANY)

        editor.sheet_var.set(followup_sheet)
        editor._render_selected_sheet()
        self.assertEqual(editor.current_followup_index, 1)
        self.assertEqual(editor.follow_vars["modalidad"].get(), "Mixta")

        editor._save_current_sheet()
        save_followup_payload.assert_called_once_with(editor.case_target, 1, ANY)

        editor.sheet_var.set(seguimientos.SHEET_FINAL)
        editor._render_selected_sheet()
        self.assertEqual(editor.status_var.get(), "Resultado final es de solo lectura.")

    def test_editor_renders_company_fields_readonly_and_blank_dates(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-blank",
            "webViewLink": "https://example.com/sheet-blank",
        }
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, f"{seguimientos.SHEET_PREFIX}1", seguimientos.SHEET_FINAL],
            "editable_sheet": f"{seguimientos.SHEET_PREFIX}1",
            "suggested_sheet": base_sheet,
            "next_followup": 1,
            "message": "Empieza por la ficha inicial del proceso.",
        }
        base_payload = {
            "fecha_visita": "2026-03-26",
            "modalidad": "Presencial",
            "nombre_empresa": "Empresa Demo",
            "ciudad_empresa": "Bogotá",
            "direccion_empresa": "Calle 1",
            "nit_empresa": "900123456",
            "correo_1": "empresa@example.com",
            "telefono_empresa": "1234567",
            "contacto_empresa": "Ana",
            "cargo": "Líder",
            "asesor": "Asesor",
            "sede_empresa": "Centro",
            "caja_compensacion": "Compensar",
            "profesional_asignado": "Profesional RECA",
            "nombre_vinculado": "Persona Demo",
            "cedula": "10101010",
            "telefono_vinculado": "3000000000",
            "correo_vinculado": "persona@example.com",
            "contacto_emergencia": "Contacto",
            "parentesco": "Familiar",
            "telefono_emergencia": "3111111111",
            "cargo_vinculado": "Analista",
            "certificado_discapacidad": "Si",
            "certificado_porcentaje": "25%",
            "discapacidad": "Auditiva",
            "tipo_contrato": "Indefinido",
            "fecha_firma_contrato": "",
            "fecha_inicio_contrato": "",
            "fecha_fin_contrato": "",
            "apoyos_ajustes": "",
            "funciones_1_5": ["", "", "", "", ""],
            "funciones_6_10": ["", "", "", "", ""],
            "seguimiento_fechas_1_3": ["", "", ""],
            "seguimiento_fechas_4_6": ["", "", ""],
        }

        patchers = [
            patch.object(app, "attach_dictation", return_value=None),
            patch.object(
                app,
                "_get_asistentes_profesionales_catalog",
                return_value={"nombres": [], "cargos": [], "name_to_cargo": {}},
            ),
            patch.object(app.seguimientos, "get_case_meta", return_value={"max_seguimientos": 3, "base_sheet_name": base_sheet}),
            patch.object(app.seguimientos, "get_workflow_state", return_value=workflow),
            patch.object(
                app.seguimientos,
                "suggest_next_step",
                return_value={"sheet": base_sheet, "message": "Editar ficha inicial del proceso"},
            ),
            patch.object(app.seguimientos, "get_base_payload", return_value=base_payload),
            patch.object(app.seguimientos, "sync_case_record_from_local", return_value=None),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        editor = app.SeguimientoEditorWindow(self.root, case_path="", case_record=case_record)
        self.addCleanup(destroy_widget, editor)

        self.assertIsNone(editor.company_name_combo)
        self.assertEqual(editor.base_company_widgets["nombre_empresa"].cget("state"), "readonly")
        self.assertEqual(editor.base_company_widgets["nit_empresa"].cget("state"), "readonly")
        self.assertEqual(editor.base_date_widgets["fecha_firma_contrato"].get().strip(), "")
        self.assertEqual(editor.base_date_widgets["fecha_inicio_contrato"].get().strip(), "")
        self.assertEqual(editor.base_date_widgets["fecha_fin_contrato"].get().strip(), "")
        self.assertEqual(len(editor.base_dates_1), 3)
        self.assertEqual(len(editor.base_dates_2), 0)
        self.assertEqual(editor.base_dates_1[0].get().strip(), "")
        self.assertEqual(editor.base_dates_1[1].get().strip(), "")
        self.assertEqual(editor.base_dates_1[2].get().strip(), "")
        self.assertEqual(str(editor.base_dates_1[0].select_button.cget("state")), "disabled")
        self.assertEqual(str(editor.base_dates_1[1].select_button.cget("state")), "disabled")
        self.assertEqual(str(editor.base_dates_1[2].select_button.cget("state")), "disabled")

    def test_editor_base_contract_end_field_accepts_text_and_no_aplica_shortcut(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-contract-end-text",
            "webViewLink": "https://example.com/sheet-contract-end-text",
        }
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, f"{seguimientos.SHEET_PREFIX}1", seguimientos.SHEET_FINAL],
            "editable_sheet": base_sheet,
            "suggested_sheet": base_sheet,
            "next_followup": 1,
            "message": "Empieza por la ficha inicial del proceso.",
        }
        base_payload = {
            "fecha_visita": "2026-03-26",
            "modalidad": "Presencial",
            "nombre_empresa": "Empresa Demo",
            "cedula": "10101010",
            "fecha_fin_contrato": "",
            "seguimiento_fechas_1_3": ["", "", ""],
            "seguimiento_fechas_4_6": ["", "", ""],
        }

        editor = self._make_editor(
            workflow=workflow,
            case_record=case_record,
            base_payload=base_payload,
        )

        field = editor.base_date_widgets["fecha_fin_contrato"]
        self.assertEqual(str(field.entry.cget("state")), "normal")
        field.entry.delete(0, tk.END)
        field.entry.insert(0, "Hasta nuevo aviso")
        self.assertEqual(field.get(), "Hasta nuevo aviso")

        editor.base_date_na_button.invoke()
        self.assertEqual(field.get(), "No aplica")

        request = editor._build_sheet_save_request(sheet_name=base_sheet, validate_base=False)
        self.assertEqual(request["payload"]["fecha_fin_contrato"], "No aplica")

    def test_editor_blocks_base_save_without_required_fecha_visita_and_modalidad(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-missing-required",
            "webViewLink": "https://example.com/sheet-missing-required",
        }
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, f"{seguimientos.SHEET_PREFIX}1", seguimientos.SHEET_FINAL],
            "editable_sheet": base_sheet,
            "suggested_sheet": base_sheet,
            "next_followup": 1,
            "message": "Completa la hoja base.",
        }
        base_payload = {
            "fecha_visita": "",
            "modalidad": "",
            "nombre_empresa": "Empresa Demo",
            "ciudad_empresa": "Bogotá",
            "direccion_empresa": "Calle 1",
            "nit_empresa": "900123456",
            "correo_1": "empresa@example.com",
            "telefono_empresa": "1234567",
            "contacto_empresa": "Ana",
            "cargo": "Líder",
            "asesor": "Asesor",
            "sede_empresa": "Centro",
            "caja_compensacion": "Compensar",
            "profesional_asignado": "Profesional RECA",
            "nombre_vinculado": "Persona Demo",
            "cedula": "10101010",
            "telefono_vinculado": "3000000000",
            "correo_vinculado": "persona@example.com",
            "contacto_emergencia": "Contacto",
            "parentesco": "Familiar",
            "telefono_emergencia": "3111111111",
            "cargo_vinculado": "Analista",
            "certificado_discapacidad": "Si",
            "certificado_porcentaje": "25%",
            "discapacidad": "Auditiva",
            "tipo_contrato": "Indefinido",
            "fecha_firma_contrato": "",
            "fecha_inicio_contrato": "",
            "fecha_fin_contrato": "",
            "apoyos_ajustes": "",
            "funciones_1_5": ["", "", "", "", ""],
            "funciones_6_10": ["", "", "", "", ""],
            "seguimiento_fechas_1_3": ["", "", ""],
            "seguimiento_fechas_4_6": ["", "", ""],
        }
        save_base_payload = Mock()

        patchers = [
            patch.object(app, "attach_dictation", return_value=None),
            patch.object(
                app,
                "_get_asistentes_profesionales_catalog",
                return_value={"nombres": [], "cargos": [], "name_to_cargo": {}},
            ),
            patch.object(app.seguimientos, "get_case_meta", return_value={"max_seguimientos": 3, "base_sheet_name": base_sheet}),
            patch.object(app.seguimientos, "get_workflow_state", return_value=workflow),
            patch.object(app.seguimientos, "suggest_next_step", return_value={"sheet": base_sheet, "message": workflow["message"]}),
            patch.object(app.seguimientos, "get_base_payload", return_value=base_payload),
            patch.object(app.seguimientos, "save_base_payload", save_base_payload),
            patch.object(app.seguimientos, "sync_case_record_from_local", return_value=None),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        editor = app.SeguimientoEditorWindow(self.root, case_path="", case_record=case_record)
        self.addCleanup(destroy_widget, editor)
        editor._run_loading_job = Mock()

        with patch.object(app.messagebox, "showerror", return_value=None) as showerror:
            editor._save_current_sheet()

        save_base_payload.assert_not_called()
        editor._run_loading_job.assert_not_called()
        showerror.assert_not_called()
        self.assertIn("fecha de visita y modalidad", editor.status_var.get().lower())

    def test_editor_base_sheet_only_enables_next_followup_date(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-followup-2",
            "webViewLink": "https://example.com/sheet-followup-2",
        }
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, f"{seguimientos.SHEET_PREFIX}1", f"{seguimientos.SHEET_PREFIX}2", seguimientos.SHEET_FINAL],
            "editable_sheet": f"{seguimientos.SHEET_PREFIX}2",
            "suggested_sheet": base_sheet,
            "next_followup": 2,
            "message": "Seguimiento 2 habilitado. En la hoja base solo puedes editar la fecha del seguimiento 2.",
        }
        base_payload = {
            "fecha_visita": "2026-03-26",
            "modalidad": "Presencial",
            "nombre_empresa": "Empresa Demo",
            "ciudad_empresa": "Bogotá",
            "direccion_empresa": "Calle 1",
            "nit_empresa": "900123456",
            "correo_1": "empresa@example.com",
            "telefono_empresa": "1234567",
            "contacto_empresa": "Ana",
            "cargo": "Líder",
            "asesor": "Asesor",
            "sede_empresa": "Centro",
            "caja_compensacion": "Compensar",
            "profesional_asignado": "Profesional RECA",
            "nombre_vinculado": "Persona Demo",
            "cedula": "10101010",
            "telefono_vinculado": "3000000000",
            "correo_vinculado": "persona@example.com",
            "contacto_emergencia": "Contacto",
            "parentesco": "Familiar",
            "telefono_emergencia": "3111111111",
            "cargo_vinculado": "Analista",
            "certificado_discapacidad": "Si",
            "certificado_porcentaje": "25%",
            "discapacidad": "Auditiva",
            "tipo_contrato": "Indefinido",
            "fecha_firma_contrato": "",
            "fecha_inicio_contrato": "",
            "fecha_fin_contrato": "",
            "apoyos_ajustes": "",
            "funciones_1_5": ["", "", "", "", ""],
            "funciones_6_10": ["", "", "", "", ""],
            "seguimiento_fechas_1_3": ["2026-04-01", "2026-05-01", ""],
            "seguimiento_fechas_4_6": ["", "", ""],
        }

        patchers = [
            patch.object(app, "attach_dictation", return_value=None),
            patch.object(
                app,
                "_get_asistentes_profesionales_catalog",
                return_value={"nombres": [], "cargos": [], "name_to_cargo": {}},
            ),
            patch.object(app.seguimientos, "get_case_meta", return_value={"max_seguimientos": 3, "base_sheet_name": base_sheet}),
            patch.object(app.seguimientos, "get_workflow_state", return_value=workflow),
            patch.object(app.seguimientos, "suggest_next_step", return_value={"sheet": base_sheet, "message": workflow["message"]}),
            patch.object(app.seguimientos, "get_base_payload", return_value=base_payload),
            patch.object(app.seguimientos, "sync_case_record_from_local", return_value=None),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        editor = app.SeguimientoEditorWindow(self.root, case_path="", case_record=case_record)
        self.addCleanup(destroy_widget, editor)

        self.assertEqual(str(editor.base_dates_1[0].select_button.cget("state")), "disabled")
        self.assertEqual(str(editor.base_dates_1[1].select_button.cget("state")), "disabled")
        self.assertEqual(str(editor.base_dates_1[2].select_button.cget("state")), "disabled")

    def test_copy_previous_followup_keeps_date_and_large_texts_but_copies_other_fields(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        followup_1 = f"{seguimientos.SHEET_PREFIX}1"
        followup_2 = f"{seguimientos.SHEET_PREFIX}2"
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-copy-followup-2",
            "webViewLink": "https://example.com/sheet-copy-followup-2",
        }
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, followup_1, followup_2, seguimientos.SHEET_FINAL],
            "editable_sheet": followup_2,
            "suggested_sheet": followup_2,
            "next_followup": 2,
            "message": "Continua con Seguimiento 2.",
        }
        current_payload = {
            "modalidad": "Virtual",
            "fecha_seguimiento": "2026-05-10",
            "tipo_apoyo": "",
            "item_labels": ["Asistencia", "Puntualidad"],
            "item_observaciones": ["Actual 1", "Actual 2"],
            "item_autoevaluacion": ["", ""],
            "item_eval_empresa": ["", ""],
            "empresa_item_labels": ["Comunicacion", "Productividad"],
            "empresa_eval": ["", ""],
            "empresa_observacion": ["Actual emp 1", "Actual emp 2"],
            "situacion_encontrada": "Texto actual situacion",
            "estrategias_ajustes": "Texto actual estrategias",
            "asistentes": [
                {"nombre": "Actual Nombre", "cargo": "Actual Cargo"},
                {"nombre": "", "cargo": ""},
            ],
        }
        previous_payload = {
            "modalidad": "Presencial",
            "fecha_seguimiento": "2026-04-15",
            "tipo_apoyo": seguimientos.TIPO_APOYO_OPTIONS[0],
            "item_labels": ["Asistencia", "Puntualidad"],
            "item_observaciones": ["Obs previa 1", "Obs previa 2"],
            "item_autoevaluacion": ["Bien", "Excelente"],
            "item_eval_empresa": ["Excelente", "Bien"],
            "empresa_item_labels": ["Comunicacion", "Productividad"],
            "empresa_eval": ["Bien", "Necesita mejorar"],
            "empresa_observacion": ["Emp previa 1", "Emp previa 2"],
            "situacion_encontrada": "Situacion previa",
            "estrategias_ajustes": "Estrategia previa",
            "asistentes": [
                {"nombre": "Profesional RECA", "cargo": "Profesional"},
                {"nombre": "Jefe directo", "cargo": "Coordinador"},
            ],
        }

        patchers = [
            patch.object(app, "attach_dictation", return_value=None),
            patch.object(
                app,
                "_get_asistentes_profesionales_catalog",
                return_value={"nombres": [], "cargos": [], "name_to_cargo": {}},
            ),
            patch.object(
                app.seguimientos,
                "get_case_meta",
                return_value={"max_seguimientos": 3, "base_sheet_name": base_sheet},
            ),
            patch.object(app.seguimientos, "get_workflow_state", return_value=workflow),
            patch.object(
                app.seguimientos,
                "suggest_next_step",
                return_value={"sheet": followup_2, "message": workflow["message"]},
            ),
            patch.object(
                app.seguimientos,
                "get_followup_payload",
                side_effect=lambda _case_ref, index: previous_payload if index == 1 else current_payload,
            ),
            patch.object(app.seguimientos, "sync_case_record_from_local", return_value=None),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        editor = app.SeguimientoEditorWindow(self.root, case_path="", case_record=case_record)
        self.addCleanup(destroy_widget, editor)
        editor.sheet_var.set(followup_2)
        editor._render_selected_sheet()

        editor._copy_previous_followup_values()

        self.assertEqual(editor.follow_vars["modalidad"].get(), "Presencial")
        self.assertEqual(editor.follow_vars["tipo_apoyo"].get(), seguimientos.TIPO_APOYO_OPTIONS[0])
        self.assertEqual(editor.follow_vars["fecha_seguimiento"].get(), "2026-05-10")
        self.assertEqual(editor.follow_item_obs[0].get(), "Obs previa 1")
        self.assertEqual(editor.follow_item_auto[0].get(), "Bien")
        self.assertEqual(editor.follow_item_emp[0].get(), "Excelente")
        self.assertEqual(editor.follow_emp_eval[1].get(), "Necesita mejorar")
        self.assertEqual(editor.follow_emp_obs[1].get(), "Emp previa 2")
        self.assertEqual(editor.follow_text["situacion_encontrada"].get("1.0", tk.END).strip(), "Texto actual situacion")
        self.assertEqual(editor.follow_text["estrategias_ajustes"].get("1.0", tk.END).strip(), "Texto actual estrategias")
        self.assertEqual(app._get_input_value(editor.follow_asistentes[0][0]), "Profesional RECA")
        self.assertEqual(app._get_input_value(editor.follow_asistentes[0][1]), "Profesional")

    def test_editor_save_confirms_before_overwriting_existing_values(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-overwrite-confirm",
            "webViewLink": "https://example.com/sheet-overwrite-confirm",
        }
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, f"{seguimientos.SHEET_PREFIX}1", seguimientos.SHEET_FINAL],
            "editable_sheet": base_sheet,
            "suggested_sheet": base_sheet,
            "next_followup": 1,
            "message": "Empieza por la ficha inicial del proceso.",
        }
        base_payload = {
            "fecha_visita": "2026-03-26",
            "modalidad": "Presencial",
            "nombre_empresa": "Empresa Demo",
            "cedula": "10101010",
            "fecha_fin_contrato": "2026-12-31",
            "seguimiento_fechas_1_3": ["", "", ""],
            "seguimiento_fechas_4_6": ["", "", ""],
        }
        save_base_payload = Mock()

        editor = self._make_editor(
            workflow=workflow,
            case_record=case_record,
            base_payload=base_payload,
            save_base_payload=save_base_payload,
        )
        editor._run_loading_job = self._run_loading_job_immediately

        field = editor.base_date_widgets["fecha_fin_contrato"]
        field.entry.delete(0, tk.END)
        field.entry.insert(0, "No aplica")
        editor._refresh_overwrite_highlights()

        self.assertEqual(field.entry.cget("bg"), app.COLOR_FIELD_WARNING_BG)

        with patch.object(app.messagebox, "askyesno", side_effect=[False, True]) as askyesno:
            editor._save_current_sheet()
            save_base_payload.assert_not_called()
            self.assertIn("cancelado", editor.status_var.get().lower())

            editor._save_current_sheet()

        self.assertEqual(askyesno.call_count, 2)
        save_base_payload.assert_called_once_with(editor.case_target, ANY)

    def test_followup_quick_actions_apply_values_to_each_group(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        followup_sheet = f"{seguimientos.SHEET_PREFIX}1"
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-quick-actions",
            "webViewLink": "https://example.com/sheet-quick-actions",
        }
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, followup_sheet, seguimientos.SHEET_FINAL],
            "editable_sheet": followup_sheet,
            "suggested_sheet": followup_sheet,
            "next_followup": 1,
            "message": "Seguimiento 1 habilitado.",
        }
        followup_payload = {
            "modalidad": "Presencial",
            "tipo_apoyo": seguimientos.TIPO_APOYO_OPTIONS[0],
            "item_labels": ["Asistencia", "Puntualidad"],
            "item_observaciones": ["", ""],
            "item_autoevaluacion": ["", ""],
            "item_eval_empresa": ["", ""],
            "empresa_item_labels": ["Comunicación", "Productividad"],
            "empresa_eval": ["", ""],
            "empresa_observacion": ["", ""],
            "situacion_encontrada": "",
            "estrategias_ajustes": "",
            "asistentes": [],
        }

        patchers = [
            patch.object(app, "attach_dictation", return_value=None),
            patch.object(
                app,
                "_get_asistentes_profesionales_catalog",
                return_value={"nombres": [], "cargos": [], "name_to_cargo": {}},
            ),
            patch.object(app.seguimientos, "get_case_meta", return_value={"max_seguimientos": 3, "base_sheet_name": base_sheet}),
            patch.object(app.seguimientos, "get_workflow_state", return_value=workflow),
            patch.object(app.seguimientos, "suggest_next_step", return_value={"sheet": followup_sheet, "message": workflow["message"]}),
            patch.object(app.seguimientos, "get_followup_payload", return_value=followup_payload),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        editor = app.SeguimientoEditorWindow(self.root, case_path="", case_record=case_record)
        self.addCleanup(destroy_widget, editor)

        editor._set_followup_eval_group_values("auto", "Bien")
        editor._set_followup_eval_group_values("item_empresa", "Excelente")
        editor._set_followup_eval_group_values("empresa_eval", "Necesita mejorar")

        self.assertEqual([var.get() for var in editor.follow_item_auto], ["Bien", "Bien"])
        self.assertEqual([var.get() for var in editor.follow_item_emp], ["Excelente", "Excelente"])
        self.assertEqual(
            [var.get() for var in editor.follow_emp_eval],
            ["Necesita mejorar", "Necesita mejorar"],
        )

    def test_editor_uses_inline_feedback_for_previous_followup_copy_without_modal(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        followup_sheet = f"{seguimientos.SHEET_PREFIX}1"
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, followup_sheet, seguimientos.SHEET_FINAL],
            "editable_sheet": followup_sheet,
            "suggested_sheet": followup_sheet,
            "next_followup": 1,
            "message": "Seguimiento 1 habilitado.",
        }
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-copy-inline",
            "webViewLink": "https://example.com/sheet-copy-inline",
        }
        followup_payload = {
            "modalidad": "Presencial",
            "tipo_apoyo": seguimientos.TIPO_APOYO_OPTIONS[0],
            "item_labels": ["Asistencia", "Puntualidad"],
            "item_observaciones": ["", ""],
            "item_autoevaluacion": ["", ""],
            "item_eval_empresa": ["", ""],
            "empresa_item_labels": ["Comunicación", "Productividad"],
            "empresa_eval": ["", ""],
            "empresa_observacion": ["", ""],
            "situacion_encontrada": "",
            "estrategias_ajustes": "",
            "asistentes": [],
        }

        editor = self._make_editor(
            workflow=workflow,
            case_record=case_record,
            followup_payload=followup_payload,
        )
        editor.sheet_var.set(followup_sheet)
        editor._render_selected_sheet()

        with patch.object(app.messagebox, "showinfo", return_value=None) as showinfo:
            with patch.object(app.messagebox, "showerror", return_value=None) as showerror:
                editor._copy_previous_followup_values()

        showinfo.assert_not_called()
        showerror.assert_not_called()
        self.assertIn("no tiene un seguimiento anterior", editor.status_var.get().lower())

    def test_editor_allows_saving_non_suggested_followup_without_blocking(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        followup_1 = f"{seguimientos.SHEET_PREFIX}1"
        followup_2 = f"{seguimientos.SHEET_PREFIX}2"
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, followup_1, followup_2, seguimientos.SHEET_FINAL],
            "editable_sheet": followup_2,
            "suggested_sheet": followup_2,
            "next_followup": 2,
            "message": "Seguimiento 2 habilitado.",
        }
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-closed-inline",
            "webViewLink": "https://example.com/sheet-closed-inline",
        }
        followup_payload = {
            "modalidad": "Presencial",
            "tipo_apoyo": seguimientos.TIPO_APOYO_OPTIONS[0],
            "item_labels": ["Asistencia", "Puntualidad"],
            "item_observaciones": ["", ""],
            "item_autoevaluacion": ["", ""],
            "item_eval_empresa": ["", ""],
            "empresa_item_labels": ["Comunicación", "Productividad"],
            "empresa_eval": ["", ""],
            "empresa_observacion": ["", ""],
            "situacion_encontrada": "",
            "estrategias_ajustes": "",
            "asistentes": [],
        }
        save_followup_payload = Mock()

        editor = self._make_editor(
            workflow=workflow,
            case_record=case_record,
            followup_payload=followup_payload,
            save_followup_payload=save_followup_payload,
        )
        editor._run_loading_job = self._run_loading_job_immediately
        editor.sheet_var.set(followup_1)
        editor._render_selected_sheet()
        self.assertIn("sugerida actual", editor.status_var.get().lower())

        with patch.object(app.messagebox, "showinfo", return_value=None) as showinfo:
            with patch.object(app.messagebox, "showwarning", return_value=None) as showwarning:
                editor._save_current_sheet()

        showinfo.assert_not_called()
        showwarning.assert_not_called()
        save_followup_payload.assert_called_once_with(editor.case_target, 1, ANY)
        self.assertIn("sugerida actual", editor.status_var.get().lower())

    def test_editor_uses_inline_feedback_for_final_sheet_save_attempt(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, f"{seguimientos.SHEET_PREFIX}1", seguimientos.SHEET_FINAL],
            "editable_sheet": f"{seguimientos.SHEET_PREFIX}1",
            "suggested_sheet": seguimientos.SHEET_FINAL,
            "next_followup": 1,
            "message": "Seguimiento 1 habilitado.",
        }
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-final-inline",
            "webViewLink": "https://example.com/sheet-final-inline",
        }

        editor = self._make_editor(workflow=workflow, case_record=case_record)
        editor.sheet_var.set(seguimientos.SHEET_FINAL)
        editor._render_selected_sheet()

        with patch.object(app.messagebox, "showinfo", return_value=None) as showinfo:
            with patch.object(app.messagebox, "showwarning", return_value=None) as showwarning:
                editor._save_current_sheet()

        showinfo.assert_not_called()
        showwarning.assert_not_called()
        self.assertIn("solo lectura", editor.status_var.get().lower())

    def test_editor_company_lookup_by_nit_uses_inline_warning_without_modal(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, f"{seguimientos.SHEET_PREFIX}1", seguimientos.SHEET_FINAL],
            "editable_sheet": base_sheet,
            "suggested_sheet": base_sheet,
            "next_followup": 1,
            "message": "Completa la hoja base.",
        }
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-company-inline",
            "webViewLink": "https://example.com/sheet-company-inline",
        }
        base_payload = {
            "fecha_visita": "2026-03-26",
            "modalidad": "Presencial",
            "nit_empresa": "900123456",
            "nombre_empresa": "",
        }

        editor = self._make_editor(
            workflow=workflow,
            case_record=case_record,
            base_payload=base_payload,
        )
        editor.base_vars["nit_empresa"].set("900123456")

        with patch.object(app.seguimientos, "get_empresa_by_nit", return_value=None):
            with patch.object(app.messagebox, "showwarning", return_value=None) as showwarning:
                with patch.object(app.messagebox, "showerror", return_value=None) as showerror:
                    editor._buscar_empresa_por_nit()

        showwarning.assert_not_called()
        showerror.assert_not_called()
        self.assertIn("no se encontr", editor.status_var.get().lower())
        self.assertIn("empresa", editor.status_var.get().lower())

    def test_editor_sheet_selector_uses_friendly_stage_titles_and_continue_button(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        followup_sheet = f"{seguimientos.SHEET_PREFIX}1"
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-stage-selector",
            "webViewLink": "https://example.com/sheet-stage-selector",
        }
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, followup_sheet, seguimientos.SHEET_FINAL],
            "editable_sheet": followup_sheet,
            "suggested_sheet": followup_sheet,
            "next_followup": 1,
            "message": "La ficha inicial esta completa. Continua con Seguimiento 1.",
            "stage_model": [
                {
                    "stage_id": "base_process",
                    "title": "Ficha inicial del proceso",
                    "label": "Ficha inicial del proceso",
                    "sheet_name": base_sheet,
                    "status": "completed",
                    "coverage_percent": 95,
                    "is_completed": True,
                    "is_suggested": False,
                    "is_editable": True,
                    "helper_text": "La ficha inicial esta completa.",
                },
                {
                    "stage_id": "followup_1",
                    "title": "Seguimiento 1",
                    "label": "Seguimiento 1",
                    "sheet_name": followup_sheet,
                    "status": "not_started",
                    "coverage_percent": 0,
                    "is_completed": False,
                    "is_suggested": True,
                    "is_editable": True,
                    "helper_text": "Esta es la etapa sugerida para continuar.",
                },
                {
                    "stage_id": "final_result",
                    "title": "Resultado final",
                    "label": "Resultado final",
                    "sheet_name": seguimientos.SHEET_FINAL,
                    "status": "review_only",
                    "coverage_percent": 0,
                    "is_completed": False,
                    "is_suggested": False,
                    "is_editable": False,
                    "helper_text": "Consolidado automatico del proceso.",
                },
            ],
        }
        base_payload = {
            "fecha_visita": "2026-03-26",
            "modalidad": "Presencial",
            "nombre_empresa": "Empresa Demo",
            "cedula": "10101010",
            "seguimiento_fechas_1_3": ["2026-04-02", "", ""],
            "seguimiento_fechas_4_6": ["", "", ""],
        }
        followup_payload = {
            "modalidad": "Mixta",
            "tipo_apoyo": "",
            "item_labels": ["Asistencia"],
            "item_observaciones": [""],
            "item_autoevaluacion": [""],
            "item_eval_empresa": [""],
            "empresa_item_labels": ["Comunicacion"],
            "empresa_eval": [""],
            "empresa_observacion": [""],
            "situacion_encontrada": "",
            "estrategias_ajustes": "",
            "asistentes": [],
        }

        editor = self._make_editor(
            workflow=workflow,
            case_record=case_record,
            base_payload=base_payload,
            followup_payload=followup_payload,
        )
        editor.sheet_var.set(base_sheet)
        editor._render_selected_sheet()

        self.assertEqual(
            tuple(editor.sheet_combo.cget("values") or ()),
            ("Ficha inicial del proceso", "Seguimiento 1", "Resultado final"),
        )
        self.assertEqual(editor.sheet_display_var.get(), "Ficha inicial del proceso")
        self.assertEqual(editor.continue_stage_button.cget("text"), "Continuar a Seguimiento 1")
        self.assertEqual(str(editor.continue_stage_button.winfo_manager()), "pack")
        self.assertIn("etapa sugerida actual: seguimiento 1", editor.status_var.get().lower())

    def test_editor_shows_inline_transition_after_base_reaches_followup_1(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        followup_sheet = f"{seguimientos.SHEET_PREFIX}1"
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-base-transition",
            "webViewLink": "https://example.com/sheet-base-transition",
        }
        initial_workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, followup_sheet, seguimientos.SHEET_FINAL],
            "editable_sheet": base_sheet,
            "suggested_sheet": base_sheet,
            "next_followup": 1,
            "base_completed": False,
            "message": "Empieza por la ficha inicial del proceso.",
            "stage_model": [
                {
                    "stage_id": "base_process",
                    "title": "Ficha inicial del proceso",
                    "label": "Ficha inicial del proceso",
                    "sheet_name": base_sheet,
                    "status": "in_progress",
                    "coverage_percent": 80,
                    "is_completed": False,
                    "is_suggested": True,
                    "is_editable": True,
                    "helper_text": "Completa la informacion inicial del proceso.",
                },
                {
                    "stage_id": "followup_1",
                    "title": "Seguimiento 1",
                    "label": "Seguimiento 1",
                    "sheet_name": followup_sheet,
                    "status": "not_started",
                    "coverage_percent": 0,
                    "is_completed": False,
                    "is_suggested": False,
                    "is_editable": True,
                    "helper_text": "Seguimiento 1 pendiente por diligenciar.",
                },
                {
                    "stage_id": "final_result",
                    "title": "Resultado final",
                    "label": "Resultado final",
                    "sheet_name": seguimientos.SHEET_FINAL,
                    "status": "review_only",
                    "coverage_percent": 0,
                    "is_completed": False,
                    "is_suggested": False,
                    "is_editable": False,
                    "helper_text": "Consolidado automatico del proceso.",
                },
            ],
        }
        updated_workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, followup_sheet, seguimientos.SHEET_FINAL],
            "editable_sheet": followup_sheet,
            "suggested_sheet": followup_sheet,
            "next_followup": 1,
            "base_completed": True,
            "message": "La ficha inicial esta completa. Continua con Seguimiento 1.",
            "stage_model": [
                {
                    "stage_id": "base_process",
                    "title": "Ficha inicial del proceso",
                    "label": "Ficha inicial del proceso",
                    "sheet_name": base_sheet,
                    "status": "completed",
                    "coverage_percent": 95,
                    "is_completed": True,
                    "is_suggested": False,
                    "is_editable": True,
                    "helper_text": "La ficha inicial esta completa.",
                },
                {
                    "stage_id": "followup_1",
                    "title": "Seguimiento 1",
                    "label": "Seguimiento 1",
                    "sheet_name": followup_sheet,
                    "status": "not_started",
                    "coverage_percent": 0,
                    "is_completed": False,
                    "is_suggested": True,
                    "is_editable": True,
                    "helper_text": "Esta es la etapa sugerida para continuar.",
                },
                {
                    "stage_id": "final_result",
                    "title": "Resultado final",
                    "label": "Resultado final",
                    "sheet_name": seguimientos.SHEET_FINAL,
                    "status": "review_only",
                    "coverage_percent": 0,
                    "is_completed": False,
                    "is_suggested": False,
                    "is_editable": False,
                    "helper_text": "Consolidado automatico del proceso.",
                },
            ],
        }
        base_payload = {
            "fecha_visita": "2026-04-07",
            "modalidad": "Presencial",
            "nombre_empresa": "Empresa Demo",
            "cedula": "10101010",
            "contacto_emergencia": "Contacto",
            "parentesco": "Hermana",
            "telefono_emergencia": "3111111111",
            "certificado_discapacidad": "Si",
            "certificado_porcentaje": "25%",
            "tipo_contrato": "Indefinido",
            "fecha_firma_contrato": "2026-03-01",
            "fecha_inicio_contrato": "2026-03-05",
            "fecha_fin_contrato": "",
            "apoyos_ajustes": "Ajuste razonable",
            "funciones_1_5": ["F1", "F2", "F3", "F4", "F5"],
            "funciones_6_10": ["F6", "F7", "F8", "F9", "F10"],
            "seguimiento_fechas_1_3": ["2026-04-20", "", ""],
            "seguimiento_fechas_4_6": ["", "", ""],
        }
        save_base_payload = Mock()

        patchers = [
            patch.object(app, "attach_dictation", return_value=None),
            patch.object(
                app,
                "_get_asistentes_profesionales_catalog",
                return_value={"nombres": [], "cargos": [], "name_to_cargo": {}},
            ),
            patch.object(
                app.seguimientos,
                "get_case_meta",
                return_value={"max_seguimientos": 3, "base_sheet_name": base_sheet},
            ),
            patch.object(
                app.seguimientos,
                "get_workflow_state",
                side_effect=[initial_workflow, updated_workflow],
            ),
            patch.object(
                app.seguimientos,
                "suggest_next_step",
                side_effect=[
                    {"sheet": base_sheet, "message": initial_workflow["message"]},
                    {"sheet": followup_sheet, "message": updated_workflow["message"]},
                ],
            ),
            patch.object(app.seguimientos, "get_base_payload", return_value=base_payload),
            patch.object(app.seguimientos, "save_base_payload", save_base_payload),
            patch.object(app.seguimientos, "sync_case_record_from_local", return_value=None),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        editor = app.SeguimientoEditorWindow(self.root, case_path="", case_record=case_record)
        self.addCleanup(destroy_widget, editor)
        editor._run_loading_job = self._run_loading_job_immediately

        editor._save_current_sheet()

        save_base_payload.assert_called_once_with(editor.case_target, ANY)
        self.assertIn("ficha inicial completa", editor.status_var.get().lower())
        self.assertIn("seguimiento 1", editor.status_var.get().lower())
        self.assertEqual(editor.continue_stage_button.cget("text"), "Continuar a Seguimiento 1")

    def test_editor_local_autosave_persists_base_sheet_draft_without_remote_save(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-autosave-base",
            "webViewLink": "https://example.com/sheet-autosave-base",
        }
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, f"{seguimientos.SHEET_PREFIX}1", seguimientos.SHEET_FINAL],
            "editable_sheet": base_sheet,
            "suggested_sheet": base_sheet,
            "next_followup": 1,
            "message": "Completa la hoja base.",
        }
        base_payload = {
            "fecha_visita": "",
            "modalidad": "",
            "nombre_empresa": "Empresa Demo",
            "cedula": "10101010",
        }
        save_base_payload = Mock()
        self._install_followup_local_draft_store()

        editor = self._make_editor(
            workflow=workflow,
            case_record=case_record,
            base_payload=base_payload,
            save_base_payload=save_base_payload,
        )

        editor.base_vars["fecha_visita"].set("2026-04-07")
        editor._schedule_sheet_autosave(delay_ms=0)
        editor._flush_pending_sheet_autosave()

        save_base_payload.assert_not_called()
        local_draft = app._get_followup_local_sheet_draft(editor.case_target, base_sheet)
        self.assertIsNotNone(local_draft)
        self.assertEqual(local_draft["payload"]["fecha_visita"], "2026-04-07")
        self.assertEqual(local_draft["payload"]["modalidad"], "")
        self.assertIn("borrador local", editor.status_var.get().lower())

    def test_editor_local_autosave_restores_base_draft_after_sheet_switch(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        followup_sheet = f"{seguimientos.SHEET_PREFIX}1"
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-autosave-switch",
            "webViewLink": "https://example.com/sheet-autosave-switch",
        }
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, followup_sheet, seguimientos.SHEET_FINAL],
            "editable_sheet": base_sheet,
            "suggested_sheet": base_sheet,
            "next_followup": 1,
            "message": "Completa la hoja base.",
        }
        base_payload = {
            "fecha_visita": "2026-04-01",
            "modalidad": "Presencial",
            "nombre_empresa": "Empresa Demo",
            "cedula": "10101010",
        }
        followup_payload = {
            "modalidad": "",
            "tipo_apoyo": "",
            "item_labels": ["Asistencia"],
            "item_observaciones": [""],
            "item_autoevaluacion": [""],
            "item_eval_empresa": [""],
            "empresa_item_labels": ["ComunicaciÃ³n"],
            "empresa_eval": [""],
            "empresa_observacion": [""],
            "situacion_encontrada": "",
            "estrategias_ajustes": "",
            "asistentes": [],
        }
        save_base_payload = Mock()
        self._install_followup_local_draft_store()

        editor = self._make_editor(
            workflow=workflow,
            case_record=case_record,
            base_payload=base_payload,
            followup_payload=followup_payload,
            save_base_payload=save_base_payload,
        )

        editor.base_vars["modalidad"].set("Mixta")
        editor.sheet_var.set(followup_sheet)
        editor._render_selected_sheet()

        save_base_payload.assert_not_called()
        local_draft = app._get_followup_local_sheet_draft(editor.case_target, base_sheet)
        self.assertIsNotNone(local_draft)
        self.assertEqual(local_draft["payload"]["modalidad"], "Mixta")
        self.assertEqual(editor.current_followup_index, 1)
        editor.sheet_var.set(base_sheet)
        editor._render_selected_sheet()
        self.assertEqual(editor.base_vars["modalidad"].get(), "Mixta")

    def test_editor_local_autosave_keeps_followup_only_in_local_draft_store(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        followup_sheet = f"{seguimientos.SHEET_PREFIX}1"
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-autosave-followup",
            "webViewLink": "https://example.com/sheet-autosave-followup",
        }
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, followup_sheet, seguimientos.SHEET_FINAL],
            "editable_sheet": followup_sheet,
            "suggested_sheet": followup_sheet,
            "next_followup": 1,
            "message": "Seguimiento 1 habilitado.",
        }
        followup_payload = {
            "modalidad": "Presencial",
            "tipo_apoyo": "",
            "item_labels": ["Asistencia"],
            "item_observaciones": [""],
            "item_autoevaluacion": [""],
            "item_eval_empresa": [""],
            "empresa_item_labels": ["ComunicaciÃ³n"],
            "empresa_eval": [""],
            "empresa_observacion": [""],
            "situacion_encontrada": "",
            "estrategias_ajustes": "",
            "asistentes": [],
        }
        save_followup_payload = Mock()
        self._install_followup_local_draft_store()

        editor = self._make_editor(
            workflow=workflow,
            case_record=case_record,
            followup_payload=followup_payload,
            save_followup_payload=save_followup_payload,
        )
        hub = Mock()
        editor._get_hub_window = Mock(return_value=hub)
        editor.sheet_var.set(followup_sheet)
        editor._render_selected_sheet()

        editor.follow_vars["tipo_apoyo"].set(seguimientos.TIPO_APOYO_OPTIONS[0])
        editor._schedule_sheet_autosave(delay_ms=0)
        editor._flush_pending_sheet_autosave()

        save_followup_payload.assert_not_called()
        hub.record_followup_completion.assert_not_called()
        local_draft = app._get_followup_local_sheet_draft(editor.case_target, followup_sheet)
        self.assertIsNotNone(local_draft)
        self.assertEqual(local_draft["payload"]["tipo_apoyo"], seguimientos.TIPO_APOYO_OPTIONS[0])
        self.assertIn("borrador local", editor.status_var.get().lower())

    def test_editor_close_flushes_pending_local_draft(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-autosave-close",
            "webViewLink": "https://example.com/sheet-autosave-close",
        }
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, f"{seguimientos.SHEET_PREFIX}1", seguimientos.SHEET_FINAL],
            "editable_sheet": base_sheet,
            "suggested_sheet": base_sheet,
            "next_followup": 1,
            "message": "Completa la hoja base.",
        }
        base_payload = {
            "fecha_visita": "2026-04-01",
            "modalidad": "Presencial",
            "nombre_empresa": "Empresa Demo",
            "cedula": "10101010",
        }
        self._install_followup_local_draft_store()

        editor = self._make_editor(
            workflow=workflow,
            case_record=case_record,
            base_payload=base_payload,
        )

        editor.base_vars["modalidad"].set("Mixta")
        editor._close_editor()

        local_draft = app._get_followup_local_sheet_draft(case_record, base_sheet)
        self.assertIsNotNone(local_draft)
        self.assertEqual(local_draft["payload"]["modalidad"], "Mixta")

    def test_hub_user_drafts_include_followup_local_drafts(self) -> None:
        self._install_empty_hub_draft_store()
        self._install_followup_local_draft_store()
        case_record = {
            "source": "drive",
            "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
            "file_id": "sheet-hub-draft",
            "webViewLink": "https://example.com/sheet-hub-draft",
        }
        request = {
            "sheet": seguimientos.SHEET_BASE,
            "save_kind": "base",
            "payload": {"nombre_empresa": "Empresa Demo", "fecha_visita": "2026-04-07"},
            "followup_index": None,
            "fingerprint": "fp-hub-draft",
        }
        app._save_followup_local_sheet_draft(
            case_record,
            request,
            metadata={
                "user_login": "demo",
                "company_name": "Empresa Demo",
                "case_record": case_record,
                "case_path": "",
                "case_label": "Empresa Demo",
            },
        )
        hub = object.__new__(app.HubWindow)
        hub.current_user_profile = {"usuario_login": "demo"}
        hub.current_user = "demo"

        drafts = app.HubWindow._get_user_drafts(hub)

        self.assertTrue(any(str(item.get("draft_type") or "") == "followup_local" for item in drafts))
        followup_draft = next(item for item in drafts if str(item.get("draft_type") or "") == "followup_local")
        self.assertEqual(followup_draft["form_id"], "seguimientos")
        self.assertEqual(followup_draft["company_name"], "Empresa Demo")
        self.assertEqual(followup_draft["sheet_name"], seguimientos.SHEET_BASE)

    def test_hub_open_draft_entry_opens_followup_editor_for_local_draft(self) -> None:
        draft = {
            "draft_id": "seguimientos:test",
            "draft_type": "followup_local",
            "sheet_name": f"{seguimientos.SHEET_PREFIX}1",
            "case_record": {
                "source": "drive",
                "mime_type": seguimientos.GOOGLE_SHEETS_MIME,
                "file_id": "sheet-open-draft",
                "webViewLink": "https://example.com/sheet-open-draft",
            },
            "case_path": "",
        }
        hub = object.__new__(app.HubWindow)
        hub.track_form_open = Mock()

        bootstrap = {
            "meta": {"max_seguimientos": 3, "base_sheet_name": seguimientos.SHEET_BASE},
            "workflow": {
                "max_seguimientos": 3,
                "base_sheet_name": seguimientos.SHEET_BASE,
                "visible_sheets": [seguimientos.SHEET_BASE, f"{seguimientos.SHEET_PREFIX}1", seguimientos.SHEET_FINAL],
                "suggested_sheet": seguimientos.SHEET_BASE,
            },
            "suggestion": {
                "sheet": seguimientos.SHEET_BASE,
                "message": "Empieza por la ficha inicial del proceso.",
                "max_seguimientos": 3,
            },
        }

        with patch.object(app, "_load_followup_editor_bootstrap", return_value=bootstrap):
            with patch.object(app, "SeguimientoEditorWindow", return_value=Mock()) as editor_ctor:
                with patch.object(app, "_focus_window", return_value=None) as focus_window:
                    app.HubWindow._open_draft_entry(hub, draft)

        editor_ctor.assert_called_once()
        _, kwargs = editor_ctor.call_args
        self.assertEqual(kwargs["case_record"]["file_id"], "sheet-open-draft")
        self.assertEqual(kwargs["bootstrap"]["suggestion"]["sheet"], f"{seguimientos.SHEET_PREFIX}1")
        focus_window.assert_called_once()
        hub.track_form_open.assert_called_once_with("seguimientos", "Seguimientos")


if __name__ == "__main__":
    unittest.main()
