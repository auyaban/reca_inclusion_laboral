from __future__ import annotations

import tkinter as tk
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
            "message": "Hoja base habilitada.",
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
            patch.object(app.seguimientos, "suggest_next_step", return_value={"sheet": base_sheet, "message": "Editar hoja base"}),
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
        self.assertEqual(editor.status_var.get(), "Ponderado final es de solo revisión.")

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
            "message": "Hoja base y seguimiento 1 habilitados hasta diligenciar el seguimiento 1.",
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
            patch.object(app.seguimientos, "suggest_next_step", return_value={"sheet": base_sheet, "message": "Editar hoja base"}),
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
        self.assertEqual(editor.base_dates_1[0].get().strip(), "")
        self.assertEqual(editor.base_dates_1[1].get().strip(), "")
        self.assertEqual(editor.base_dates_2[0].get().strip(), "")
        self.assertEqual(str(editor.base_dates_1[0].select_button.cget("state")), "normal")
        self.assertEqual(str(editor.base_dates_1[1].select_button.cget("state")), "disabled")
        self.assertEqual(str(editor.base_dates_2[0].select_button.cget("state")), "disabled")

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
        self.assertEqual(str(editor.base_dates_1[1].select_button.cget("state")), "normal")
        self.assertEqual(str(editor.base_dates_1[2].select_button.cget("state")), "disabled")

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

    def test_editor_uses_inline_feedback_for_closed_sheet_save_attempt(self) -> None:
        base_sheet = seguimientos.SHEET_BASE
        followup_1 = f"{seguimientos.SHEET_PREFIX}1"
        followup_2 = f"{seguimientos.SHEET_PREFIX}2"
        workflow = {
            "max_seguimientos": 3,
            "base_sheet_name": base_sheet,
            "visible_sheets": [base_sheet, followup_1, followup_2, seguimientos.SHEET_FINAL],
            "editable_sheet": followup_2,
            "suggested_sheet": followup_1,
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

        editor = self._make_editor(
            workflow=workflow,
            case_record=case_record,
            followup_payload=followup_payload,
        )
        editor.sheet_var.set(followup_1)
        editor._render_selected_sheet()

        with patch.object(app.messagebox, "showinfo", return_value=None) as showinfo:
            with patch.object(app.messagebox, "showwarning", return_value=None) as showwarning:
                editor._save_current_sheet()

        showinfo.assert_not_called()
        showwarning.assert_not_called()
        self.assertIn("solo se puede editar", editor.status_var.get().lower())

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
        self.assertIn("no se diligencia manualmente", editor.status_var.get().lower())

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


if __name__ == "__main__":
    unittest.main()
