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


if __name__ == "__main__":
    unittest.main()
