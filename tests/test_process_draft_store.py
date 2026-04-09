from __future__ import annotations

import copy
import os
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

import app


class _ModuleStub:
    def get_form_cache(self):
        return {}

    def save_cache_to_file(self):
        return None

    def clear_form_cache(self):
        return None


class _HubStub:
    current_user_profile = {"usuario_login": "tester@example.com"}
    current_user = ""
    _drafts_btn = None

    def __init__(self, *, cache_snapshot, ui_section):
        self._draft_capture = (cache_snapshot, [], ui_section)
        self.badge_refreshes = 0
        self.toasts = []

    def _get_current_user_login(self):
        return "tester@example.com"

    def _capture_window_draft_state(self, _window, _module):
        cache_snapshot, ui_snapshot, ui_section = self._draft_capture
        return copy.deepcopy(cache_snapshot), copy.deepcopy(ui_snapshot), ui_section

    def _draft_state_has_content(self, cache_snapshot, ui_snapshot):
        return app.HubWindow._draft_state_has_content(self, cache_snapshot, ui_snapshot)

    def _refresh_drafts_badge(self):
        self.badge_refreshes += 1

    def show_toast(self, text):
        self.toasts.append(text)


def _make_window(form_id, form_name, *, session_key):
    return SimpleNamespace(
        _form_id=form_id,
        _form_name=form_name,
        _draft_id="",
        _draft_last_fingerprint="",
        _draft_session_key=session_key,
        _draft_autosave_after_id=None,
    )


class ProcessDraftStoreTests(unittest.TestCase):
    def setUp(self) -> None:
        super().setUp()
        self.tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(self.tempdir.cleanup)
        self.drafts_path = os.path.join(self.tempdir.name, "form_drafts_il.json")
        self.process_drafts_dir = os.path.join(self.tempdir.name, "form_drafts")

        patchers = [
            patch.object(app, "_get_drafts_path", return_value=self.drafts_path),
            patch.object(app, "_get_form_drafts_dir", return_value=self.process_drafts_dir),
            patch.object(app, "_update_window_save_status", return_value=None),
            patch.object(app, "_resolve_form_meta", side_effect=lambda form_id: {"id": form_id, "name": str(form_id), "supports_drafts": True}),
            patch.object(app.messagebox, "showerror", return_value=None),
            patch.object(app.messagebox, "showinfo", return_value=None),
            patch.object(app.messagebox, "showwarning", return_value=None),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        self.form_modules = {
            "seleccion_incluyente": _ModuleStub(),
            "contratacion_incluyente": _ModuleStub(),
            "condiciones_vacante": _ModuleStub(),
        }
        self.form_module_patcher = patch.dict(app.FORM_MODULE_MAP, self.form_modules, clear=False)
        self.form_module_patcher.start()
        self.addCleanup(self.form_module_patcher.stop)

    def _section1_cache(self, *, company_name="Empresa Demo", nit="900123456"):
        return {
            "section_1": {
                "fecha_visita": "2026-04-09",
                "modalidad": "Presencial",
                "nit_empresa": nit,
                "nombre_empresa": company_name,
                "direccion_empresa": "Calle 1",
                "correo_1": "empresa@example.com",
                "contacto_empresa": "Ana",
                "telefono_empresa": "1234567",
                "cargo": "Lider",
                "ciudad_empresa": "Bogota",
                "sede_empresa": "Centro",
                "caja_compensacion": "Compensar",
                "asesor": "Carlos",
                "profesional_asignado": "Laura",
            }
        }

    def test_persist_form_draft_keeps_multiple_processes_for_same_company(self) -> None:
        for form_id, form_name in (
            ("seleccion_incluyente", "Seleccion Incluyente"),
            ("contratacion_incluyente", "Contratacion Incluyente"),
            ("condiciones_vacante", "Condiciones de Vacante"),
        ):
            with self.subTest(form_id=form_id):
                if os.path.exists(self.drafts_path):
                    os.remove(self.drafts_path)
                if os.path.isdir(self.process_drafts_dir):
                    for name in os.listdir(self.process_drafts_dir):
                        os.remove(os.path.join(self.process_drafts_dir, name))

                created_ids = []
                for index in range(3):
                    hub = _HubStub(cache_snapshot=self._section1_cache(), ui_section="section_1")
                    window = _make_window(form_id, form_name, session_key=f"{form_id}-session-{index}")
                    result = app.HubWindow._persist_form_draft(
                        hub,
                        window,
                        allow_empty=False,
                        silent=True,
                        source="manual",
                    )
                    self.assertTrue(result)
                    created_ids.append(window._draft_id)

                store = app._load_drafts_store()
                drafts = store["users"]["tester@example.com"]
                self.assertEqual(len(drafts), 3)
                self.assertEqual(len({row["draft_id"] for row in drafts}), 3)
                self.assertEqual(set(created_ids), {row["draft_id"] for row in drafts})
                self.assertEqual(len(os.listdir(self.process_drafts_dir)), 3)
                for row in drafts:
                    self.assertEqual(row["company_name"], "Empresa Demo")
                    self.assertEqual(row["draft_format_version"], app.FORM_PROCESS_DRAFT_VERSION)
                    self.assertTrue(os.path.exists(row["draft_path"]))

    def test_delete_window_draft_removes_only_current_process(self) -> None:
        created = []
        for index in range(2):
            hub = _HubStub(cache_snapshot=self._section1_cache(), ui_section="section_1")
            window = _make_window("seleccion_incluyente", "Seleccion Incluyente", session_key=f"session-{index}")
            result = app.HubWindow._persist_form_draft(
                hub,
                window,
                allow_empty=False,
                silent=True,
                source="manual",
            )
            self.assertTrue(result)
            created.append((window._draft_id, app._load_drafts_store()["users"]["tester@example.com"][index]["draft_path"]))

        delete_hub = SimpleNamespace(
            current_user_profile={"usuario_login": "tester@example.com"},
            current_user="",
            _drafts_btn=None,
            _get_current_user_login=lambda: "tester@example.com",
            _refresh_drafts_badge=lambda: None,
        )
        window = SimpleNamespace(_draft_id=created[0][0], _draft_last_fingerprint="")

        deleted = app.HubWindow._delete_window_draft(delete_hub, window)

        self.assertTrue(deleted)
        store = app._load_drafts_store()
        drafts = store["users"]["tester@example.com"]
        self.assertEqual(len(drafts), 1)
        self.assertEqual(drafts[0]["draft_id"], created[1][0])
        self.assertFalse(os.path.exists(created[0][1]))
        self.assertTrue(os.path.exists(created[1][1]))

    def test_confirm_section1_and_continue_creates_process_draft(self) -> None:
        hub = SimpleNamespace(_persist_form_draft=Mock(return_value=True))
        window = SimpleNamespace(
            company_data={"nombre_empresa": "Empresa Demo", "nit_empresa": "900123456"},
            master=hub,
        )
        confirm_fn = Mock()
        next_step = Mock()

        with patch.object(app, "ensure_process_draft", return_value=True) as ensure_mock:
            with patch.object(app.ui_feedback, "clear_field_errors", return_value=None):
                with patch.object(app, "_clear_inline_feedback", return_value=None):
                    with patch.object(app, "_get_required_fecha_visita", return_value="2026-04-09"):
                        with patch.object(app, "_get_required_modalidad", return_value="Presencial"):
                            with patch.object(app, "_section1_search_value", return_value="900123456"):
                                app._confirm_section1_and_continue(
                                    window,
                                    confirm_fn=confirm_fn,
                                    next_step=next_step,
                                )

        confirm_fn.assert_called_once()
        ensure_mock.assert_called_once_with(window, silent=True)
        next_step.assert_called_once()


if __name__ == "__main__":
    unittest.main()
