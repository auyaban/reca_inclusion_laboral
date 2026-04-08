import unittest
from unittest.mock import patch

import app
from tests.tk_test_utils import TkTestCase, destroy_widget


class _HubRuntimeStub:
    _empresa_names_cache = []

    def _install_form_autosave_bindings(self, _window):
        return None

    def _schedule_window_draft_autosave(self, _window, delay_ms=250):
        return None


class CondicionesVacanteDraftRestoreTests(TkTestCase):
    def setUp(self) -> None:
        super().setUp()
        patchers = [
            patch.object(app, "_maximize_window", lambda _window: None),
            patch.object(app.messagebox, "showerror", return_value=None),
            patch.object(app.messagebox, "showinfo", return_value=None),
            patch.object(app.messagebox, "showwarning", return_value=None),
            patch.object(
                app,
                "_get_asistentes_profesionales_catalog",
                return_value={
                    "nombres": ["Andres Eduardo Montes Agudelo"],
                    "cargos": ["Gestor inclusión laboral"],
                    "name_to_cargo": {"andres eduardo montes agudelo": "Gestor inclusión laboral"},
                },
            ),
            patch.object(
                app,
                "_get_asesores_agencia_catalog",
                return_value={
                    "nombres": ["Asesor Agencia Demo"],
                    "cargos": ["Asesor Agencia"],
                    "name_to_cargo": {"asesor agencia demo": "Asesor Agencia"},
                },
            ),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)
        app.condiciones_vacante.clear_form_cache()
        self.addCleanup(app.condiciones_vacante.clear_form_cache)

    def test_bind_form_runtime_replays_pending_draft_section_with_wrapped_route(self) -> None:
        self.root._pending_draft_restore = {
            "form_id": "condiciones_vacante",
            "cache": {
                "_last_section": "section_8",
                "section_8": [
                    {
                        "nombre": "Andres Eduardo Montes Agudelo",
                        "cargo": "Gestor inclusión laboral",
                    }
                ],
            },
            "ui_section": "section_8",
            "ui_snapshot": [],
        }

        window = app.CondicionesVacanteWindow(self.root)
        self.addCleanup(destroy_widget, window)

        hub = _HubRuntimeStub()
        with patch.object(app, "_attach_dictation_for_section", return_value=None):
            app.HubWindow._bind_form_runtime(hub, window, app.condiciones_vacante.register_form())

        window.update_idletasks()

        self.assertEqual(window._current_section, "section_8")
        self.assertEqual(window.header_title.cget("text"), "8. ASISTENTES")
        self.assertEqual(window.header_progress_label.cget("text"), "Sección 9 de 9")
        self.assertTrue(window.section8_add_btn.winfo_exists())
        self.assertGreaterEqual(len(window.section8_rows), 3)
        self.assertEqual(app._get_input_value(window.section8_rows[0][0]), "Andres Eduardo Montes Agudelo")
        self.assertEqual(app._get_input_value(window.section8_rows[0][1]), "Gestor inclusión laboral")


if __name__ == "__main__":
    unittest.main()
