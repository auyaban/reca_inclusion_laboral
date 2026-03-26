from __future__ import annotations

import tkinter as tk
import unittest
from unittest.mock import patch

import app


class _TkTestCase(unittest.TestCase):
    def setUp(self) -> None:
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk no disponible: {exc}")
        self.root.withdraw()
        self.addCleanup(self._cleanup_root)

        patchers = [
            patch.object(app, "_maximize_window", lambda _window: None),
            patch.object(app.messagebox, "showerror", return_value=None),
            patch.object(app.messagebox, "showinfo", return_value=None),
            patch.object(app.messagebox, "showwarning", return_value=None),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

    def _cleanup_root(self) -> None:
        try:
            for child in self.root.winfo_children():
                child.destroy()
            self.root.destroy()
        except Exception:
            pass


class EvaluacionAccesibilidadWindowTests(_TkTestCase):
    def setUp(self) -> None:
        super().setUp()
        app.evaluacion_accesibilidad.clear_form_cache()
        self.addCleanup(app.evaluacion_accesibilidad.clear_form_cache)
        patchers = [
            patch.object(app, "_consume_pending_draft_restore", return_value=False),
            patch.object(app.evaluacion_accesibilidad, "cache_file_exists", return_value=False),
            patch.object(app.evaluacion_accesibilidad, "save_cache_to_file", return_value=None),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

    def _create_window(self) -> app.EvaluacionAccesibilidadWindow:
        window = app.EvaluacionAccesibilidadWindow(self.root)
        self.addCleanup(lambda: window.winfo_exists() and window.destroy())
        window.update_idletasks()
        return window

    def test_maybe_resume_form_builds_section_routes_without_missing_methods(self) -> None:
        window = object.__new__(app.EvaluacionAccesibilidadWindow)

        with patch.object(app, "_consume_pending_draft_restore", return_value=False):
            with patch.object(app.evaluacion_accesibilidad, "cache_file_exists", return_value=False):
                result = window._maybe_resume_form()

        self.assertFalse(result)

    def test_section_3_transition_to_section_4_renders_and_autosaves_payload(self) -> None:
        window = self._create_window()

        window._show_section_3()
        window._confirm_section_3()

        self.assertTrue(hasattr(window, "section4_level_var"))
        self.assertTrue(window.section4_desc.winfo_exists())

        window.section4_level_var.set("Medio")
        window._update_section4_description()
        window._pending_autosave()

        payload = app.evaluacion_accesibilidad.get_form_cache().get("section_4", {})
        self.assertEqual(payload.get("nivel_accesibilidad"), "Medio")
        self.assertEqual(
            payload.get("descripcion"),
            app.evaluacion_accesibilidad.SECTION_4["descriptions"]["Medio"],
        )
        self.assertEqual(
            window.section4_desc.get("1.0", "end-1c"),
            app.evaluacion_accesibilidad.SECTION_4["descriptions"]["Medio"],
        )

    def test_modalidad_alias_restores_legacy_mixta_as_mixto(self) -> None:
        window = self._create_window()
        modalidad = window.fields["modalidad"]

        applied = app._set_widget_value_from_snapshot(modalidad, "Mixta")

        self.assertTrue(applied)
        self.assertEqual(modalidad.get(), "Mixto")
        self.assertEqual(getattr(modalidad, "_snapshot_value_aliases", {}), {"Mixta": "Mixto"})

        company_data = {
            field["id"]: f"valor-{field['id']}"
            for field in app.evaluacion_accesibilidad.SECTION_1["fields"]
            if field["source"] != "input"
        }
        payload = app.evaluacion_accesibilidad.confirm_section_1(
            company_data,
            {
                "fecha_visita": "2026-03-26",
                "modalidad": modalidad.get(),
                "nit_empresa": "900123456",
            },
        )

        self.assertEqual(payload["modalidad"], "Mixto")
        self.assertEqual(
            app.evaluacion_accesibilidad.get_form_cache()["section_1"]["modalidad"],
            "Mixto",
        )


if __name__ == "__main__":
    unittest.main()
