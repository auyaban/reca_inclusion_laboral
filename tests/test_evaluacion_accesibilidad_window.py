from __future__ import annotations

import tkinter as tk
import unittest
from unittest.mock import patch

import app
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
        self.addCleanup(destroy_widget, window)
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

    def test_section_2_1_observaciones_use_multiline_text_widget(self) -> None:
        window = self._create_window()

        window._show_section_2()
        observaciones = window.section2_1_fields["transporte_publico"]["observaciones"]

        self.assertIsInstance(observaciones, tk.Text)

        observaciones.insert("1.0", "Primera linea.\nSegunda linea.")
        window._pending_autosave()

        payload = app.evaluacion_accesibilidad.get_form_cache().get("section_2_1", {})
        self.assertEqual(
            payload.get("transporte_publico_observaciones"),
            "Primera linea.\nSegunda linea.",
        )

    def test_section_2_free_text_fields_use_multiline_text_widgets(self) -> None:
        window = self._create_window()

        sections = [
            ("_show_section_2", "section2_1_fields"),
            ("_show_section_2_2", "section2_2_fields"),
            ("_show_section_2_3", "section2_3_fields"),
            ("_show_section_2_4", "section2_4_fields"),
            ("_show_section_2_5", "section2_5_fields"),
            ("_show_section_2_6", "section2_6_fields"),
        ]

        for render_method, fields_attr in sections:
            with self.subTest(section=fields_attr):
                getattr(window, render_method)()
                fields = getattr(window, fields_attr)
                free_text_widgets = []
                for widgets in fields.values():
                    for key, widget in widgets.items():
                        if key in {"observaciones", "detalle", "texto"}:
                            free_text_widgets.append(widget)
                            self.assertIsInstance(widget, tk.Text)
                self.assertTrue(free_text_widgets)

    def test_detail_text_widgets_start_with_two_lines(self) -> None:
        window = self._create_window()

        window._show_section_2_2()
        detail_widget = next(
            widgets["texto"]
            for widgets in window.section2_2_fields.values()
            if "texto" in widgets
        )

        self.assertIsInstance(detail_widget, tk.Text)
        self.assertEqual(int(detail_widget.cget("height")), 2)

        window._show_section_2_4()
        detail_widget = next(
            widgets["detalle"]
            for widgets in window.section2_4_fields.values()
            if "detalle" in widgets
        )

        self.assertIsInstance(detail_widget, tk.Text)
        self.assertEqual(int(detail_widget.cget("height")), 2)

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

    def test_section_5_hides_suggested_adjustments_until_aplica_is_selected(self) -> None:
        window = self._create_window()

        window._show_section_5()
        first_item = app.evaluacion_accesibilidad.SECTION_5["items"][0]["id"]
        widgets = window.section5_fields[first_item]
        combo = widgets["lista"]
        suggested_label = widgets["_suggested_label"]
        suggested_value = widgets["_suggested_value"]

        window.update_idletasks()
        self.assertFalse(suggested_label.winfo_ismapped())
        self.assertFalse(suggested_value.winfo_ismapped())

        combo.set("Aplica")
        combo.event_generate("<<ComboboxSelected>>")
        window.update_idletasks()

        self.assertTrue(suggested_label.winfo_ismapped())
        self.assertTrue(suggested_value.winfo_ismapped())

    def test_confirm_section_5_sets_no_aplica_adjustments_for_all_items(self) -> None:
        payload = {}
        for item in app.evaluacion_accesibilidad.SECTION_5["items"]:
            item_id = item["id"]
            payload[item_id] = "No aplica"
            payload[f"{item_id}_nota"] = "Nota: no aplica"

        result = app.evaluacion_accesibilidad.confirm_section_5(payload)

        for item in app.evaluacion_accesibilidad.SECTION_5["items"]:
            with self.subTest(item=item["id"]):
                self.assertEqual(result[f"{item['id']}_ajustes"], "No aplica")

    def test_section_5_continue_ignores_internal_ui_labels(self) -> None:
        window = self._create_window()

        window._show_section_5()
        for item in app.evaluacion_accesibilidad.SECTION_5["items"]:
            widgets = window.section5_fields[item["id"]]
            widgets["lista"].set("No aplica")
            widgets["nota"].insert(0, "Sin ajuste")

        window._confirm_section_5()

        payload = app.evaluacion_accesibilidad.get_form_cache().get("section_5", {})
        self.assertTrue(payload)
        self.assertEqual(payload["discapacidad_fisica"], "No aplica")
        self.assertEqual(payload["discapacidad_fisica_ajustes"], "No aplica")
        self.assertEqual(window.header_title.cget("text"), "6. OBSERVACIONES")


if __name__ == "__main__":
    unittest.main()
