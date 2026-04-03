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
            patch.object(app, "_attach_autoexpand", lambda *_args, **_kwargs: None),
            patch.object(app.messagebox, "showerror", return_value=None),
            patch.object(app.messagebox, "showinfo", return_value=None),
            patch.object(app.messagebox, "showwarning", return_value=None),
            patch.object(app, "_consume_pending_draft_restore", return_value=False),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)


class GroupalSection2LayoutTests(_TkTestCase):
    def setUp(self) -> None:
        super().setUp()
        app.seleccion_incluyente.clear_form_cache()
        app.contratacion_incluyente.clear_form_cache()
        self.addCleanup(app.seleccion_incluyente.clear_form_cache)
        self.addCleanup(app.contratacion_incluyente.clear_form_cache)

    def test_seleccion_section_2_keeps_shared_activity_before_oferentes(self) -> None:
        with patch.object(app.seleccion_incluyente, "cache_file_exists", return_value=False):
            with patch.object(app.seleccion_incluyente, "get_usuarios_reca_cedulas", return_value=[]):
                window = app.SeleccionIncluyenteWindow(self.root)
        self.addCleanup(destroy_widget, window)

        window._show_section_2()
        window.update_idletasks()

        self.assertEqual(window.header_title.cget("text"), "2. DESARROLLO DE LA ACTIVIDAD")
        self.assertEqual(
            window.section2_shared_desarrollo_frame.cget("text"),
            "2. DESARROLLO DE LA ACTIVIDAD",
        )
        self.assertEqual(window.oferente_frames[0].cget("text"), "Oferente 1")
        self.assertEqual(
            getattr(window.oferente_frames[0], "_section2_frame").cget("text"),
            "3. DATOS DEL OFERENTE",
        )
        self.assertIs(
            window.section2_shared_desarrollo_frame.master,
            window.oferente_frames[0].master,
        )

    def test_seleccion_section_2_rebinds_dictation_after_adding_oferente(self) -> None:
        with patch.object(app.seleccion_incluyente, "cache_file_exists", return_value=False):
            with patch.object(app.seleccion_incluyente, "get_usuarios_reca_cedulas", return_value=[]):
                window = app.SeleccionIncluyenteWindow(self.root)
        self.addCleanup(destroy_widget, window)

        window._form_id = "seleccion_incluyente"
        window._current_section = "section_2"

        with patch.object(app, "_attach_dictation_for_section", return_value=None) as attach_mock:
            window._show_section_2()
            window.update()
            initial_calls = attach_mock.call_count

            add_button = next(
                widget
                for widget in app._iter_widget_tree(window.section_container)
                if isinstance(widget, app.ttk.Button) and widget.cget("text") == "Agregar oferente"
            )
            add_button.invoke()
            window.update()

        self.assertGreater(attach_mock.call_count, initial_calls)
        self.assertEqual(attach_mock.call_args.args[1:], ("seleccion_incluyente", "section_2"))

    def test_contratacion_section_2_uses_groupal_titles_and_options_with_single_vinculado(self) -> None:
        with patch.object(app.contratacion_incluyente, "cache_file_exists", return_value=False):
            with patch.object(app.contratacion_incluyente, "get_usuarios_reca_cedulas", return_value=[]):
                window = app.ContratacionIncluyenteWindow(self.root)
        self.addCleanup(destroy_widget, window)

        window._show_section_2()
        window.update_idletasks()

        self.assertEqual(window.header_title.cget("text"), "2. DESARROLLO DE LA ACTIVIDAD")
        self.assertEqual(
            window.section2_shared_desarrollo_frame.cget("text"),
            "2. DESARROLLO DE LA ACTIVIDAD",
        )
        self.assertEqual(
            getattr(window.oferente_frames[0], "_section2_frame").cget("text"),
            "3. DATOS DEL VINCULADO",
        )
        self.assertEqual(
            getattr(window.oferente_frames[0], "_section3_frame").cget("text"),
            "4. DATOS ADICIONALES",
        )
        contrato_lee_values = tuple(window.oferente_blocks[0]["contrato_lee_observacion"].cget("values"))
        self.assertNotIn("0. No requiere apoyo.", contrato_lee_values)

        contrato_lee_nivel = window.oferente_blocks[0]["contrato_lee_nivel_apoyo"]
        contrato_lee_nivel.set("2. Nivel de apoyo medio.")
        getattr(contrato_lee_nivel, "_nivel_apoyo_observacion_sync")()
        self.assertTrue(window.oferente_blocks[0]["contrato_lee_observacion"].get().startswith("2."))

        contrato_lee_nivel.set("No aplica.")
        getattr(contrato_lee_nivel, "_nivel_apoyo_observacion_sync")()
        self.assertEqual(window.oferente_blocks[0]["contrato_lee_observacion"].get(), "No aplica.")

        condiciones_nivel = window.oferente_blocks[0]["condiciones_salariales_nivel_apoyo"]
        condiciones_nivel.set("0. No requiere apoyo.")
        getattr(condiciones_nivel, "_nivel_apoyo_observacion_sync")()
        self.assertTrue(
            window.oferente_blocks[0]["condiciones_salariales_observacion"].get().startswith("0.")
        )

        window._show_section_6()
        self.assertEqual(window.header_title.cget("text"), "5. AJUSTES RAZONABLES Y RECOMENDACIONES")

        window._show_section_7()
        self.assertEqual(window.header_title.cget("text"), "6. ASISTENTES")

    def test_contratacion_section_2_rebinds_dictation_after_adding_vinculado(self) -> None:
        with patch.object(app.contratacion_incluyente, "cache_file_exists", return_value=False):
            with patch.object(app.contratacion_incluyente, "get_usuarios_reca_cedulas", return_value=[]):
                window = app.ContratacionIncluyenteWindow(self.root)
        self.addCleanup(destroy_widget, window)

        window._form_id = "contratacion_incluyente"
        window._current_section = "section_2"

        with patch.object(app, "_attach_dictation_for_section", return_value=None) as attach_mock:
            window._show_section_2()
            window.update()
            initial_calls = attach_mock.call_count

            add_button = next(
                widget
                for widget in app._iter_widget_tree(window.section_container)
                if isinstance(widget, app.ttk.Button) and widget.cget("text") == "Agregar vinculado"
            )
            add_button.invoke()
            window.update()

        self.assertGreater(attach_mock.call_count, initial_calls)
        self.assertEqual(attach_mock.call_args.args[1:], ("contratacion_incluyente", "section_2"))


if __name__ == "__main__":
    unittest.main()
