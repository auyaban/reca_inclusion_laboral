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


class PrefixedDropdownSyncTests(_TkTestCase):
    def setUp(self) -> None:
        super().setUp()
        app.seleccion_incluyente.clear_form_cache()
        app.contratacion_incluyente.clear_form_cache()
        app.induccion_operativa.clear_form_cache()
        self.addCleanup(app.seleccion_incluyente.clear_form_cache)
        self.addCleanup(app.contratacion_incluyente.clear_form_cache)
        self.addCleanup(app.induccion_operativa.clear_form_cache)

    def test_seleccion_section_4_syncs_both_directions(self) -> None:
        with patch.object(app.seleccion_incluyente, "cache_file_exists", return_value=False):
            with patch.object(app.seleccion_incluyente, "get_usuarios_reca_cedulas", return_value=[]):
                window = app.SeleccionIncluyenteWindow(self.root)
        self.addCleanup(destroy_widget, window)

        window._show_section_2()
        fields = window.oferente_blocks[0]

        alergias_nivel = fields["alergias_nivel_apoyo"]
        alergias_tipo = fields["alergias_tipo"]
        alergias_nivel.set("0. No requiere apoyo.")
        getattr(alergias_nivel, "_nivel_apoyo_observacion_sync")()
        self.assertTrue(alergias_tipo.get().startswith("0."))

        alergias_tipo.set("2. No conoce si presenta alguna alergia.")
        getattr(alergias_tipo, "_prefixed_dropdown_sync")()
        self.assertTrue(alergias_nivel.get().startswith("2."))

        aseo_nivel = fields["aseo_nivel_apoyo"]
        aseo_obs = fields["alimentacion"]
        aseo_nivel.set("3. Nivel de apoyo alto.")
        getattr(aseo_nivel, "_nivel_apoyo_observacion_sync")()
        self.assertTrue(aseo_obs.get().startswith("3."))

    def test_contratacion_section_5_syncs_both_directions(self) -> None:
        with patch.object(app.contratacion_incluyente, "cache_file_exists", return_value=False):
            with patch.object(app.contratacion_incluyente, "get_usuarios_reca_cedulas", return_value=[]):
                window = app.ContratacionIncluyenteWindow(self.root)
        self.addCleanup(destroy_widget, window)

        window._show_section_2()
        fields = window.oferente_blocks[0]

        condiciones_nivel = fields["condiciones_salariales_nivel_apoyo"]
        condiciones_obs = fields["condiciones_salariales_observacion"]
        condiciones_obs.set("3. Se explica de manera completa las condiciones salariales asignadas al cargo.")
        getattr(condiciones_obs, "_prefixed_dropdown_sync")()
        self.assertTrue(condiciones_nivel.get().startswith("3."))

        condiciones_nivel.set("0. No requiere apoyo.")
        getattr(condiciones_nivel, "_nivel_apoyo_observacion_sync")()
        self.assertTrue(condiciones_obs.get().startswith("0."))

        contrato_lee_nivel = fields["contrato_lee_nivel_apoyo"]
        contrato_lee_obs = fields["contrato_lee_observacion"]
        contrato_lee_nivel.set("0. No requiere apoyo.")
        getattr(contrato_lee_nivel, "_nivel_apoyo_observacion_sync")()
        self.assertEqual(contrato_lee_obs.get(), "No aplica.")

    def test_induccion_operativa_section_4_syncs_both_directions(self) -> None:
        with patch.object(app.induccion_operativa, "cache_file_exists", return_value=False):
            window = app.InduccionOperativaWindow(self.root)
        self.addCleanup(destroy_widget, window)

        window._show_section_4()
        widgets = window.section4_item_widgets["reconoce_instrucciones"]
        nivel = widgets["nivel_apoyo"]
        observaciones = widgets["observaciones"]

        nivel.set("2. Nivel de apoyo medio.")
        getattr(nivel, "_nivel_apoyo_observacion_sync")()
        self.assertTrue(observaciones.get().startswith("2."))

        observaciones.set("1. Requiere especificacion de instrucciones.")
        getattr(observaciones, "_prefixed_dropdown_sync")()
        self.assertTrue(nivel.get().startswith("1."))


if __name__ == "__main__":
    unittest.main()
