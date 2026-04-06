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


class Section1ResumeStateTests(_TkTestCase):
    def setUp(self) -> None:
        super().setUp()
        for module in (
            app.seleccion_incluyente,
            app.condiciones_vacante,
        ):
            module.clear_form_cache()
            self.addCleanup(module.clear_form_cache)

    def test_seleccion_section_1_restores_cached_company_and_enables_continue(self) -> None:
        app.seleccion_incluyente.set_section_cache(
            "section_1",
            {
                "fecha_visita": "2026-04-06",
                "modalidad": "Virtual",
                "nit_empresa": "900123456",
                "nombre_empresa": "Colvanes",
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
            },
        )
        with patch.object(app.seleccion_incluyente, "cache_file_exists", return_value=False):
            with patch.object(app.seleccion_incluyente, "get_usuarios_reca_cedulas", return_value=[]):
                window = app.SeleccionIncluyenteWindow(self.root)
        self.addCleanup(destroy_widget, window)
        window.update_idletasks()

        self.assertEqual(window.fields["nit_empresa"].get(), "900123456")
        self.assertEqual(window.fields["nombre_busqueda"].get(), "Colvanes")
        self.assertEqual(window.company_data.get("nombre_empresa"), "Colvanes")
        self.assertEqual(str(window.continue_btn.cget("state")), "normal")

    def test_condiciones_section_1_restores_cached_company_without_custom_prefill(self) -> None:
        app.condiciones_vacante.set_section_cache(
            "section_1",
            {
                "fecha_visita": "2026-04-06",
                "modalidad": "Presencial",
                "nit_empresa": "800999111",
                "nombre_empresa": "Empresa Demo",
                "direccion_empresa": "Carrera 7",
                "correo_1": "demo@example.com",
                "contacto_empresa": "Pedro",
                "telefono_empresa": "7654321",
                "cargo": "Coordinador",
                "ciudad_empresa": "Bogota",
                "sede_empresa": "Norte",
                "caja_compensacion": "Compensar",
                "asesor": "Marta",
                "profesional_asignado": "Diego",
            },
        )
        with patch.object(app.condiciones_vacante, "cache_file_exists", return_value=False):
            window = app.CondicionesVacanteWindow(self.root)
        self.addCleanup(destroy_widget, window)
        window.update_idletasks()

        self.assertEqual(window.fields["nit_empresa"].get(), "800999111")
        self.assertEqual(window.fields["nombre_busqueda"].get(), "Empresa Demo")
        self.assertEqual(window.company_data.get("nombre_empresa"), "Empresa Demo")
        self.assertEqual(str(window.continue_btn.cget("state")), "normal")


if __name__ == "__main__":
    unittest.main()
