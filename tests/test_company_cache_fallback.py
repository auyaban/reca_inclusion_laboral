import unittest
from unittest.mock import patch

from formularios import common
from formularios.evaluacion_programa import evaluacion_accesibilidad
from formularios.induccion_organizacional import induccion_organizacional


class CompanyCacheFallbackTests(unittest.TestCase):
    def test_merge_company_row_fills_missing_values_from_cache(self):
        row = {
            "nit_empresa": "900696296-4",
            "nombre_empresa": "CORONA INDUSTRIAL SAS",
            "direccion_empresa": None,
            "zona_empresa": "",
            "sede_empresa": "",
        }
        cached = {
            "nit_empresa": "900696296-4",
            "nombre_empresa": "CORONA INDUSTRIAL SAS",
            "direccion_empresa": "Avenida el dorado #86-85",
            "zona_empresa": "Chapinero",
        }

        with patch.object(common, "_find_cached_company_row", return_value=cached):
            merged = common._merge_company_row_with_cache(
                row,
                field_map=evaluacion_accesibilidad.SECTION_1_SUPABASE_MAP,
                nit="900696296-4",
            )

        self.assertEqual(merged["direccion_empresa"], "Avenida el dorado #86-85")
        self.assertEqual(merged["zona_empresa"], "Chapinero")
        self.assertEqual(merged["sede_empresa"], "Chapinero")

    def test_get_empresa_by_nit_uses_cache_when_online_row_has_blank_address(self):
        online_row = {
            "asesor": "Deimi Yisela Torres Reyes",
            "caja_compensacion": "No Compensar",
            "cargo": "Coordinador Desarrollo de Talento",
            "ciudad_empresa": "Bogotá",
            "contacto_empresa": "Brajhan Plazas Chavez",
            "correo_1": "anlopeze@corona.com.co",
            "direccion_empresa": None,
            "nit_empresa": "900696296-4",
            "nombre_empresa": "CORONA INDUSTRIAL SAS",
            "profesional_asignado": "Laura Alejandra Perez Bustacara",
            "telefono_empresa": "3104765912",
            "zona_empresa": "Chapinero",
        }
        cached = dict(online_row, direccion_empresa="Avenida el dorado #86-85")

        with patch.object(evaluacion_accesibilidad, "_supabase_get", return_value=[online_row]):
            with patch.object(common, "_find_cached_company_row", return_value=cached):
                company = evaluacion_accesibilidad.get_empresa_by_nit("900696296-4")

        self.assertIsNotNone(company)
        self.assertEqual(company["direccion_empresa"], "Avenida el dorado #86-85")
        self.assertEqual(company["sede_empresa"], "Chapinero")

    def test_induccion_organizacional_normalizes_legacy_direccion_key(self):
        payload = {
            "fecha_visita": "2026-04-02",
            "modalidad": "Presencial",
            "nombre_empresa": "CORONA INDUSTRIAL SAS",
            "ciudad_empresa": "Bogotá",
            "dirección_empresa": "Avenida el dorado #86-85",
            "nit_empresa": "900696296-4",
            "correo_1": "anlopeze@corona.com.co",
            "telefono_empresa": "3104765912",
            "contacto_empresa": "Brajhan Plazas Chavez",
            "cargo": "Coordinador",
            "caja_compensacion": "No Compensar",
            "sede_empresa": "Chapinero",
            "asesor": "Deimi",
            "profesional_asignado": "Laura",
        }

        induccion_organizacional.FORM_CACHE.clear()
        induccion_organizacional.SECTION_1_CACHE.clear()
        induccion_organizacional.set_section_cache("section_1", payload)

        section_1 = induccion_organizacional.get_form_cache()["section_1"]
        self.assertEqual(section_1["direccion_empresa"], "Avenida el dorado #86-85")
        self.assertNotIn("dirección_empresa", section_1)


if __name__ == "__main__":
    unittest.main()
