from __future__ import annotations

import unittest
from unittest.mock import patch

from formularios.seleccion_incluyente import seleccion_incluyente as seleccion
from formularios.contratacion_incluyente import contratacion_incluyente as contratacion


class SyncUsuariosRecaShapeTests(unittest.TestCase):
    def tearDown(self) -> None:
        seleccion.clear_form_cache()
        contratacion.clear_form_cache()

    def test_seleccion_incluyente_upsert_uses_consistent_keys(self) -> None:
        seleccion.FORM_CACHE["section_2"] = [
            {
                "cedula": "1000061994",
                "nombre_oferente": "Ana",
                "telefono_oferente": "3001112233",
            },
            {
                "cedula": "1000061995",
                "nombre_oferente": "Bruno",
                "telefono_oferente": "",
                "tipo_pension": "Colpensiones",
            },
        ]

        captured = {}

        def _fake_upsert(table, rows, env_path=".env", on_conflict=None):
            captured["table"] = table
            captured["rows"] = rows
            captured["on_conflict"] = on_conflict
            return {"status": "synced"}

        with patch.object(seleccion, "_supabase_upsert_with_queue", side_effect=_fake_upsert):
            count = seleccion.sync_usuarios_reca()

        self.assertEqual(count, 2)
        self.assertEqual(captured["table"], "usuarios_reca")
        self.assertEqual(captured["on_conflict"], "cedula_usuario")
        self.assertEqual(len(captured["rows"]), 2)
        self.assertEqual(set(captured["rows"][0].keys()), set(captured["rows"][1].keys()))
        self.assertIsNone(captured["rows"][1]["telefono_oferente"])
        self.assertEqual(captured["rows"][1]["tipo_pension"], "Colpensiones")

    def test_contratacion_incluyente_upsert_uses_consistent_keys(self) -> None:
        contratacion.FORM_CACHE["section_2"] = [
            {
                "cedula": "1000061994",
                "nombre_oferente": "Ana",
                "tipo_contrato": "Indefinido",
            },
            {
                "cedula": "1000061995",
                "nombre_oferente": "Bruno",
                "telefono_oferente": "3001112233",
            },
        ]
        contratacion.SECTION_1_CACHE["nit_empresa"] = "900123456"
        contratacion.SECTION_1_CACHE["nombre_empresa"] = "Empresa Test"

        captured = {}

        def _fake_upsert(table, rows, env_path=".env", on_conflict=None):
            captured["table"] = table
            captured["rows"] = rows
            captured["on_conflict"] = on_conflict
            return {"status": "synced"}

        with patch.object(contratacion, "_supabase_upsert_with_queue", side_effect=_fake_upsert):
            count = contratacion.sync_usuarios_reca()

        self.assertEqual(count, 2)
        self.assertEqual(captured["table"], "usuarios_reca")
        self.assertEqual(captured["on_conflict"], "cedula_usuario")
        self.assertEqual(len(captured["rows"]), 2)
        self.assertEqual(set(captured["rows"][0].keys()), set(captured["rows"][1].keys()))
        self.assertIsNone(captured["rows"][0]["telefono_oferente"])
        self.assertEqual(captured["rows"][0]["empresa_nit"], "900123456")
        self.assertEqual(captured["rows"][1]["empresa_nombre"], "Empresa Test")

    def test_contratacion_incluyente_upsert_deduplicates_duplicate_cedulas_keeping_last(self) -> None:
        contratacion.FORM_CACHE["section_2"] = [
            {
                "cedula": "1000061994",
                "nombre_oferente": "Ana inicial",
                "telefono_oferente": "",
            },
            {
                "cedula": "1000061994",
                "nombre_oferente": "Ana final",
                "telefono_oferente": "3001112233",
            },
        ]
        contratacion.SECTION_1_CACHE["nit_empresa"] = "900123456"
        contratacion.SECTION_1_CACHE["nombre_empresa"] = "Empresa Test"

        captured = {}

        def _fake_upsert(table, rows, env_path=".env", on_conflict=None):
            captured["table"] = table
            captured["rows"] = rows
            captured["on_conflict"] = on_conflict
            return {"status": "synced"}

        with patch.object(contratacion, "_supabase_upsert_with_queue", side_effect=_fake_upsert):
            count = contratacion.sync_usuarios_reca()

        self.assertEqual(count, 1)
        self.assertEqual(captured["table"], "usuarios_reca")
        self.assertEqual(captured["on_conflict"], "cedula_usuario")
        self.assertEqual(len(captured["rows"]), 1)
        self.assertEqual(captured["rows"][0]["cedula_usuario"], "1000061994")
        self.assertEqual(captured["rows"][0]["nombre_usuario"], "Ana final")
        self.assertEqual(captured["rows"][0]["telefono_oferente"], "3001112233")


if __name__ == "__main__":
    unittest.main()
