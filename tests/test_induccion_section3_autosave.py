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


class InduccionSection3AutosaveTests(_TkTestCase):
    def setUp(self) -> None:
        super().setUp()
        app.induccion_organizacional.clear_form_cache()
        app.induccion_operativa.clear_form_cache()
        self.addCleanup(app.induccion_organizacional.clear_form_cache)
        self.addCleanup(app.induccion_operativa.clear_form_cache)

        patchers = [
            patch.object(app, "_consume_pending_draft_restore", return_value=False),
            patch.object(app.induccion_organizacional, "cache_file_exists", return_value=False),
            patch.object(app.induccion_organizacional, "save_cache_to_file", return_value=None),
            patch.object(app.induccion_operativa, "cache_file_exists", return_value=False),
            patch.object(app.induccion_operativa, "save_cache_to_file", return_value=None),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

    def _create_org_window(self) -> app.InduccionOrganizacionalWindow:
        window = app.InduccionOrganizacionalWindow(self.root)
        self.addCleanup(lambda: window.winfo_exists() and window.destroy())
        window.update_idletasks()
        return window

    def _create_operativa_window(self) -> app.InduccionOperativaWindow:
        window = app.InduccionOperativaWindow(self.root)
        self.addCleanup(lambda: window.winfo_exists() and window.destroy())
        window.update_idletasks()
        return window

    def test_induccion_organizacional_section_3_autosave_keeps_nested_payload(self) -> None:
        window = self._create_org_window()

        window._show_section_3()
        widgets = window.section3_fields["historia_empresa"]
        widgets["visto"].set("Si")
        widgets["responsable"].insert(0, "Andrea")
        widgets["medio_socializacion"].set("Video")
        widgets["descripcion"].insert(0, "Se socializo la historia de la empresa.")

        window._confirm_section_3()

        payload = app.induccion_organizacional.get_form_cache().get("section_3", {})
        self.assertEqual(payload["historia_empresa"]["visto"], "Si")
        self.assertEqual(payload["historia_empresa"]["responsable"], "Andrea")
        self.assertEqual(payload["historia_empresa"]["medio_socializacion"], "Video")
        self.assertEqual(
            payload["historia_empresa"]["descripcion"],
            "Se socializo la historia de la empresa.",
        )

    def test_induccion_operativa_section_3_autosave_keeps_nested_payload(self) -> None:
        window = self._create_operativa_window()

        window._show_section_3()
        first_item_id = app.induccion_operativa.SECTION_3["items"][0]["id"]
        widgets = window.section3_fields[first_item_id]
        widgets["ejecucion"].set("Si")
        widgets["observaciones"].insert(0, "Actividad ejecutada y explicada.")

        window._confirm_section_3()

        payload = app.induccion_operativa.get_form_cache().get("section_3", {})
        self.assertEqual(payload[first_item_id]["ejecucion"], "Si")
        self.assertEqual(
            payload[first_item_id]["observaciones"],
            "Actividad ejecutada y explicada.",
        )


class InduccionSection3GuardTests(unittest.TestCase):
    def setUp(self) -> None:
        app.induccion_organizacional.clear_form_cache()
        app.induccion_operativa.clear_form_cache()
        self.addCleanup(app.induccion_organizacional.clear_form_cache)
        self.addCleanup(app.induccion_operativa.clear_form_cache)

    def test_autosave_keeps_existing_meaningful_section_when_new_payload_is_empty(self) -> None:
        original_payload = {
            "historia_empresa": {
                "visto": "Si",
                "responsable": "Andrea",
                "medio_socializacion": "Video",
                "descripcion": "Contenido diligenciado",
            }
        }
        app.induccion_organizacional.set_section_cache("section_3", original_payload)

        with patch.object(app.induccion_organizacional, "save_cache_to_file", return_value=None):
            app._autosave_section(app.induccion_organizacional, "section_3", lambda: {})

        self.assertEqual(
            app.induccion_organizacional.get_form_cache().get("section_3"),
            original_payload,
        )

    def test_set_section_cache_records_history_snapshot(self) -> None:
        payload = {
            "historia_empresa": {
                "visto": "Si",
                "responsable": "Andrea",
            }
        }

        app.induccion_organizacional.set_section_cache("section_3", payload, source="manual")

        cache = app.induccion_organizacional.get_form_cache()
        history = cache.get("_section_history", {}).get("section_3", [])
        self.assertEqual(len(history), 1)
        self.assertEqual(history[0]["payload"], payload)
        self.assertEqual(history[0]["source"], "manual")
        self.assertEqual(cache.get("_last_saved_section"), "section_3")

    def test_induccion_organizacional_export_validation_rejects_empty_section_3(self) -> None:
        app.induccion_organizacional.set_section_cache(
            "section_4",
            [{"medio": "No aplica", "recomendacion": ""}],
        )
        app.induccion_organizacional.set_section_cache(
            "section_6",
            [{"nombre": "Leidy", "cargo": "Profesional"}],
        )

        with self.assertRaises(RuntimeError):
            app.induccion_organizacional._validate_cache_before_export()

    def test_induccion_operativa_export_validation_rejects_empty_section_3(self) -> None:
        app.induccion_operativa.set_section_cache(
            "section_4",
            {"items": {"foo": {"nivel_apoyo": "Medio", "observaciones": ""}}},
        )
        app.induccion_operativa.set_section_cache(
            "section_9",
            [{"nombre": "Leidy", "cargo": "Profesional"}],
        )

        with self.assertRaises(RuntimeError):
            app.induccion_operativa._validate_cache_before_export()

    def test_find_guarded_missing_sections_detects_required_section_lost_after_history(self) -> None:
        cache_snapshot = {
            "section_3": {},
            "_section_history": {
                "section_3": [
                    {
                        "saved_at": "2026-03-27 17:00:00",
                        "source": "autosave",
                        "payload": {"historia_empresa": {"visto": "Si"}},
                    }
                ]
            },
        }

        missing = app._find_guarded_missing_sections("induccion_organizacional", cache_snapshot)

        self.assertEqual(missing, [("section_3", "Seccion 3")])


if __name__ == "__main__":
    unittest.main()
