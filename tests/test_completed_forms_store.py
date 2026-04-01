from __future__ import annotations

import tempfile
import unittest
from datetime import datetime, timezone
from unittest.mock import patch

import app


class CompletedFormsStoreTests(unittest.TestCase):
    def test_prune_completed_forms_entries_keeps_only_last_30_days(self) -> None:
        now = datetime(2026, 3, 27, 17, 0, 0, tzinfo=timezone.utc)
        entries = [
            {"registro_id": "recent", "finalizado_at_iso": "2026-03-20T17:00:00+00:00"},
            {"registro_id": "old", "finalizado_at_iso": "2026-02-20T17:00:00+00:00"},
        ]

        pruned = app._prune_completed_forms_entries(entries, now=now)

        self.assertEqual([item["registro_id"] for item in pruned], ["recent"])

    def test_store_completed_form_locally_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            target = f"{tmpdir}\\completed_forms_il.json"
            row = {
                "registro_id": "abc",
                "usuario_login": "tester",
                "nombre_formato": "Induccion Organizacional",
                "nombre_empresa": "Empresa Demo",
                "path_formato": r"C:\tmp\demo.xlsx",
                "upload_status": "synced",
                "finalizado_at_iso": "2026-03-27T12:00:00-05:00",
                "finalizado_at_colombia": "2026-03-27 12:00:00",
                "payload_generated_at": "2026-03-27T17:00:00Z",
                "source_item_key": "induccion_organizacional:session-1",
                "payload_raw": {
                    "form_id": "induccion_organizacional",
                    "cache_snapshot": {"section_1": {"nombre_empresa": "Empresa Demo"}},
                },
            }

            with patch.object(app, "_get_completed_forms_path", return_value=target):
                app._store_completed_form_locally(row)
                store = app._load_completed_forms_store()

        self.assertIn("tester", store["users"])
        self.assertEqual(len(store["users"]["tester"]), 1)
        self.assertEqual(store["users"]["tester"][0]["registro_id"], "abc")

    def test_resolve_completed_restore_form_meta_maps_labs_to_base_form(self) -> None:
        entry = {
            "payload_raw": {
                "form_id": "seleccion_incluyente_labs",
                "cache_snapshot": {"section_1": {"nombre_empresa": "Empresa Demo"}},
            }
        }

        form_meta = app._resolve_completed_restore_form_meta(entry)

        self.assertIsNotNone(form_meta)
        self.assertEqual(form_meta["id"], "seleccion_incluyente")


if __name__ == "__main__":
    unittest.main()
