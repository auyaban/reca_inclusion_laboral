from __future__ import annotations

import base64
import json
import unittest

from formularios import common


def _build_jwt(payload: dict) -> str:
    encoded = base64.urlsafe_b64encode(json.dumps(payload).encode("utf-8")).decode("ascii").rstrip("=")
    return f"header.{encoded}.signature"


class CommonSecurityHardeningTests(unittest.TestCase):
    def test_get_cache_scope_uses_only_subject_claim(self) -> None:
        token = _build_jwt({"sub": "user-123", "role": "service_role"})
        self.assertEqual(common._get_cache_scope(token), "user:user-123")
        self.assertEqual(common._get_cache_scope(_build_jwt({"role": "service_role"})), "anon")

    def test_safe_json_for_log_redacts_pii(self) -> None:
        payload = {
            "email": "persona@example.com",
            "nit_empresa": "900123456",
            "nombre_profesional": "Persona Demo",
        }

        rendered = common._safe_json_for_log(payload)

        self.assertNotIn("persona@example.com", rendered)
        self.assertNotIn("900123456", rendered)
        self.assertNotIn("Persona Demo", rendered)
        self.assertIn("@example.com", rendered)

    def test_truncate_failed_queue_value_limits_nested_strings(self) -> None:
        rendered = common._truncate_failed_queue_value(
            {"rows": [{"nombre": "x" * 400}]},
            max_string=50,
        )

        self.assertEqual(len(rendered["rows"][0]["nombre"]), 53)
        self.assertTrue(rendered["rows"][0]["nombre"].endswith("..."))

    def test_build_failed_queue_summary_does_not_include_payload_values(self) -> None:
        job = {
            "op": "upsert",
            "table": "usuarios_reca",
            "attempts": 2,
            "rows": [{"nombre": "Persona Demo", "telefono": "+57 300 123 4567"}],
            "filters": {"id": "123"},
            "values": {"correo": "persona@example.com"},
            "on_conflict": "id",
        }

        summary = common._build_failed_queue_summary(job, RuntimeError("boom"))

        self.assertEqual(summary["row_count"], 1)
        self.assertEqual(summary["filter_keys"], ["id"])
        self.assertEqual(summary["value_keys"], ["correo"])
        self.assertNotIn("rows", summary)
        self.assertNotIn("Persona Demo", json.dumps(summary))


if __name__ == "__main__":
    unittest.main()
