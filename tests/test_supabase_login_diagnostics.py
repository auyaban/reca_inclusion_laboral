from __future__ import annotations

import io
import unittest
import urllib.error
from unittest.mock import patch

from formularios import common, user_messages


class _FakeResponse:
    def __init__(self, status: int = 200, body: str = "") -> None:
        self.status = status
        self._body = body.encode("utf-8")

    def read(self) -> bytes:
        return self._body

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        return None


class SupabaseLoginDiagnosticsTests(unittest.TestCase):
    def test_probe_supabase_service_uses_auth_health_with_supabase_headers(self) -> None:
        captured = {}

        def _fake_urlopen(request, timeout=0):
            captured["url"] = request.full_url
            captured["timeout"] = timeout
            captured["headers"] = dict(request.header_items())
            return _FakeResponse(status=200, body='{"status":"ok"}')

        with patch.object(
            common,
            "_load_supabase_credentials",
            return_value=("https://example.supabase.co", "demo-key"),
        ):
            with patch.object(common.urllib.request, "urlopen", side_effect=_fake_urlopen):
                result = common.probe_supabase_service(timeout=7)

        self.assertTrue(result["ok"])
        self.assertEqual(captured["url"], "https://example.supabase.co/auth/v1/health")
        self.assertEqual(captured["timeout"], 7)
        self.assertEqual(captured["headers"].get("Apikey"), "demo-key")
        self.assertEqual(captured["headers"].get("Authorization"), "Bearer demo-key")

    def test_password_login_reports_email_confirmation_requirement(self) -> None:
        http_error = urllib.error.HTTPError(
            url="https://example.supabase.co/auth/v1/token?grant_type=password",
            code=400,
            msg="Bad Request",
            hdrs=None,
            fp=io.BytesIO(b'{"error_code":"email_not_confirmed","msg":"Email not confirmed"}'),
        )

        with patch.object(
            common,
            "_load_supabase_credentials",
            return_value=("https://example.supabase.co", "demo-key"),
        ):
            with patch.object(common.urllib.request, "urlopen", side_effect=http_error):
                with self.assertRaisesRegex(RuntimeError, "confirmar tu correo"):
                    common._supabase_auth_password_login("demo@example.com", "secret")

    def test_login_message_maps_email_not_confirmed_copy(self) -> None:
        message = user_messages.map_exception_to_user_message(
            "login",
            RuntimeError("Debes confirmar tu correo antes de iniciar sesión."),
        )

        self.assertEqual(message, "Debes confirmar tu correo antes de iniciar sesión.")

    def test_login_message_maps_invalid_credentials_copy(self) -> None:
        message = user_messages.map_exception_to_user_message(
            "login",
            RuntimeError("Usuario y contraseña incorrectos."),
        )

        self.assertEqual(message, "Usuario o contraseña incorrectos.")


if __name__ == "__main__":
    unittest.main()
