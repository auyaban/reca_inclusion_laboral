from __future__ import annotations

import unittest
from unittest.mock import patch

import app


class LoginProfileFallbackTests(unittest.TestCase):
    def test_authenticate_user_uses_profile_fallback_when_profesionales_is_denied(self) -> None:
        def _fake_rpc(name, params=None, use_session=True):
            if name == "resolve_login_email":
                return {"resolve_login_email": "test@example.com"}
            if name == "get_my_profesional_profile":
                raise RuntimeError(
                    "Supabase no esta disponible (HTTP 401): "
                    "{'code':'42501','message':'permission denied for table profesionales'}"
                )
            raise AssertionError(f"RPC inesperado: {name}")

        with patch.object(app, "_clear_supabase_session"):
            with patch.object(app, "_supabase_auth_password_login"):
                with patch.object(app, "_supabase_rpc", side_effect=_fake_rpc):
                    profile = app.HubWindow._authenticate_user(object(), "Test", "secret")

        self.assertEqual(profile["usuario_login"], "test")
        self.assertEqual(profile["nombre_profesional"], "test")
        self.assertTrue(profile["_profile_fallback"])
        self.assertEqual(profile["_auth_source"], "jwt_profile_fallback")
        self.assertEqual(profile["_resolved_email"], "test@example.com")


if __name__ == "__main__":
    unittest.main()
