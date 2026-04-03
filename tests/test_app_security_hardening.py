from __future__ import annotations

import json
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import app


class AppSecurityHardeningTests(unittest.TestCase):
    def test_authenticate_user_offline_rejects_expired_cache_entry(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "offline_auth_users.json"
            path.write_text(
                json.dumps(
                    {
                        "version": app.OFFLINE_AUTH_STORE_VERSION,
                        "users": {
                            "tester": {
                                "id": "1",
                                "usuario_login": "tester",
                                "usuario_pass_hash": app._hash_password("secret"),
                                "nombre_profesional": "Persona Demo",
                                "programa": "IL",
                                "cached_at": time.time() - ((app.OFFLINE_AUTH_TTL_DAYS + 1) * 86400),
                                "ttl_days": app.OFFLINE_AUTH_TTL_DAYS,
                            }
                        },
                    }
                ),
                encoding="utf-8",
            )

            with patch.object(app, "_get_offline_auth_path", return_value=str(path)):
                result = app.HubWindow._authenticate_user_offline(object(), "tester", "secret")

            self.assertIsNone(result)
            payload = json.loads(path.read_text(encoding="utf-8"))
            self.assertEqual(payload["users"], {})

    def test_authenticate_user_offline_accepts_fresh_cache_entry(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "offline_auth_users.json"
            path.write_text(
                json.dumps(
                    {
                        "version": app.OFFLINE_AUTH_STORE_VERSION,
                        "users": {
                            "tester": {
                                "id": "1",
                                "usuario_login": "tester",
                                "usuario_pass_hash": app._hash_password("secret"),
                                "nombre_profesional": "Persona Demo",
                                "programa": "IL",
                                "cached_at": time.time(),
                                "ttl_days": app.OFFLINE_AUTH_TTL_DAYS,
                            }
                        },
                    }
                ),
                encoding="utf-8",
            )

            with patch.object(app, "_get_offline_auth_path", return_value=str(path)):
                result = app.HubWindow._authenticate_user_offline(object(), "tester", "secret")

        self.assertEqual(result["usuario_login"], "tester")
        self.assertEqual(result["_auth_source"], "offline")

    def test_open_local_file_safely_rejects_paths_outside_allowed_roots(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = Path(tmpdir) / "demo.xlsx"
            file_path.write_text("demo", encoding="utf-8")

            with self.assertRaisesRegex(RuntimeError, "no está permitida"):
                app._open_local_file_safely(str(file_path), allowed_roots=[str(Path(tmpdir) / "otro")])

    def test_open_local_file_safely_allows_expected_extensions_under_allowed_root(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = Path(tmpdir) / "demo.xlsx"
            file_path.write_text("demo", encoding="utf-8")

            with patch.object(app.os, "startfile", Mock()) as startfile:
                app._open_local_file_safely(str(file_path), allowed_roots=[tmpdir])

            startfile.assert_called_once_with(str(file_path))

    def test_password_candidates_rejects_overlong_passwords(self) -> None:
        self.assertEqual(app._password_candidates("x" * (app.MAX_PASSWORD_LENGTH + 1)), [])

    def test_finish_with_loading_allows_opening_generated_file_from_custom_folder(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            file_path = Path(tmpdir) / "resultado.xlsx"
            file_path.write_text("demo", encoding="utf-8")
            loading = SimpleNamespace(
                set_status=Mock(),
                set_progress=Mock(),
                window=SimpleNamespace(grab_release=Mock()),
                close=Mock(),
            )

            with patch.object(app.messagebox, "askyesno", return_value=True):
                with patch.object(app, "_open_local_file_safely") as open_local_file:
                    app._finish_with_loading(loading, "Listo", open_target=str(file_path))

            open_local_file.assert_called_once()
            kwargs = open_local_file.call_args.kwargs
            self.assertIn(str(file_path.parent), kwargs["allowed_roots"])


if __name__ == "__main__":
    unittest.main()
