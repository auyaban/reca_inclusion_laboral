from __future__ import annotations

import tempfile
import tkinter as tk
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

import app
from tests.tk_test_utils import TkTestCase, destroy_widget


class LoginAutoLoginPersistenceTests(unittest.TestCase):
    def test_saved_login_credentials_roundtrip_includes_auto_login(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = f"{temp_dir}\\login_credentials.json"
            with patch.object(app, "_get_login_credentials_path", return_value=path):
                with patch.object(app, "_dpapi_encrypt_text", side_effect=lambda text: f"enc::{text}"):
                    with patch.object(
                        app,
                        "_dpapi_decrypt_text",
                        side_effect=lambda text: str(text or "").replace("enc::", "", 1),
                    ):
                        app._save_login_credentials(
                            "tester",
                            "secret",
                            resolved_email="tester@example.com",
                            auto_login=True,
                        )
                        payload = app._load_saved_login_credentials()

        self.assertTrue(payload["remember"])
        self.assertTrue(payload["auto_login"])
        self.assertEqual(payload["username"], "tester")
        self.assertEqual(payload["password"], "secret")
        self.assertEqual(payload["resolved_email"], "tester@example.com")


class LoginAutoLoginBehaviorTests(TkTestCase):
    def test_finalize_login_success_schedules_profesional_normalization(self) -> None:
        dummy = SimpleNamespace(
            remember_login_var=SimpleNamespace(get=lambda: False),
            auto_login_var=SimpleNamespace(get=lambda: False),
            current_user="",
            current_user_profile={},
            login_frame=None,
            _cache_offline_user_auth=Mock(),
            _schedule_profesional_asignado_normalization=Mock(),
            _start_usage_session=Mock(),
            _build_header=Mock(),
            _build_body=Mock(),
            _auto_login_in_progress=True,
        )

        with patch.object(app, "_clear_login_credentials", return_value=None):
            app.HubWindow._finalize_login_success(
                dummy,
                user_row={"usuario_login": "tester"},
                username_input="tester",
                username="tester",
                password="secret",
            )

        dummy._schedule_profesional_asignado_normalization.assert_called_once_with()
        dummy._start_usage_session.assert_called_once_with()
        self.assertFalse(dummy._auto_login_in_progress)

    def test_schedule_profesional_normalization_absorbs_background_errors(self) -> None:
        class _ImmediateThread:
            def __init__(self, target=None, daemon=None, **kwargs):
                self._target = target

            def start(self):
                if self._target is not None:
                    self._target()

        dummy = SimpleNamespace(
            _profesional_normalization_in_progress=False,
            _normalize_profesional_asignado=Mock(side_effect=RuntimeError("timeout")),
        )

        with patch.object(app.threading, "Thread", _ImmediateThread):
            with patch.object(app, "_log_capture", return_value=None) as log_capture:
                result = app.HubWindow._schedule_profesional_asignado_normalization(dummy)

        self.assertTrue(result)
        self.assertFalse(dummy._profesional_normalization_in_progress)
        dummy._normalize_profesional_asignado.assert_called_once_with()
        log_capture.assert_called()

    def test_update_auto_login_toggle_state_disables_checkbox_when_remember_is_false(self) -> None:
        dummy = SimpleNamespace()
        dummy.remember_login_var = tk.BooleanVar(master=self.root, value=False)
        dummy.auto_login_var = tk.BooleanVar(master=self.root, value=True)
        dummy._auto_login_cb = tk.Checkbutton(self.root, variable=dummy.auto_login_var)

        app.HubWindow._update_auto_login_toggle_state(dummy)

        self.assertFalse(dummy.auto_login_var.get())
        self.assertEqual(str(dummy._auto_login_cb.cget("state")), "disabled")

    def test_maybe_attempt_auto_login_starts_background_login_once(self) -> None:
        status = Mock()
        dummy = SimpleNamespace(
            _auto_login_attempted=False,
            _auto_login_in_progress=False,
            _startup_precheck_completed=True,
            login_frame=object(),
            remember_login_var=SimpleNamespace(get=lambda: True),
            auto_login_var=SimpleNamespace(get=lambda: True),
            login_user_entry=SimpleNamespace(get=lambda: "tester"),
            login_pass_entry=SimpleNamespace(get=lambda: "secret"),
            login_status=status,
            _start_auto_login_async=Mock(),
        )

        app.HubWindow._maybe_attempt_auto_login(dummy)

        self.assertTrue(dummy._auto_login_attempted)
        self.assertTrue(dummy._auto_login_in_progress)
        dummy._start_auto_login_async.assert_called_once_with(username_input="tester", password="secret")
        status.config.assert_called_once()

    def test_maybe_attempt_auto_login_skips_without_password(self) -> None:
        dummy = SimpleNamespace(
            _auto_login_attempted=False,
            _auto_login_in_progress=False,
            _startup_precheck_completed=True,
            login_frame=object(),
            remember_login_var=SimpleNamespace(get=lambda: True),
            auto_login_var=SimpleNamespace(get=lambda: True),
            login_user_entry=SimpleNamespace(get=lambda: "tester"),
            login_pass_entry=SimpleNamespace(get=lambda: ""),
            login_status=Mock(),
            _start_auto_login_async=Mock(),
        )

        app.HubWindow._maybe_attempt_auto_login(dummy)

        self.assertFalse(dummy._auto_login_attempted)
        self.assertFalse(dummy._auto_login_in_progress)
        dummy._start_auto_login_async.assert_not_called()

    def test_maybe_attempt_auto_login_can_ignore_precheck_gate(self) -> None:
        status = Mock()
        dummy = SimpleNamespace(
            _auto_login_attempted=False,
            _auto_login_in_progress=False,
            _startup_precheck_completed=False,
            login_frame=object(),
            remember_login_var=SimpleNamespace(get=lambda: True),
            auto_login_var=SimpleNamespace(get=lambda: True),
            login_user_entry=SimpleNamespace(get=lambda: "tester"),
            login_pass_entry=SimpleNamespace(get=lambda: "secret"),
            login_status=status,
            _start_auto_login_async=Mock(),
        )

        app.HubWindow._maybe_attempt_auto_login(dummy, ignore_precheck=True)

        self.assertTrue(dummy._auto_login_attempted)
        self.assertTrue(dummy._auto_login_in_progress)
        dummy._start_auto_login_async.assert_called_once_with(username_input="tester", password="secret")
        status.config.assert_called_once()

    def test_handle_login_uses_resolved_auth_result_without_waiting_for_precheck(self) -> None:
        dummy = SimpleNamespace(
            login_user_entry=SimpleNamespace(get=lambda: "tester"),
            login_pass_entry=SimpleNamespace(get=lambda: "secret"),
            login_status=Mock(),
            update_idletasks=Mock(),
            _resolve_login_attempt=Mock(
                return_value={
                    "user_row": {"usuario_login": "tester"},
                    "used_offline": True,
                    "auth_exc": TimeoutError("timed out"),
                }
            ),
            _complete_login_with_auth_result=Mock(return_value=True),
            _auto_login_in_progress=False,
        )

        with patch.object(app, "_load_saved_login_credentials", return_value={}):
            result = app.HubWindow._handle_login(dummy)

        self.assertTrue(result)
        dummy._resolve_login_attempt.assert_called_once_with("tester", "secret", cached_email="")
        dummy._complete_login_with_auth_result.assert_called_once()


class LoginPrecheckInteractionTests(unittest.TestCase):
    def _create_window(self):
        patchers = [
            patch.object(app, "_maximize_window", lambda _window: None),
            patch.object(app, "_ensure_drive_upload_worker", return_value=None),
            patch.object(app.HubWindow, "_run_startup_precheck_async", return_value=None),
            patch.object(app, "_load_saved_login_credentials", return_value={}),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)
        try:
            window = app.HubWindow()
        except tk.TclError as exc:
            self.skipTest(f"Tk no disponible: {exc}")
        self.addCleanup(destroy_widget, window)
        window.update_idletasks()
        window.update()
        return window

    def test_login_button_stays_enabled_while_precheck_is_pending(self) -> None:
        window = self._create_window()

        self.assertEqual(str(window._login_btn.cget("state")), "normal")
        window._set_login_ready_state(False, "Verificando servicios...")

        self.assertEqual(str(window._login_btn.cget("state")), "normal")
        self.assertEqual(window.login_status.cget("text"), "Verificando servicios...")

    def test_login_button_invokes_submit_even_before_precheck_finishes(self) -> None:
        with patch.object(app.HubWindow, "_handle_login", autospec=True, return_value=True) as handle_login:
            window = self._create_window()

            handle_login.reset_mock()
            window._startup_precheck_completed = False
            window._login_btn.invoke()

        handle_login.assert_called_once_with(window)

    def test_login_return_key_invokes_submit_even_before_precheck_finishes(self) -> None:
        with patch.object(app.HubWindow, "_handle_login", autospec=True, return_value=True) as handle_login:
            window = self._create_window()

            handle_login.reset_mock()
            window._startup_precheck_completed = False
            window.login_user_entry.focus_force()
            window.update()
            window.login_user_entry.event_generate("<Return>")
            window.update()

        handle_login.assert_called_once_with(window)


if __name__ == "__main__":
    unittest.main()
