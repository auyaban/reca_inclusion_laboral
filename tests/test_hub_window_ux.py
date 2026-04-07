from __future__ import annotations

import tkinter as tk
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

import app
from tkinter import ttk

from tests.tk_test_utils import destroy_widget


class _ImmediateThread:
    def __init__(self, target=None, daemon=None, **kwargs):
        self._target = target

    def start(self):
        if self._target is not None:
            self._target()

    def is_alive(self):
        return False


class _UpdateFlowDummy:
    _normalize_update_snapshot = app.HubWindow._normalize_update_snapshot
    _confirm_and_start_update = app.HubWindow._confirm_and_start_update
    _maybe_prompt_startup_update = app.HubWindow._maybe_prompt_startup_update
    _handle_manual_update_snapshot = app.HubWindow._handle_manual_update_snapshot


def _collect_label_texts(widget):
    texts = []
    for child in widget.winfo_children():
        try:
            text = str(child.cget("text") or "").strip()
        except Exception:
            text = ""
        if text:
            texts.append(text)
        texts.extend(_collect_label_texts(child))
    return texts


class HubWindowUiTests(unittest.TestCase):
    def _create_window(self):
        patchers = [
            patch.object(app, "_maximize_window", lambda _window: None),
            patch.object(app, "_ensure_drive_upload_worker", return_value=None),
            patch.object(app.HubWindow, "_build_login", return_value=None),
            patch.object(app.HubWindow, "_refresh_drafts_badge", return_value=None),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)
        try:
            window = app.HubWindow()
        except tk.TclError as exc:
            self.skipTest(f"Tk no disponible: {exc}")
        self.addCleanup(destroy_widget, window)
        window.withdraw()
        window.update_idletasks()
        window.update()
        return window

    def test_build_body_renders_form_description_and_company_hint(self) -> None:
        window = self._create_window()

        with patch.object(
            app,
            "get_forms",
            return_value=[
                {
                    "id": "demo",
                    "name": "Formulario Demo",
                    "hub_description": "Descripción demo",
                }
            ],
        ):
            with patch.object(window, "_get_assigned_companies", return_value=[]):
                window._build_body()
                window.update_idletasks()

        texts = _collect_label_texts(window.body)

        self.assertIn("Descripción demo", texts)
        self.assertIn("Doble clic para editar la empresa seleccionada.", texts)

    def test_form_card_state_changes_when_singleton_window_opens_and_closes(self) -> None:
        window = self._create_window()

        with patch.object(
            app,
            "get_forms",
            return_value=[
                {
                    "id": "demo",
                    "name": "Formulario Demo",
                    "hub_description": "Descripción demo",
                }
            ],
        ):
            with patch.object(window, "_get_assigned_companies", return_value=[]):
                window._build_body()
                window.update_idletasks()

        button = window._form_action_buttons["demo"]

        window._register_open_form_window("demo", object())
        self.assertEqual(button.cget("text"), "En progreso...")
        self.assertEqual(str(button.cget("state")), "disabled")

        window._release_form_window("demo")
        self.assertEqual(button.cget("text"), "Abrir")
        self.assertEqual(str(button.cget("state")), "normal")

    def test_open_form_reuses_existing_singleton_window(self) -> None:
        existing = Mock()
        existing.winfo_exists.return_value = True
        hub = SimpleNamespace(_open_form_windows={"presentacion_programa": existing})

        with patch.object(app, "_focus_window") as focus_window:
            result = app.HubWindow._open_form(
                hub,
                {
                    "id": "presentacion_programa",
                    "name": "Presentación",
                    "singleton_window": True,
                },
            )

        self.assertIs(result, existing)
        focus_window.assert_called_once_with(existing)

    def test_refresh_database_cache_reports_success_inline(self) -> None:
        window = self._create_window()
        button = ttk.Button(window, text="Actualizar Base de Datos")
        status_label = tk.Label(window, text="")

        def _after(delay, callback):
            if delay == 0:
                callback()
            return None

        dummy = SimpleNamespace(
            _refresh_db_btn=button,
            _refresh_db_status_label=status_label,
            _clear_form_memory_caches=Mock(),
            _get_assigned_companies=Mock(return_value=[{"nombre_empresa": "Empresa Demo"}]),
            _render_companies=Mock(),
            _companies_all=[],
            _empresa_names_cache=[],
            after=_after,
        )

        with patch.object(app.threading, "Thread", _ImmediateThread):
            app.HubWindow._refresh_database_cache(dummy)

        self.assertIn("Base de datos actualizada", status_label.cget("text"))
        dummy._clear_form_memory_caches.assert_called_once()
        dummy._render_companies.assert_called_once()
        self.assertEqual(str(button.cget("state")), "normal")

    def test_refresh_database_cache_reports_error_inline(self) -> None:
        window = self._create_window()
        button = ttk.Button(window, text="Actualizar Base de Datos")
        status_label = tk.Label(window, text="")

        def _after(delay, callback):
            if delay == 0:
                callback()
            return None

        dummy = SimpleNamespace(
            _refresh_db_btn=button,
            _refresh_db_status_label=status_label,
            _clear_form_memory_caches=Mock(),
            _get_assigned_companies=Mock(side_effect=RuntimeError("Supabase unavailable")),
            _render_companies=Mock(),
            _companies_all=[],
            _empresa_names_cache=[],
            after=_after,
        )

        with patch.object(app.threading, "Thread", _ImmediateThread):
            app.HubWindow._refresh_database_cache(dummy)

        self.assertIn("base de datos", status_label.cget("text").lower())
        dummy._render_companies.assert_not_called()
        self.assertEqual(str(button.cget("state")), "normal")

    def test_sync_panel_shows_separate_service_status_lines(self) -> None:
        window = self._create_window()
        window._is_online = True
        window._service_probe_cache = {
            "internet": {"ok": True, "status_text": "Todo correcto", "error_code": "", "detail": ""},
            "supabase": {"ok": True, "status_text": "Disponible", "error_code": "", "detail": ""},
            "drive": {"ok": False, "status_text": "Timeout", "error_code": "timeout", "detail": ""},
        }
        before_children = set(window.winfo_children())

        with patch.object(app, "_get_supabase_write_queue_snapshot", return_value=[]):
            with patch.object(app, "_get_supabase_failed_writes_snapshot", return_value=[]):
                with patch.object(app, "_get_drive_upload_queue_snapshot", return_value=[]):
                    with patch.object(app, "_get_drive_failed_uploads_snapshot", return_value=[]):
                        window._open_sync_panel()
                        window.update_idletasks()

        modal = next(child for child in window.winfo_children() if child not in before_children)
        texts = _collect_label_texts(modal)

        self.assertIn("Estado de servicios", texts)
        self.assertTrue(any(text.startswith("Internet:") for text in texts))
        self.assertTrue(any(text.startswith("Supabase:") for text in texts))
        self.assertTrue(any(text.startswith("Drive:") for text in texts))

    def test_apply_update_snapshot_updates_version_label_and_schedules_startup_prompt(self) -> None:
        window = self._create_window()
        snapshot = {
            "local_version": "1.2.3",
            "remote_version": "1.2.4",
            "assets": {},
            "error": None,
        }

        with patch.object(window, "_maybe_prompt_startup_update") as maybe_prompt:
            window._apply_update_snapshot(snapshot, prompt_startup=True)
            window.update_idletasks()
            window.update()

        self.assertEqual(window._version_var.get(), "Versión local: 1.2.3 | GitHub: 1.2.4")
        self.assertEqual(window._latest_update_snapshot["remote_version"], "1.2.4")
        maybe_prompt.assert_called_once()


class HubWindowUpdateFlowTests(unittest.TestCase):
    def test_maybe_prompt_startup_update_accepts_once_per_app_run(self) -> None:
        dummy = _UpdateFlowDummy()
        dummy._startup_update_prompt_shown = False
        dummy._update_assets_are_usable = Mock(return_value=True)
        dummy._start_manual_update = Mock()
        snapshot = {
            "local_version": "1.2.3",
            "remote_version": "1.2.4",
            "assets": {"RECA_INCLUSION_LABORAL_Setup.exe": "https://example.com/setup.exe"},
            "error": None,
        }

        with patch.object(app.messagebox, "askyesno", return_value=True) as askyesno:
            first = app.HubWindow._maybe_prompt_startup_update(dummy, snapshot)
            second = app.HubWindow._maybe_prompt_startup_update(dummy, snapshot)

        self.assertTrue(first)
        self.assertFalse(second)
        self.assertTrue(dummy._startup_update_prompt_shown)
        askyesno.assert_called_once()
        dummy._start_manual_update.assert_called_once_with(snapshot["assets"])

    def test_maybe_prompt_startup_update_decline_does_not_repeat_prompt(self) -> None:
        dummy = _UpdateFlowDummy()
        dummy._startup_update_prompt_shown = False
        dummy._update_assets_are_usable = Mock(return_value=True)
        dummy._start_manual_update = Mock()
        snapshot = {
            "local_version": "1.2.3",
            "remote_version": "1.2.4",
            "assets": {"RECA_INCLUSION_LABORAL_Setup.exe": "https://example.com/setup.exe"},
            "error": None,
        }

        with patch.object(app.messagebox, "askyesno", return_value=False) as askyesno:
            first = app.HubWindow._maybe_prompt_startup_update(dummy, snapshot)
            second = app.HubWindow._maybe_prompt_startup_update(dummy, snapshot)

        self.assertFalse(first)
        self.assertFalse(second)
        self.assertTrue(dummy._startup_update_prompt_shown)
        askyesno.assert_called_once()
        dummy._start_manual_update.assert_not_called()

    def test_maybe_prompt_startup_update_skips_errors_silently(self) -> None:
        dummy = _UpdateFlowDummy()
        dummy._startup_update_prompt_shown = False
        dummy._update_assets_are_usable = Mock(return_value=True)
        dummy._start_manual_update = Mock()
        snapshot = {
            "local_version": "1.2.3",
            "remote_version": None,
            "assets": {},
            "error": "timeout",
        }

        with patch.object(app.messagebox, "askyesno") as askyesno:
            result = app.HubWindow._maybe_prompt_startup_update(dummy, snapshot)

        self.assertFalse(result)
        self.assertFalse(dummy._startup_update_prompt_shown)
        askyesno.assert_not_called()
        dummy._start_manual_update.assert_not_called()

    def test_handle_manual_update_snapshot_reports_missing_installer_asset(self) -> None:
        dummy = _UpdateFlowDummy()
        dummy._update_assets_are_usable = Mock(return_value=False)
        dummy._apply_update_snapshot = Mock(
            side_effect=lambda snapshot, prompt_startup=False: app.HubWindow._normalize_update_snapshot(
                dummy, snapshot
            )
        )
        dummy._get_expected_update_installer_asset_name = Mock(
            return_value="RECA_INCLUSION_LABORAL_Setup.exe"
        )
        dummy._confirm_and_start_update = Mock()
        snapshot = {
            "local_version": "1.2.3",
            "remote_version": "1.2.4",
            "assets": {},
            "error": None,
        }

        with patch.object(app.messagebox, "showerror") as showerror:
            result = app.HubWindow._handle_manual_update_snapshot(dummy, snapshot)

        self.assertFalse(result)
        dummy._confirm_and_start_update.assert_not_called()
        showerror.assert_called_once()
        self.assertIn("RECA_INCLUSION_LABORAL_Setup.exe", showerror.call_args.args[1])


if __name__ == "__main__":
    unittest.main()
