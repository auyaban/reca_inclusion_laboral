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


if __name__ == "__main__":
    unittest.main()
