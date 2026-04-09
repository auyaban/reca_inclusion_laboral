from __future__ import annotations

import copy
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

import app


class _ModuleStub:
    def get_form_cache(self):
        return {}

    def save_cache_to_file(self):
        return None


class _HubStub:
    current_user_profile = {"usuario_login": "tester@example.com"}
    current_user = ""
    _drafts_btn = None

    def __init__(self, *, cache_snapshot, ui_snapshot, ui_section):
        self._draft_capture = (cache_snapshot, ui_snapshot, ui_section)
        self.badge_refreshes = 0
        self.toasts = []

    def _get_current_user_login(self):
        return "tester@example.com"

    def _capture_window_draft_state(self, _window, _module):
        cache_snapshot, ui_snapshot, ui_section = self._draft_capture
        return copy.deepcopy(cache_snapshot), copy.deepcopy(ui_snapshot), ui_section

    def _draft_state_has_content(self, _cache_snapshot, _ui_snapshot):
        return True

    def _refresh_drafts_badge(self):
        self.badge_refreshes += 1

    def show_toast(self, text):
        self.toasts.append(text)


def _make_window(**overrides):
    payload = {
        "_form_id": "evaluacion_accesibilidad",
        "_form_name": "Evaluacion de Accesibilidad",
        "_draft_id": "",
        "_draft_last_fingerprint": "",
        "_draft_session_key": "session-demo",
        "_draft_autosave_after_id": None,
    }
    payload.update(overrides)
    return SimpleNamespace(**payload)


class DraftAutosaveTests(unittest.TestCase):
    def _draft_patches(self, store=None):
        store_data = {"version": app.FORM_DRAFTS_INDEX_VERSION, "users": {"tester@example.com": []}} if store is None else store
        saved_payloads = []
        patchers = [
            patch.object(app, "_resolve_form_meta", return_value={"supports_drafts": True}),
            patch.dict(app.FORM_MODULE_MAP, {"evaluacion_accesibilidad": _ModuleStub()}),
            patch.object(app, "_load_drafts_store", return_value=store_data),
            patch.object(
                app,
                "_save_process_draft_document",
                side_effect=lambda doc, draft_path="": draft_path or f"C:\\drafts\\{doc['draft_id']}.json",
            ),
            patch.object(
                app,
                "_save_drafts_store",
                side_effect=lambda data: saved_payloads.append(copy.deepcopy(data)),
            ),
            patch.object(app, "_update_window_save_status", return_value=None),
        ]
        return patchers, saved_payloads

    def test_save_current_form_draft_forwards_manual_source(self):
        hub = SimpleNamespace(_persist_form_draft=Mock())
        window = object()

        app.HubWindow._save_current_form_draft(
            hub,
            window,
            silent=True,
            allow_empty=True,
            toast_text="",
        )

        hub._persist_form_draft.assert_called_once_with(
            window,
            allow_empty=True,
            silent=True,
            toast_text="",
            source="manual",
        )

    def test_schedule_window_draft_autosave_forwards_autosave_source(self):
        calls = []
        hub = SimpleNamespace(
            _persist_form_draft=lambda *args, **kwargs: calls.append((args, kwargs))
        )
        window = _make_window()
        window.winfo_exists = lambda: True
        window.after_cancel = lambda _after_id: None

        def _after(_delay_ms, callback):
            callback()
            return "after-1"

        window.after = _after

        with patch.object(app, "_form_supports_drafts", return_value=True):
            app.HubWindow._schedule_window_draft_autosave(hub, window, delay_ms=0)

        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0][0], (window,))
        self.assertEqual(
            calls[0][1],
            {
                "allow_empty": True,
                "silent": True,
                "toast_text": "",
                "source": "autosave",
            },
        )

    def test_persist_form_draft_skips_section_1_autosave(self):
        hub = _HubStub(
            cache_snapshot={"section_1": {"nombre_empresa": "Empresa Demo"}},
            ui_snapshot=[{"path": "section_1.nombre_empresa", "value": "Empresa Demo"}],
            ui_section="section_1",
        )
        window = _make_window(_draft_last_fingerprint="fingerprint-previo")
        patchers, saved_payloads = self._draft_patches()

        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        result = app.HubWindow._persist_form_draft(
            hub,
            window,
            allow_empty=True,
            silent=True,
            source="autosave",
        )

        self.assertFalse(result)
        self.assertEqual(saved_payloads, [])
        self.assertEqual(window._draft_last_fingerprint, "fingerprint-previo")
        self.assertEqual(hub.badge_refreshes, 0)

    def test_persist_form_draft_allows_manual_save_in_section_1(self):
        hub = _HubStub(
            cache_snapshot={"section_1": {"nombre_empresa": "Empresa Demo"}},
            ui_snapshot=[{"path": "section_1.nombre_empresa", "value": "Empresa Demo"}],
            ui_section="section_1",
        )
        window = _make_window()
        patchers, saved_payloads = self._draft_patches()

        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        result = app.HubWindow._persist_form_draft(
            hub,
            window,
            allow_empty=True,
            silent=True,
            source="manual",
        )

        self.assertTrue(result)
        self.assertEqual(len(saved_payloads), 1)
        saved_entry = saved_payloads[0]["users"]["tester@example.com"][0]
        self.assertEqual(saved_entry["last_section"], "section_1")
        self.assertEqual(saved_entry["draft_format_version"], 3)
        self.assertEqual(saved_entry["draft_type"], app.FORM_PROCESS_DRAFT_TYPE)
        self.assertTrue(saved_entry["draft_path"].endswith(".json"))
        self.assertTrue(window._draft_id)
        self.assertEqual(hub.badge_refreshes, 1)

    def test_persist_form_draft_autosave_still_saves_after_section_1(self):
        hub = _HubStub(
            cache_snapshot={"section_2": {"observaciones": "Texto de prueba"}},
            ui_snapshot=[{"path": "section_2.observaciones", "value": "Texto de prueba"}],
            ui_section="section_2",
        )
        window = _make_window()
        patchers, saved_payloads = self._draft_patches()

        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        result = app.HubWindow._persist_form_draft(
            hub,
            window,
            allow_empty=True,
            silent=True,
            source="autosave",
        )

        self.assertTrue(result)
        self.assertEqual(len(saved_payloads), 1)
        saved_entry = saved_payloads[0]["users"]["tester@example.com"][0]
        self.assertEqual(saved_entry["last_section"], "section_2")
        self.assertEqual(saved_entry["draft_format_version"], 3)
        self.assertEqual(saved_entry["draft_type"], app.FORM_PROCESS_DRAFT_TYPE)
        self.assertTrue(window._draft_last_fingerprint)

    def test_persist_form_draft_flushes_pending_section_autosave_before_capture(self):
        hub = _HubStub(
            cache_snapshot={"section_2": {"observaciones": "Texto de prueba"}},
            ui_snapshot=[{"path": "section_2.observaciones", "value": "Texto de prueba"}],
            ui_section="section_2",
        )
        window = _make_window()
        patchers, _saved_payloads = self._draft_patches()

        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        with patch.object(app, "_run_pending_section_autosave", return_value=None) as flush_mock:
            result = app.HubWindow._persist_form_draft(
                hub,
                window,
                allow_empty=True,
                silent=True,
                source="manual",
            )

        self.assertTrue(result)
        flush_mock.assert_called_once_with(window)


if __name__ == "__main__":
    unittest.main()
