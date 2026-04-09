from __future__ import annotations

import os
import tempfile
import tkinter as tk
from copy import deepcopy
from types import SimpleNamespace
from unittest.mock import patch

import app
from tests.tk_test_utils import TkTestCase, destroy_widget


class Section2DraftRestoreTests(TkTestCase):
    def setUp(self) -> None:
        super().setUp()
        self.tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(self.tempdir.cleanup)

        self.selection_cache_path = os.path.join(self.tempdir.name, "seleccion_incluyente.json")
        self.contratacion_cache_path = os.path.join(self.tempdir.name, "contratacion_incluyente.json")
        self.drafts_path = os.path.join(self.tempdir.name, "form_drafts_il.json")
        self.process_drafts_dir = os.path.join(self.tempdir.name, "form_drafts")

        self.showerror_patcher = patch.object(app.messagebox, "showerror", return_value=None)
        self.showinfo_patcher = patch.object(app.messagebox, "showinfo", return_value=None)
        self.showwarning_patcher = patch.object(app.messagebox, "showwarning", return_value=None)
        self.showerror = self.showerror_patcher.start()
        self.showinfo = self.showinfo_patcher.start()
        self.showwarning = self.showwarning_patcher.start()
        self.addCleanup(self.showerror_patcher.stop)
        self.addCleanup(self.showinfo_patcher.stop)
        self.addCleanup(self.showwarning_patcher.stop)

        patchers = [
            patch.object(app, "_maximize_window", lambda _window: None),
            patch.object(app, "_attach_autoexpand", lambda *_args, **_kwargs: None),
            patch.object(app, "_attach_dictation_for_section", return_value=None),
            patch.object(app, "_get_drafts_path", return_value=self.drafts_path),
            patch.object(app, "_get_form_drafts_dir", return_value=self.process_drafts_dir),
            patch.object(app.seleccion_incluyente, "_get_cache_path", return_value=self.selection_cache_path),
            patch.object(app.contratacion_incluyente, "_get_cache_path", return_value=self.contratacion_cache_path),
            patch.object(app.seleccion_incluyente, "get_usuarios_reca_cedulas", return_value=[]),
            patch.object(app.contratacion_incluyente, "get_usuarios_reca_cedulas", return_value=[]),
        ]
        for patcher in patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        app.seleccion_incluyente.clear_form_cache()
        app.contratacion_incluyente.clear_form_cache()
        self.addCleanup(app.seleccion_incluyente.clear_form_cache)
        self.addCleanup(app.contratacion_incluyente.clear_form_cache)

    def _make_hub(self):
        hub = self.root
        hub.current_user_profile = {"usuario_login": "tester@example.com"}
        hub.current_user = ""
        hub._drafts_btn = None
        hub._empresa_names_cache = []
        hub._open_form_windows = {}
        hub._pending_draft_restore = None
        hub._release_form_window = lambda _form_id, _window: None
        hub._register_open_form_window = lambda form_id, window: hub._open_form_windows.__setitem__(
            form_id, window
        )
        hub._refresh_drafts_badge = lambda: None
        hub.show_toast = lambda _text: None
        hub._get_current_user_login = lambda: "tester@example.com"
        hub._capture_window_draft_state = (
            lambda window, module: app.HubWindow._capture_window_draft_state(hub, window, module)
        )
        hub._draft_state_has_content = (
            lambda cache_snapshot, ui_snapshot: app.HubWindow._draft_state_has_content(
                hub, cache_snapshot, ui_snapshot
            )
        )
        hub._persist_form_draft = (
            lambda window, **kwargs: app.HubWindow._persist_form_draft(hub, window, **kwargs)
        )
        hub._save_current_form_draft = (
            lambda window, **kwargs: app.HubWindow._save_current_form_draft(hub, window, **kwargs)
        )
        hub._install_form_autosave_bindings = lambda _window: None
        hub._schedule_window_draft_autosave = lambda _window, delay_ms=250: None
        return hub

    def _bind_runtime(self, hub, window, form_meta):
        app.HubWindow._bind_form_runtime(hub, window, form_meta)
        window.update_idletasks()

    def _seed_section1(self, module):
        section_1 = {
            "nombre_empresa": "Empresa Demo",
            "nit_empresa": "900123456",
        }
        module.set_section_cache("section_1", section_1, source="manual")
        module.save_cache_to_file()

    def _read_single_draft(self):
        data = app._load_drafts_store()
        return deepcopy(data["users"]["tester@example.com"][0])

    def _read_process_draft(self, draft_summary):
        return deepcopy(app._load_process_draft_document(draft_summary["draft_path"]))

    def _write_raw_drafts_index(self, payload):
        folder = os.path.dirname(self.drafts_path)
        if folder:
            os.makedirs(folder, exist_ok=True)
        with open(self.drafts_path, "w", encoding="utf-8") as handle:
            import json

            json.dump(payload, handle, ensure_ascii=False, indent=2)

    def _reopen_draft(self, hub, draft_summary, *, window_cls, form_meta):
        created_windows = []

        def _open_form(_form_meta):
            window = window_cls(hub)
            created_windows.append(window)
            self.addCleanup(destroy_widget, window)
            self._bind_runtime(hub, window, form_meta)
            return window

        hub._open_form = _open_form
        app.HubWindow._open_draft_entry(hub, draft_summary)
        self.assertTrue(created_windows)
        window = created_windows[-1]
        window.update_idletasks()
        return window

    def test_selection_section2_restore_migrates_v1_draft_with_three_oferentes(self) -> None:
        hub = self._make_hub()
        form_meta = app.seleccion_incluyente.register_form()
        self._seed_section1(app.seleccion_incluyente)

        window = app.SeleccionIncluyenteWindow(hub)
        self.addCleanup(destroy_widget, window)
        self._bind_runtime(hub, window, form_meta)
        window._show_section_2()
        window._section2_add_block()
        window._section2_add_block()
        window.update_idletasks()

        window.section2_shared_desarrollo_widget.insert("1.0", "Desarrollo seleccion demo")
        for idx, fields in enumerate(window.oferente_blocks, start=1):
            fields["nombre_oferente"].insert(0, f"Oferente {idx}")
            fields["cedula"].set(f"1000{idx}")
            fields["telefono_oferente"].insert(0, f"30000000{idx}")
            fields["fecha_nacimiento"].insert(0, f"0{idx}/0{idx}/200{idx}")

        app.HubWindow._save_current_form_draft(
            hub,
            window,
            silent=True,
            allow_empty=True,
            toast_text="",
        )

        draft_summary = self._read_single_draft()
        cache_snapshot, ui_snapshot, ui_section = hub._capture_window_draft_state(window, app.seleccion_incluyente)
        legacy_draft = {
            "draft_id": draft_summary["draft_id"],
            "form_id": draft_summary["form_id"],
            "form_name": draft_summary["form_name"],
            "company_key": draft_summary["company_key"],
            "company_name": draft_summary["company_name"],
            "created_at": draft_summary["created_at"],
            "updated_at": draft_summary["updated_at"],
            "last_section": draft_summary["last_section"],
            "draft_session_key": getattr(window, "_draft_session_key", ""),
            "cache": deepcopy(cache_snapshot),
            "ui_section": ui_section,
            "ui_snapshot": deepcopy(ui_snapshot),
        }
        legacy_draft["cache"]["section_2"] = []
        self._write_raw_drafts_index({"version": 1, "users": {"tester@example.com": [legacy_draft]}})

        destroy_widget(window)
        app.seleccion_incluyente.clear_form_cache()

        reopened = self._reopen_draft(
            hub,
            self._read_single_draft(),
            window_cls=app.SeleccionIncluyenteWindow,
            form_meta=form_meta,
        )

        self.assertEqual(len(reopened.oferente_blocks), 3)
        self.assertEqual(
            reopened.section2_shared_desarrollo_widget.get("1.0", tk.END).strip(),
            "Desarrollo seleccion demo",
        )
        self.assertEqual(reopened.oferente_blocks[0]["nombre_oferente"].get().strip(), "Oferente 1")
        self.assertEqual(reopened.oferente_blocks[1]["nombre_oferente"].get().strip(), "Oferente 2")
        self.assertEqual(reopened.oferente_blocks[2]["nombre_oferente"].get().strip(), "Oferente 3")
        self.assertEqual(
            reopened.oferente_blocks[0]["edad"].get().strip(),
            str(app._calc_age_from_digits("01012001", min_year=1900)),
        )

        migrated = self._read_single_draft()
        migrated_doc = self._read_process_draft(migrated)
        self.assertEqual(migrated["draft_format_version"], app.FORM_PROCESS_DRAFT_VERSION)
        self.assertEqual(len(migrated_doc["cache"]["section_2"]), 3)
        self.assertEqual(
            migrated_doc["cache"]["section_2"][1]["desarrollo_actividad"],
            "Desarrollo seleccion demo",
        )

    def test_contratacion_section2_restore_migrates_v1_draft_with_three_vinculados(self) -> None:
        hub = self._make_hub()
        form_meta = app.contratacion_incluyente.register_form()
        self._seed_section1(app.contratacion_incluyente)

        window = app.ContratacionIncluyenteWindow(hub)
        self.addCleanup(destroy_widget, window)
        self._bind_runtime(hub, window, form_meta)
        window._show_section_2()
        window._section2_add_block()
        window._section2_add_block()
        window.update_idletasks()

        window.section2_shared_desarrollo_widget.insert("1.0", "Desarrollo contratacion demo")
        etnico_options = tuple(window.oferente_blocks[0]["grupo_etnico_cual"].cget("values"))
        etnico_value = next(
            (value for value in etnico_options if "no aplica" not in str(value).lower()),
            etnico_options[0],
        )
        for idx, fields in enumerate(window.oferente_blocks, start=1):
            fields["nombre_oferente"].insert(0, f"Vinculado {idx}")
            fields["cedula"].set(f"2000{idx}")
            fields["fecha_nacimiento"].insert(0, f"0{idx}/0{idx}/199{idx}")
            fields["grupo_etnico"].set("Si")
            fields["grupo_etnico_cual"].set(etnico_value)
            sync_fn = getattr(fields["grupo_etnico"], "_no_aplica_dropdown_sync", None)
            if callable(sync_fn):
                sync_fn()

        app.HubWindow._save_current_form_draft(
            hub,
            window,
            silent=True,
            allow_empty=True,
            toast_text="",
        )

        draft_summary = self._read_single_draft()
        cache_snapshot, ui_snapshot, ui_section = hub._capture_window_draft_state(window, app.contratacion_incluyente)
        legacy_draft = {
            "draft_id": draft_summary["draft_id"],
            "form_id": draft_summary["form_id"],
            "form_name": draft_summary["form_name"],
            "company_key": draft_summary["company_key"],
            "company_name": draft_summary["company_name"],
            "created_at": draft_summary["created_at"],
            "updated_at": draft_summary["updated_at"],
            "last_section": draft_summary["last_section"],
            "draft_session_key": getattr(window, "_draft_session_key", ""),
            "cache": deepcopy(cache_snapshot),
            "ui_section": ui_section,
            "ui_snapshot": deepcopy(ui_snapshot),
        }
        legacy_draft["cache"]["section_2"] = []
        self._write_raw_drafts_index({"version": 1, "users": {"tester@example.com": [legacy_draft]}})

        destroy_widget(window)
        app.contratacion_incluyente.clear_form_cache()

        reopened = self._reopen_draft(
            hub,
            self._read_single_draft(),
            window_cls=app.ContratacionIncluyenteWindow,
            form_meta=form_meta,
        )

        self.assertEqual(len(reopened.oferente_blocks), 3)
        self.assertEqual(
            reopened.section2_shared_desarrollo_widget.get("1.0", tk.END).strip(),
            "Desarrollo contratacion demo",
        )
        self.assertEqual(reopened.oferente_blocks[0]["nombre_oferente"].get().strip(), "Vinculado 1")
        self.assertEqual(reopened.oferente_blocks[1]["nombre_oferente"].get().strip(), "Vinculado 2")
        self.assertEqual(reopened.oferente_blocks[2]["nombre_oferente"].get().strip(), "Vinculado 3")
        self.assertEqual(
            reopened.oferente_blocks[0]["edad"].get().strip(),
            str(app._calc_age_from_digits("01011991", min_year=1900)),
        )
        self.assertEqual(reopened.oferente_blocks[1]["grupo_etnico"].get().strip(), "Si")
        self.assertEqual(reopened.oferente_blocks[2]["grupo_etnico_cual"].get().strip(), etnico_value)

        migrated = self._read_single_draft()
        migrated_doc = self._read_process_draft(migrated)
        self.assertEqual(migrated["draft_format_version"], app.FORM_PROCESS_DRAFT_VERSION)
        self.assertEqual(len(migrated_doc["cache"]["section_2"]), 3)
        self.assertEqual(
            migrated_doc["cache"]["section_2"][2]["desarrollo_actividad"],
            "Desarrollo contratacion demo",
        )

    def test_selection_section2_partial_restore_warns_and_keeps_recovered_data(self) -> None:
        hub = self._make_hub()
        form_meta = app.seleccion_incluyente.register_form()
        self._seed_section1(app.seleccion_incluyente)

        window = app.SeleccionIncluyenteWindow(hub)
        self.addCleanup(destroy_widget, window)
        self._bind_runtime(hub, window, form_meta)
        window._show_section_2()
        window._section2_add_block()
        window._section2_add_block()
        window.update_idletasks()

        window.section2_shared_desarrollo_widget.insert("1.0", "Desarrollo parcial demo")
        for idx, fields in enumerate(window.oferente_blocks, start=1):
            fields["nombre_oferente"].insert(0, f"Parcial {idx}")
            fields["cedula"].set(f"3000{idx}")

        app.HubWindow._save_current_form_draft(
            hub,
            window,
            silent=True,
            allow_empty=True,
            toast_text="",
        )

        draft_summary = self._read_single_draft()
        cache_snapshot, ui_snapshot, ui_section = hub._capture_window_draft_state(window, app.seleccion_incluyente)
        legacy_draft = {
            "draft_id": draft_summary["draft_id"],
            "form_id": draft_summary["form_id"],
            "form_name": draft_summary["form_name"],
            "company_key": draft_summary["company_key"],
            "company_name": draft_summary["company_name"],
            "created_at": draft_summary["created_at"],
            "updated_at": draft_summary["updated_at"],
            "last_section": draft_summary["last_section"],
            "draft_session_key": getattr(window, "_draft_session_key", ""),
            "cache": deepcopy(cache_snapshot),
            "ui_section": ui_section,
            "ui_snapshot": deepcopy(ui_snapshot),
        }
        legacy_draft["cache"]["section_2"] = []
        legacy_draft["ui_snapshot"].append({"path": "999.999", "class": "Entry", "value": "imposible"})
        self._write_raw_drafts_index({"version": 1, "users": {"tester@example.com": [legacy_draft]}})

        destroy_widget(window)
        app.seleccion_incluyente.clear_form_cache()

        reopened = self._reopen_draft(
            hub,
            self._read_single_draft(),
            window_cls=app.SeleccionIncluyenteWindow,
            form_meta=form_meta,
        )

        self.assertEqual(len(reopened.oferente_blocks), 3)
        self.assertEqual(reopened.oferente_blocks[0]["nombre_oferente"].get().strip(), "Parcial 1")
        self.showwarning.assert_called_once()

        migrated = self._read_single_draft()
        migrated_doc = self._read_process_draft(migrated)
        self.assertEqual(migrated["draft_format_version"], app.FORM_PROCESS_DRAFT_VERSION)
        self.assertEqual(len(migrated_doc["cache"]["section_2"]), 3)

    def test_selection_section2_prefill_uses_window_draft_cache_over_dirty_module_cache(self) -> None:
        hub = self._make_hub()
        form_meta = app.seleccion_incluyente.register_form()

        app.seleccion_incluyente.FORM_CACHE["section_2"] = [
            {
                "nombre_oferente": "Dato incorrecto",
                "cedula": "999999",
                "desarrollo_actividad": "Incorrecto",
            }
        ]

        window = app.SeleccionIncluyenteWindow(hub)
        self.addCleanup(destroy_widget, window)
        window._draft_cache = {
            "section_1": {"nombre_empresa": "Empresa Demo", "nit_empresa": "900123456"},
            "section_2": [
                {
                    "nombre_oferente": "Dato correcto",
                    "cedula": "123456",
                    "telefono_oferente": "3000000000",
                    "desarrollo_actividad": "Desde draft",
                }
            ],
        }
        self._bind_runtime(hub, window, form_meta)

        window._show_section_2()
        window.update_idletasks()

        self.assertEqual(window.oferente_blocks[0]["nombre_oferente"].get().strip(), "Dato correcto")
        self.assertEqual(window.oferente_blocks[0]["cedula"].get().strip(), "123456")
        self.assertEqual(window.section2_shared_desarrollo_widget.get("1.0", tk.END).strip(), "Desde draft")

    def test_selection_fresh_open_clears_legacy_cache_file(self) -> None:
        hub = self._make_hub()
        form_meta = app.seleccion_incluyente.register_form()
        app.seleccion_incluyente.set_section_cache(
            "section_1",
            {"nombre_empresa": "Empresa Legacy", "nit_empresa": "900999999"},
            source="manual",
        )
        app.seleccion_incluyente.save_cache_to_file()
        self.assertTrue(os.path.exists(self.selection_cache_path))

        window = app.SeleccionIncluyenteWindow(hub)
        self.addCleanup(destroy_widget, window)
        self._bind_runtime(hub, window, form_meta)
        window.update_idletasks()

        self.assertFalse(os.path.exists(self.selection_cache_path))
        self.assertEqual(app.seleccion_incluyente.get_form_cache(), {})
