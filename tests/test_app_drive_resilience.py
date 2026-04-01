from __future__ import annotations

import sys
import types
import unittest
from unittest.mock import patch

import app
import drive_upload


class DriveResilienceTests(unittest.TestCase):
    def test_background_finalization_skips_local_exists_for_remote_sheet_exports(self) -> None:
        captured = {}

        class _DummyHub:
            def __init__(self) -> None:
                self.current_session_id = "session-demo"

            def track_form_finished(self, form_id):
                captured["tracked_form_id"] = form_id

            def finalize_form_delivery(self, output_path, **kwargs):
                captured["finalize_call"] = {"output_path": output_path, **kwargs}
                return {
                    "status": "synced",
                    "output_path": output_path,
                    "remote_url": kwargs.get("drive_job", {}).get("remote_url", ""),
                    "drive_file_id": kwargs.get("drive_job", {}).get("drive_file_id", ""),
                    "error": "",
                }

        class _DummyWindow:
            def __init__(self) -> None:
                self.master = _DummyHub()
                self._finalize_in_progress = False
                self.destroyed = False

            def destroy(self):
                self.destroyed = True

        class _DummyLoading:
            def close(self):
                captured["loading_closed"] = True

        class _ImmediateThread:
            def __init__(self, target=None, daemon=None, **kwargs):
                self._target = target

            def start(self):
                if self._target is not None:
                    self._target()

        window = _DummyWindow()
        loading = _DummyLoading()

        with patch.object(app, "_guard_form_action", return_value=False):
            with patch.object(app.threading, "Thread", _ImmediateThread):
                with patch.object(app, "HubWindow", _DummyHub):
                    with patch.object(app, "_safe_widget_after", side_effect=lambda _window, fn: fn()):
                        with patch.object(app, "_update_loading_async", return_value=None):
                            with patch.object(app, "_finalize_export_flow", return_value=None):
                                with patch.object(app, "_return_to_hub", return_value=None):
                                    with patch.object(app.os.path, "exists", side_effect=AssertionError("os.path.exists should not run for remote exports")):
                                        app._start_background_finalization(
                                            window,
                                            loading,
                                            form_name="Demo",
                                            company_name="Empresa Demo",
                                            form_id="demo_remote",
                                            worker_fn=lambda: {
                                                "output_path": "https://docs.google.com/spreadsheets/d/sheet-123/edit",
                                                "drive_file_id": "sheet-123",
                                                "already_in_drive": True,
                                            },
                                        )

        self.assertTrue(window.destroyed)
        self.assertEqual(captured["tracked_form_id"], "demo_remote")
        self.assertEqual(
            captured["finalize_call"]["output_path"],
            "https://docs.google.com/spreadsheets/d/sheet-123/edit",
        )
        self.assertEqual(
            captured["finalize_call"]["drive_job"],
            {
                "status": "synced",
                "drive_file_id": "sheet-123",
                "remote_url": "https://docs.google.com/spreadsheets/d/sheet-123/edit",
            },
        )

    def test_finalize_form_delivery_does_not_enqueue_synced_remote_sheet(self) -> None:
        class _DummyHub:
            def create_form_completion_record(self, *args, **kwargs):
                return "registro-demo"

        dummy = _DummyHub()
        remote_url = "https://docs.google.com/spreadsheets/d/sheet-123/edit"

        with patch.object(app, "_update_form_completion_upload_status", return_value=True) as update_status:
            with patch.object(app, "_enqueue_drive_upload_job") as enqueue_job:
                result = app.HubWindow.finalize_form_delivery(
                    dummy,
                    remote_url,
                    form_name="Demo",
                    company_name="Empresa Demo",
                    drive_job={
                        "status": "synced",
                        "drive_file_id": "sheet-123",
                        "remote_url": remote_url,
                    },
                )

        self.assertEqual(result["status"], "synced")
        self.assertEqual(result["remote_url"], remote_url)
        enqueue_job.assert_not_called()
        update_status.assert_called_once_with(
            "registro-demo",
            upload_status="synced",
            upload_error="",
            upload_attempted_at=unittest.mock.ANY,
            uploaded_at=unittest.mock.ANY,
            path_formato=remote_url,
            drive_file_id="sheet-123",
        )

    def test_drive_probe_read_only_skips_write_check(self) -> None:
        fake_google = types.ModuleType("google")
        fake_google_oauth2 = types.ModuleType("google.oauth2")
        fake_service_account = types.ModuleType("google.oauth2.service_account")
        fake_googleapiclient = types.ModuleType("googleapiclient")
        fake_discovery = types.ModuleType("googleapiclient.discovery")

        class _FakeCredentials:
            @staticmethod
            def from_service_account_file(*args, **kwargs):
                return object()

        fake_service_account.Credentials = _FakeCredentials
        fake_discovery.build = lambda *args, **kwargs: object()

        with patch.dict(
            sys.modules,
            {
                "google": fake_google,
                "google.oauth2": fake_google_oauth2,
                "google.oauth2.service_account": fake_service_account,
                "googleapiclient": fake_googleapiclient,
                "googleapiclient.discovery": fake_discovery,
            },
        ):
            with patch.object(drive_upload, "_get_credentials_path", return_value="creds.json"):
                with patch.object(drive_upload, "_get_excel_folder_id", return_value="root-folder"):
                    with patch.object(drive_upload, "_resolve_target_root_id", return_value="resolved-root"):
                        with patch.object(
                            drive_upload,
                            "_probe_parent_read_access",
                            return_value={"sample_id": "sample-1"},
                        ):
                            with patch.object(
                                drive_upload,
                                "_probe_parent_write_access",
                            ) as write_probe:
                                result = drive_upload.probe_drive_service(require_write=False)

        self.assertTrue(result["ok"])
        self.assertIn("autenticado", result["status_text"].lower())
        write_probe.assert_not_called()

    def test_probe_startup_services_can_skip_drive_write_probe(self) -> None:
        with patch.object(app, "check_internet", return_value={"ok": True}):
            with patch.object(app, "probe_supabase_service", return_value={"ok": True}):
                with patch.object(
                    app.drive_upload,
                    "probe_drive_service",
                    return_value={"ok": True},
                ) as drive_probe:
                    result = app.probe_startup_services(require_drive_write=False)

        self.assertTrue(result["internet"]["ok"])
        drive_probe.assert_called_once_with(log_enabled=False, require_write=False)

    def test_probe_startup_services_defaults_to_read_only_drive_probe(self) -> None:
        with patch.object(app, "check_internet", return_value={"ok": True}):
            with patch.object(app, "probe_supabase_service", return_value={"ok": True}):
                with patch.object(
                    app.drive_upload,
                    "probe_drive_service",
                    return_value={"ok": True},
                ) as drive_probe:
                    result = app.probe_startup_services()

        self.assertTrue(result["internet"]["ok"])
        drive_probe.assert_called_once_with(log_enabled=False, require_write=False)

    def test_read_followup_case_state_returns_error_instead_of_raising(self) -> None:
        with patch.object(app.seguimientos, "suggest_next_step", side_effect=RuntimeError("boom")):
            with patch.object(app.seguimientos, "describe_case") as describe_case:
                result = app._read_followup_case_state({"file_id": "demo"})

        self.assertIsNotNone(result["error"])
        self.assertIsNone(result["suggestion"])
        self.assertIsNone(result["summary"])
        describe_case.assert_not_called()

    def test_format_followup_case_state_error_hides_transient_transport_detail(self) -> None:
        exc = RuntimeError(
            "[WinError 10053] Se ha anulado una conexion establecida "
            "por el software en su equipo host"
        )

        message = app._format_followup_case_state_error(exc)

        self.assertIn("falla temporal de conexión", message)
        self.assertNotIn("WinError 10053", message)

    def test_transient_drive_exception_detects_wrapped_winerror_10053_text(self) -> None:
        try:
            raise RuntimeError(
                "[WinError 10053] Se ha anulado una conexion establecida "
                "por el software en su equipo host"
            )
        except RuntimeError as inner:
            try:
                raise RuntimeError("outer") from inner
            except RuntimeError as wrapped:
                self.assertTrue(app._is_transient_drive_exception(wrapped))

    def test_transient_drive_exception_does_not_treat_local_os_errors_as_transient(self) -> None:
        cases = (
            PermissionError("permission denied"),
            FileNotFoundError("No such file or directory"),
        )

        for exc in cases:
            with self.subTest(exc=type(exc).__name__):
                self.assertFalse(app._is_transient_drive_exception(exc))

    def test_load_followup_editor_bootstrap_formats_transient_error(self) -> None:
        winerror = RuntimeError(
            "[WinError 10053] Se ha anulado una conexion establecida "
            "por el software en su equipo host"
        )

        with patch.object(app.seguimientos, "get_case_meta", side_effect=winerror):
            with self.assertRaises(RuntimeError) as caught:
                app._load_followup_editor_bootstrap({"file_id": "demo"})

        self.assertIn("falla temporal de conexion", str(caught.exception).lower())
        self.assertNotIn("WinError 10053", str(caught.exception))

    def test_open_editor_bootstraps_before_constructing_window(self) -> None:
        captured = {}

        class _DummyVar:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

            def set(self, value):
                self.value = value

        class _DummyButton:
            def __init__(self):
                self.last_config = {}

            def config(self, **kwargs):
                self.last_config.update(kwargs)

        class _DummySeguimientosWindow:
            def __init__(self):
                self.case_record = {
                    "source": "drive",
                    "mime_type": app.seguimientos.GOOGLE_SHEETS_MIME,
                    "webViewLink": "https://example.test/case",
                    "local_path": "C:/tmp/demo.xlsx",
                }
                self.case_path = "C:/tmp/demo.xlsx"
                self.cedula_combo = _DummyVar("123456")
                self.compensar_var = _DummyVar("Si")
                self.user_row = {}
                self.linked_company = {}
                self.path_var = _DummyVar("")
                self.open_btn = _DummyButton()

            def _get_case_target(self):
                return self.case_record

            def _run_loading_job(self, **kwargs):
                captured.update(kwargs)

        dummy = _DummySeguimientosWindow()
        bootstrap = {
            "meta": {"max_seguimientos": 3, "base_sheet_name": app.seguimientos.SHEET_BASE},
            "workflow": {
                "base_sheet_name": app.seguimientos.SHEET_BASE,
                "visible_sheets": [app.seguimientos.SHEET_BASE],
                "suggested_sheet": app.seguimientos.SHEET_BASE,
                "message": "ok",
                "max_seguimientos": 3,
            },
            "suggestion": {
                "sheet": app.seguimientos.SHEET_BASE,
                "message": "ok",
                "max_seguimientos": 3,
            },
        }
        editor_instance = object()

        with patch.object(app, "_load_followup_editor_bootstrap", return_value=bootstrap) as load_bootstrap:
            with patch.object(app, "SeguimientoEditorWindow", return_value=editor_instance) as editor_window:
                with patch.object(app, "_focus_window") as focus_window:
                    app.SeguimientosWindow._open_editor(dummy)
                    payload = captured["worker"](lambda *args, **kwargs: None)
                    captured["on_success"](payload)

        self.assertEqual(captured["title"], "Abriendo seguimiento")
        self.assertEqual(captured["on_error_title"], "Seguimientos")
        load_bootstrap.assert_called_once_with(dummy.case_record)
        editor_window.assert_called_once_with(
            dummy,
            "C:/tmp/demo.xlsx",
            case_record=dummy.case_record,
            bootstrap=bootstrap,
        )
        focus_window.assert_called_once_with(editor_instance)

    def test_finalize_form_delivery_falls_back_to_pending_for_unexpected_drive_transport_error(
        self,
    ) -> None:
        class _DummyHub:
            def create_form_completion_record(self, *args, **kwargs):
                return "registro-demo"

        dummy = _DummyHub()
        winerror = RuntimeError(
            "[WinError 10053] Se ha anulado una conexion establecida "
            "por el software en su equipo host"
        )

        with patch.object(app, "_perform_drive_upload_attempt", side_effect=winerror):
            with patch.object(app, "_update_form_completion_upload_status", return_value=False):
                with patch.object(app, "_enqueue_drive_upload_job") as enqueue_job:
                    result = app.HubWindow.finalize_form_delivery(
                        dummy,
                        "C:/tmp/demo.xlsx",
                        form_name="Demo",
                        company_name="Empresa Demo",
                    )

        self.assertEqual(result["status"], "pending")
        self.assertEqual(result["registro_id"], "registro-demo")
        enqueue_job.assert_called_once()


if __name__ == "__main__":
    unittest.main()
