from __future__ import annotations

import unittest
from unittest.mock import ANY, Mock, patch

import drive_upload
from formularios.seguimientos import seguimientos


class SeguimientosPdfBundleTests(unittest.TestCase):
    def test_list_pdf_followup_candidates_uses_only_persisted_stages(self) -> None:
        workflow = {
            "max_seguimientos": 3,
            "sheet_progress": [
                {"sheet_name": seguimientos.SHEET_BASE, "status": "completed", "title": "Ficha inicial del proceso"},
                {"sheet_name": f"{seguimientos.SHEET_PREFIX}1", "status": "completed", "title": "Seguimiento 1"},
                {"sheet_name": f"{seguimientos.SHEET_PREFIX}2", "status": "in_progress", "title": "Seguimiento 2"},
                {"sheet_name": f"{seguimientos.SHEET_PREFIX}3", "status": "not_started", "title": "Seguimiento 3"},
            ],
        }
        base_payload = {
            "seguimiento_fechas_1_3": ["2026-04-10", "2026-04-20", ""],
            "seguimiento_fechas_4_6": ["", "", ""],
        }
        followups = {
            1: {"fecha_seguimiento": "2026-04-11"},
            2: {"fecha_seguimiento": ""},
            3: {"fecha_seguimiento": ""},
        }

        with patch.object(seguimientos, "get_workflow_state", return_value=workflow):
            with patch.object(seguimientos, "get_base_payload", return_value=base_payload):
                with patch.object(
                    seguimientos,
                    "get_followup_payload",
                    side_effect=lambda _case_ref, index: followups[index],
                ):
                    candidates = seguimientos.list_pdf_followup_candidates({"file_id": "sheet-1"})

        self.assertEqual([item["followup_index"] for item in candidates], [1, 2])
        self.assertEqual(candidates[0]["fecha_seguimiento"], "2026-04-11")
        self.assertEqual(candidates[1]["fecha_seguimiento"], "2026-04-20")

    def test_build_pdf_export_bundle_base_only(self) -> None:
        case_ref = {"file_id": "sheet-1", "folder_id": "case-folder-1"}
        base_payload = {
            "fecha_visita": "2026-04-07",
            "modalidad": "Presencial",
            "nombre_empresa": "Empresa Demo",
            "nit_empresa": "900123456",
            "profesional_asignado": "Profesional RECA",
            "nombre_vinculado": "Persona Demo",
            "cedula": "10101010",
            "cargo_vinculado": "Analista",
        }

        with patch.object(
            seguimientos,
            "get_case_meta",
            return_value={"base_sheet_name": seguimientos.SHEET_BASE},
        ):
            with patch.object(seguimientos, "get_base_payload", return_value=base_payload):
                bundle = seguimientos.build_pdf_export_bundle(case_ref)

        self.assertEqual(bundle["tipo_acta"], "seguimiento")
        self.assertEqual(bundle["fecha_servicio"], "2026-04-07")
        self.assertEqual(bundle["extra_name"], "Ficha inicial")
        self.assertEqual(bundle["selected_sheet_names"], [seguimientos.SHEET_BASE])
        self.assertEqual(bundle["temp_parent_folder_id"], "case-folder-1")
        self.assertEqual(bundle["acta_metadata"]["document_variant"], "base_only")
        self.assertIsNone(bundle["acta_metadata"]["included_followup_index"])
        self.assertEqual(bundle["acta_metadata"]["participantes"][0]["cedula"], "10101010")

    def test_build_pdf_export_bundle_with_followup(self) -> None:
        case_ref = {"file_id": "sheet-2", "folder_id": "case-folder-2"}
        base_payload = {
            "fecha_visita": "2026-04-07",
            "modalidad": "Presencial",
            "nombre_empresa": "Empresa Demo",
            "nit_empresa": "900123456",
            "profesional_asignado": "Profesional RECA",
            "nombre_vinculado": "Persona Demo",
            "cedula": "10101010",
            "cargo_vinculado": "Analista",
        }
        followup_payload = {
            "seguimiento_numero": "2",
            "fecha_seguimiento": "2026-04-15",
            "modalidad": "Virtual",
            "tipo_apoyo": "Ajustes razonables",
            "situacion_encontrada": "Sin novedades",
            "estrategias_ajustes": "Seguimiento mensual",
            "asistentes": [
                {"nombre": "Profesional RECA", "cargo": "Profesional"},
                {"nombre": "", "cargo": ""},
            ],
        }

        with patch.object(
            seguimientos,
            "get_case_meta",
            return_value={"base_sheet_name": seguimientos.SHEET_BASE},
        ):
            with patch.object(seguimientos, "get_base_payload", return_value=base_payload):
                with patch.object(seguimientos, "get_followup_payload", return_value=followup_payload):
                    bundle = seguimientos.build_pdf_export_bundle(case_ref, followup_index=2)

        self.assertEqual(bundle["fecha_servicio"], "2026-04-15")
        self.assertEqual(bundle["extra_name"], "Seguimiento 2")
        self.assertEqual(
            bundle["selected_sheet_names"],
            [seguimientos.SHEET_BASE, f"{seguimientos.SHEET_PREFIX}2"],
        )
        self.assertEqual(bundle["acta_metadata"]["document_variant"], "base_plus_followup")
        self.assertEqual(bundle["acta_metadata"]["included_followup_index"], 2)
        self.assertEqual(bundle["acta_metadata"]["modalidad_servicio"], "Virtual")
        self.assertEqual(bundle["acta_metadata"]["asistentes"][0]["nombre"], "Profesional RECA")


class DriveUploadPdfExportTests(unittest.TestCase):
    def test_create_and_upload_acta_pdf_uses_selected_sheets_bundle(self) -> None:
        service = Mock()

        with patch.object(drive_upload, "_get_or_create_folder", return_value="company-folder"):
            with patch.object(drive_upload, "_export_google_sheet_selection_as_pdf", return_value=b"pdf-raw") as export_pdf:
                with patch.object(drive_upload, "inject_reca_metadata", return_value=b"pdf-meta") as inject_meta:
                    with patch.object(
                        drive_upload,
                        "upload_pdf_to_folder",
                        return_value={"file_id": "pdf-1", "file_name": "demo.pdf", "webViewLink": "https://example.com/pdf"},
                    ) as upload_pdf:
                        result = drive_upload.create_and_upload_acta_pdf(
                            service=service,
                            sheet_file_id="sheet-1",
                            tipo_acta="seguimiento",
                            acta_metadata={"nombre_empresa": "Empresa Demo"},
                            fecha_servicio=__import__("datetime").date(2026, 4, 7),
                            folder_id="pdf-root",
                            folder_name="Empresa Demo",
                            extra="Ficha inicial",
                            selected_sheet_names=[seguimientos.SHEET_BASE],
                            temp_parent_folder_id="case-folder-1",
                        )

        export_pdf.assert_called_once_with(
            service,
            "sheet-1",
            selected_sheet_names=[seguimientos.SHEET_BASE],
            temp_parent_folder_id="case-folder-1",
        )
        inject_meta.assert_called_once_with(b"pdf-raw", {"nombre_empresa": "Empresa Demo"})
        upload_pdf.assert_called_once_with(service, b"pdf-meta", ANY, "company-folder")
        self.assertEqual(result["file_id"], "pdf-1")

    def test_export_google_sheet_selection_as_pdf_trashes_temp_copy_on_failure(self) -> None:
        service = Mock()

        with patch.object(drive_upload, "_copy_spreadsheet_for_pdf_export", return_value="temp-sheet-1"):
            with patch("google_sheets_client.hide_sheets", side_effect=RuntimeError("boom")):
                with patch.object(drive_upload, "_trash_drive_file") as trash_file:
                    with self.assertRaises(RuntimeError):
                        drive_upload._export_google_sheet_selection_as_pdf(
                            service,
                            "sheet-1",
                            selected_sheet_names=[seguimientos.SHEET_BASE],
                            temp_parent_folder_id="case-folder-1",
                        )

        trash_file.assert_called_once_with(service, "temp-sheet-1")


if __name__ == "__main__":
    unittest.main()
