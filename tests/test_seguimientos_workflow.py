from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from formularios.seguimientos import seguimientos


class SeguimientosWorkflowTests(unittest.TestCase):
    def test_workflow_keeps_followup_1_open_when_only_date_is_registered(self) -> None:
        base_payload = {
            "nombre_vinculado": "Persona Demo",
            "cedula": "123456",
            "cargo_vinculado": "Auxiliar",
            "discapacidad": "Fisica",
            "seguimiento_fechas_1_3": ["2026-04-01", "", ""],
            "seguimiento_fechas_4_6": ["", "", ""],
        }

        with patch.object(
            seguimientos,
            "get_case_meta",
            return_value={"base_sheet_name": seguimientos.SHEET_BASE, "max_seguimientos": 3},
        ):
            with patch.object(seguimientos, "get_base_payload", return_value=base_payload):
                with patch.object(
                    seguimientos,
                    "get_followup_payload",
                    return_value={
                        "item_autoevaluacion": ["", ""],
                        "item_eval_empresa": ["", ""],
                        "tipo_apoyo": "",
                    },
                ):
                    workflow = seguimientos.get_workflow_state({"source": "drive", "file_id": "demo"})

        self.assertEqual(workflow["next_followup"], 1)
        self.assertEqual(workflow["completed_followups"], [])
        self.assertEqual(workflow["editable_sheet"], f"{seguimientos.SHEET_PREFIX}1")
        self.assertEqual(
            workflow["message"],
            "Hoja base y seguimiento 1 habilitados hasta diligenciar el seguimiento 1.",
        )

    def test_workflow_advances_to_followup_2_only_after_followup_1_is_completed(self) -> None:
        base_payload = {
            "nombre_vinculado": "Persona Demo",
            "cedula": "123456",
            "cargo_vinculado": "Auxiliar",
            "discapacidad": "Fisica",
            "seguimiento_fechas_1_3": ["2026-04-01", "", ""],
            "seguimiento_fechas_4_6": ["", "", ""],
        }

        def _payload_for_index(_case_ref, index):
            if index == 1:
                return {
                    "item_autoevaluacion": ["Excelente"],
                    "item_eval_empresa": ["Bien"],
                    "tipo_apoyo": "Requiere apoyo bajo.",
                }
            return {
                "item_autoevaluacion": [""],
                "item_eval_empresa": [""],
                "tipo_apoyo": "",
            }

        with patch.object(
            seguimientos,
            "get_case_meta",
            return_value={"base_sheet_name": seguimientos.SHEET_BASE, "max_seguimientos": 3},
        ):
            with patch.object(seguimientos, "get_base_payload", return_value=base_payload):
                with patch.object(seguimientos, "get_followup_payload", side_effect=_payload_for_index):
                    workflow = seguimientos.get_workflow_state({"source": "drive", "file_id": "demo"})

        self.assertEqual(workflow["next_followup"], 2)
        self.assertEqual(workflow["completed_followups"], [1])
        self.assertEqual(workflow["editable_sheet"], f"{seguimientos.SHEET_PREFIX}2")


class SeguimientosBaseSheetMappingTests(unittest.TestCase):
    def test_build_base_sheet_updates_writes_apoyos_ajustes_to_e21(self) -> None:
        updates = seguimientos._build_base_sheet_updates(
            {"apoyos_ajustes": "Texto de ajuste"},
            base_sheet_name=seguimientos.SHEET_BASE,
        )

        target = next(
            item for item in updates if item["value"] == "Texto de ajuste"
        )

        self.assertEqual(target["range"], f"'{seguimientos.SHEET_BASE}'!E21")

    def test_local_roundtrip_keeps_label_in_a21_and_value_in_e21(self) -> None:
        repo_root = Path(__file__).resolve().parents[1]
        template_path = repo_root / "templates" / "seguimientos.xlsx"
        with tempfile.TemporaryDirectory() as tmp_dir:
            case_path = Path(tmp_dir) / "seguimiento_demo.xlsx"
            shutil.copy2(template_path, case_path)

            seguimientos.save_base_payload(
                str(case_path),
                {"apoyos_ajustes": "Ajuste razonable de prueba"},
            )

            workbook = load_workbook(case_path, data_only=False)
            try:
                sheet_name = seguimientos._get_base_sheet_name_from_workbook(workbook)
                ws = workbook[sheet_name]
                self.assertEqual(
                    str(ws["A21"].value or "").strip(),
                    "Apoyos y/o ajustes razonables requeridos:",
                )
                self.assertEqual(ws["E21"].value, "Ajuste razonable de prueba")
            finally:
                workbook.close()

            payload = seguimientos.get_base_payload(str(case_path))
            self.assertEqual(payload["apoyos_ajustes"], "Ajuste razonable de prueba")


if __name__ == "__main__":
    unittest.main()
