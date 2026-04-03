from __future__ import annotations

import unittest
from unittest.mock import patch

from formularios.condiciones_vacante import condiciones_vacante as cv


class _DummyRow:
    def __init__(self, worksheet: "_DummyWorksheet", row: int) -> None:
        self.worksheet = worksheet
        self.row = row

    def Insert(self) -> None:
        self.worksheet.insert_calls.append(self.row)

    def Copy(self, target: "_DummyRow") -> None:
        self.worksheet.copy_calls.append((self.row, target.row))


class _DummyWorksheet:
    def __init__(self) -> None:
        self.insert_calls: list[int] = []
        self.copy_calls: list[tuple[int, int]] = []

    def Rows(self, row: int) -> _DummyRow:
        return _DummyRow(self, row)


class CondicionesVacanteTemplateWriterTests(unittest.TestCase):
    def setUp(self) -> None:
        self._previous_cache = cv.FORM_CACHE.copy()
        cv.FORM_CACHE.clear()

    def tearDown(self) -> None:
        cv.FORM_CACHE.clear()
        cv.FORM_CACHE.update(self._previous_cache)

    def test_section_6_overflow_keeps_inserting_at_same_boundary_row(self) -> None:
        ws = _DummyWorksheet()
        payload = [{"discapacidad": f"D{i}"} for i in range(8)]

        with patch.object(cv, "ws_write") as write_mock, patch.object(cv, "_log_excel"):
            cv._write_section_with_ws(ws, "section_6", payload)

        self.assertEqual(ws.insert_calls, [157, 157, 157, 157])
        self.assertEqual(ws.copy_calls, [(156, 157)] * 4)
        written_cells = [call.args[1] for call in write_mock.call_args_list]
        self.assertEqual(
            written_cells,
            ["A153", "A154", "A155", "A156", "A157", "A158", "A159", "A160"],
        )

    def test_section_7_and_section_8_shift_after_section_6_overflow(self) -> None:
        ws = _DummyWorksheet()
        cv.FORM_CACHE["section_6"] = [{"discapacidad": f"D{i}"} for i in range(8)]
        section_7 = {"observaciones_recomendaciones": "Observacion"}
        section_8 = {
            "asistentes": [
                {"nombre": "Uno", "cargo": "Cargo 1"},
                {"nombre": "Dos", "cargo": "Cargo 2"},
                {"nombre": "Tres", "cargo": "Cargo 3"},
            ]
        }

        with patch.object(cv, "ws_write") as write_mock, patch.object(cv, "_log_excel"):
            cv._write_section_with_ws(ws, "section_7", section_7)
            cv._write_section_with_ws(ws, "section_8", section_8)

        writes = [(call.args[1], call.args[2]) for call in write_mock.call_args_list]
        self.assertEqual(
            writes,
            [
                ("A163", "Observacion"),
                ("E165", "Uno"),
                ("L165", "Cargo 1"),
                ("E166", "Dos"),
                ("L166", "Cargo 2"),
                ("E167", "Tres"),
                ("L167", "Cargo 3"),
            ],
        )

    def test_section_8_overflow_keeps_inserting_at_same_boundary_row(self) -> None:
        ws = _DummyWorksheet()
        section_8 = {
            "asistentes": [
                {"nombre": "Uno", "cargo": "Cargo 1"},
                {"nombre": "Dos", "cargo": "Cargo 2"},
                {"nombre": "Tres", "cargo": "Cargo 3"},
                {"nombre": "Cuatro", "cargo": "Cargo 4"},
                {"nombre": "Cinco", "cargo": "Cargo 5"},
            ]
        }

        with patch.object(cv, "ws_write"), patch.object(cv, "_log_excel"):
            cv._write_section_with_ws(ws, "section_8", section_8)

        self.assertEqual(ws.insert_calls, [164, 164])
        self.assertEqual(ws.copy_calls, [(163, 164)] * 2)

    def test_section_8_overflow_shifts_insert_boundary_after_section_6_growth(self) -> None:
        ws = _DummyWorksheet()
        cv.FORM_CACHE["section_6"] = [{"discapacidad": f"D{i}"} for i in range(8)]
        section_8 = {
            "asistentes": [
                {"nombre": "Uno", "cargo": "Cargo 1"},
                {"nombre": "Dos", "cargo": "Cargo 2"},
                {"nombre": "Tres", "cargo": "Cargo 3"},
                {"nombre": "Cuatro", "cargo": "Cargo 4"},
                {"nombre": "Cinco", "cargo": "Cargo 5"},
            ]
        }

        with patch.object(cv, "ws_write"), patch.object(cv, "_log_excel"):
            cv._write_section_with_ws(ws, "section_8", section_8)

        self.assertEqual(ws.insert_calls, [168, 168])
        self.assertEqual(ws.copy_calls, [(167, 168)] * 2)


if __name__ == "__main__":
    unittest.main()
