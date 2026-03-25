from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import mock_open, patch

from formularios.condiciones_vacante import condiciones_vacante as cv


class CondicionesVacanteDictionaryTests(unittest.TestCase):
    def setUp(self) -> None:
        self._previous_cache = cv._DISABILITY_DICT
        cv._DISABILITY_DICT = None

    def tearDown(self) -> None:
        cv._DISABILITY_DICT = self._previous_cache

    def test_text_dictionary_loader_parses_tea_heading_as_own_entry(self) -> None:
        sample = """
DISCAPACIDAD INTELECTUAL
"1. Instrucción uno
2. Instrucción dos"

TEA / AUTISMO
"1. Anticipar tareas
2. Anticipar cambios"
""".strip()

        with patch("formularios.condiciones_vacante.condiciones_vacante.os.path.exists", return_value=True):
            with patch("builtins.open", mock_open(read_data=sample)):
                entries = cv._load_disability_descriptions_from_text()

        self.assertIn(cv.normalize_disability_key("TEA / AUTISMO"), entries)
        self.assertIn("Anticipar tareas", entries[cv.normalize_disability_key("TEA / AUTISMO")])
        self.assertNotIn("TEA / AUTISMO", entries[cv.normalize_disability_key("DISCAPACIDAD INTELECTUAL")])

    def test_get_disability_descriptions_prefers_sheet_source(self) -> None:
        expected = {cv.normalize_disability_key("TEA / AUTISMO"): "Texto oficial"}

        with patch.object(cv, "_load_disability_descriptions_from_sheet", return_value=expected):
            with patch.object(cv, "_load_disability_descriptions_from_text") as text_loader:
                result = cv.get_disability_descriptions()

        self.assertEqual(result, expected)
        text_loader.assert_not_called()

    def test_get_disability_descriptions_uses_text_fallback_if_sheet_fails(self) -> None:
        expected = {cv.normalize_disability_key("TEA / AUTISMO"): "Texto fallback"}

        with patch.object(cv, "_load_disability_descriptions_from_sheet", side_effect=RuntimeError("sin acceso")):
            with patch.object(cv, "_load_disability_descriptions_from_text", return_value=expected):
                result = cv.get_disability_descriptions()

        self.assertEqual(result, expected)

    def test_text_dictionary_loader_uses_bundle_path_when_meipass_is_set(self) -> None:
        sample = """
TEA / AUTISMO
"1. Anticipar tareas
2. Anticipar cambios"
""".strip()

        with tempfile.TemporaryDirectory() as tmpdir:
            bundle_root = Path(tmpdir)
            (bundle_root / "Diccionario.txt").write_text(sample, encoding="utf-8")

            with patch("version_info.sys._MEIPASS", str(bundle_root), create=True):
                entries = cv._load_disability_descriptions_from_text()

        self.assertIn(cv.normalize_disability_key("TEA / AUTISMO"), entries)
        self.assertIn("Anticipar tareas", entries[cv.normalize_disability_key("TEA / AUTISMO")])


if __name__ == "__main__":
    unittest.main()
