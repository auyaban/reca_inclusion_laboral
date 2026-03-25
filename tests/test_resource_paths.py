from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from formularios.contratacion_incluyente import contratacion_incluyente as contratacion
from formularios.seleccion_incluyente import seleccion_incluyente as seleccion
from formularios.seguimientos import seguimientos
from version_info import resource_path


class ResourcePathTests(unittest.TestCase):
    def test_resource_path_prefers_meipass_when_present(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            bundle_root = Path(tmpdir)
            with patch("version_info.sys._MEIPASS", str(bundle_root), create=True):
                self.assertEqual(
                    resource_path("templates/demo.xlsx"),
                    bundle_root / "templates" / "demo.xlsx",
                )

    def test_contratacion_template_path_uses_bundle_templates(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            bundle_root = Path(tmpdir)
            templates_dir = bundle_root / "templates"
            templates_dir.mkdir(parents=True, exist_ok=True)
            expected = templates_dir / "contratacion_incluyente_grupal_2_4.xlsx"
            expected.write_bytes(b"")

            with patch("version_info.sys._MEIPASS", str(bundle_root), create=True):
                resolved = contratacion._find_template_path(
                    contratacion.TEMPLATE_VARIANT_GROUP_2_PLUS
                )

        self.assertEqual(Path(resolved), expected)

    def test_seleccion_template_path_matches_normalized_bundle_name(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            bundle_root = Path(tmpdir)
            templates_dir = bundle_root / "templates"
            templates_dir.mkdir(parents=True, exist_ok=True)
            expected = templates_dir / "seleccion incluyente grupal 2-4.xlsx"
            expected.write_bytes(b"")

            with patch("version_info.sys._MEIPASS", str(bundle_root), create=True):
                resolved = seleccion._find_template_path(
                    seleccion.TEMPLATE_VARIANT_GROUP_2_PLUS
                )

        self.assertEqual(Path(resolved), expected)

    def test_seguimientos_template_path_uses_bundle_templates(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            bundle_root = Path(tmpdir)
            templates_dir = bundle_root / "templates"
            templates_dir.mkdir(parents=True, exist_ok=True)
            expected = templates_dir / "seguimientos.xlsx"
            expected.write_bytes(b"")

            with patch("version_info.sys._MEIPASS", str(bundle_root), create=True):
                resolved = seguimientos._find_template_path()

        self.assertEqual(Path(resolved), expected)


if __name__ == "__main__":
    unittest.main()
