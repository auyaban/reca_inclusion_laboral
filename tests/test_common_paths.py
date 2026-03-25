from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from formularios import common


class CommonPathTests(unittest.TestCase):
    def test_sanitize_filename_handles_reserved_windows_names(self) -> None:
        self.assertEqual(common._sanitize_filename("CON.", default="Empresa"), "CON_")
        self.assertEqual(common._sanitize_filename("aux.txt", default="Empresa"), "aux_.txt")

    def test_build_process_output_path_falls_back_when_desktop_candidate_is_too_long(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            long_desktop = str(Path(tmpdir) / ("desktop_" + ("x" * 240)))
            local_app_data = str(Path(tmpdir) / "localappdata")
            with patch.object(common, "_get_desktop_dir", return_value=long_desktop):
                with patch.dict(os.environ, {"LOCALAPPDATA": local_app_data}, clear=False):
                    output_path = common._build_process_output_path(
                        "Empresa Demo",
                        "Proceso de Seleccion Incluyente",
                    )

        self.assertTrue(str(output_path).startswith(local_app_data))


if __name__ == "__main__":
    unittest.main()
