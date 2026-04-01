from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

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


if __name__ == "__main__":
    unittest.main()
