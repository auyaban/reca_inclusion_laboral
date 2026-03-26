import tempfile
import unittest
from pathlib import Path
from unittest import mock

import version_info


class VersionInfoTests(unittest.TestCase):
    def test_get_version_accepts_utf8_bom(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            version_path = Path(tmpdir) / "VERSION"
            version_path.write_bytes(b"\xef\xbb\xbf1.2.4\n")
            with mock.patch.object(version_info, "resource_path", return_value=version_path):
                self.assertEqual(version_info.get_version(), "1.2.4")


if __name__ == "__main__":
    unittest.main()
