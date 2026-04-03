from __future__ import annotations

import unittest
from pathlib import Path

import app


class EncodingGuardTests(unittest.TestCase):
    def test_repo_python_sources_do_not_contain_mojibake(self) -> None:
        root = Path(__file__).resolve().parents[1]

        issues = app._detect_mojibake_issues(str(root))

        self.assertEqual(issues, [])

    def test_editorconfig_enforces_utf8(self) -> None:
        editorconfig = Path(__file__).resolve().parents[1] / ".editorconfig"

        self.assertTrue(editorconfig.exists())
        content = editorconfig.read_text(encoding="utf-8")
        self.assertIn("charset = utf-8", content)


if __name__ == "__main__":
    unittest.main()
