from __future__ import annotations

import unittest
from unittest.mock import patch

import app


class _DummyHub:
    def __init__(self, login: str) -> None:
        self.current_user_profile = {"usuario_login": login}
        self.current_user = login


class _DummyWindow:
    def __init__(self, login: str, form_id: str = "presentacion_programa") -> None:
        self.master = _DummyHub(login)
        self._form_id = form_id


class TestFillHelpersTests(unittest.TestCase):
    def test_login_allows_test_fill_only_for_testaaron(self) -> None:
        with patch.dict(
            "os.environ",
            {"RECA_ENABLE_TEST_FILL": "1", "RECA_TEST_FILL_LOGIN": "testaaron"},
            clear=False,
        ):
            self.assertTrue(app._login_allows_test_fill("testaaron"))
            self.assertTrue(app._login_allows_test_fill(" TestAaron "))
            self.assertFalse(app._login_allows_test_fill("otro.usuario"))

    def test_pick_test_combobox_value_uses_first_non_empty_value(self) -> None:
        self.assertEqual(app._pick_test_combobox_value(["", "Primera", "Segunda"]), "Primera")
        self.assertEqual(app._pick_test_combobox_value(["Unica"]), "Unica")
        self.assertEqual(app._pick_test_combobox_value([]), "")

    def test_get_test_fill_entry_value_returns_minimal_values_by_kind(self) -> None:
        self.assertEqual(app._get_test_fill_entry_value("numeric", max_len=6), "1111")
        self.assertEqual(app._get_test_fill_entry_value("decimal"), "1")
        self.assertEqual(app._get_test_fill_entry_value("birthdate"), "01/01/2000")
        self.assertEqual(app._get_test_fill_entry_value("name"), "Pendiente")
        self.assertEqual(app._get_test_fill_entry_value(""), "Pendiente")

    def test_get_test_fill_command_requires_allowed_login(self) -> None:
        with patch.dict(
            "os.environ",
            {"RECA_ENABLE_TEST_FILL": "1", "RECA_TEST_FILL_LOGIN": "testaaron"},
            clear=False,
        ):
            allowed = app._get_test_fill_command(_DummyWindow("testaaron"))
            blocked = app._get_test_fill_command(_DummyWindow("otro.usuario"))

        self.assertTrue(callable(allowed))
        self.assertIsNone(blocked)


if __name__ == "__main__":
    unittest.main()
