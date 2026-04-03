import unittest
from unittest.mock import patch

import app


class _DummyHub:
    def __init__(self, login):
        self.current_user_profile = {"usuario_login": login}
        self.current_user = login


class _DummyWindow:
    def __init__(self, login, form_id="seleccion_incluyente"):
        self._form_id = form_id
        self.master = _DummyHub(login)


class TestFillAccessTests(unittest.TestCase):
    def test_login_allows_test_fill_only_for_testaaron(self):
        with patch.dict(
            "os.environ",
            {"RECA_ENABLE_TEST_FILL": "1", "RECA_TEST_FILL_LOGIN": "testaaron"},
            clear=False,
        ):
            self.assertTrue(app._login_allows_test_fill("testaaron"))
            self.assertTrue(app._login_allows_test_fill("TESTAARON"))
            self.assertFalse(app._login_allows_test_fill("aarontest"))
            self.assertFalse(app._login_allows_test_fill("aaron"))

    def test_window_allows_test_fill_only_for_supported_form_and_login(self):
        with patch.dict(
            "os.environ",
            {"RECA_ENABLE_TEST_FILL": "1", "RECA_TEST_FILL_LOGIN": "testaaron"},
            clear=False,
        ):
            self.assertTrue(app._window_allows_test_fill(_DummyWindow("testaaron")))
            self.assertFalse(app._window_allows_test_fill(_DummyWindow("otro_usuario")))
            self.assertFalse(app._window_allows_test_fill(_DummyWindow("testaaron", form_id="")))


if __name__ == "__main__":
    unittest.main()
