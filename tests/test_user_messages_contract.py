from __future__ import annotations

import re
import unittest
from pathlib import Path

from formularios import user_messages


CANONICAL_CONTEXTS = (
    "login",
    "company_search",
    "linked_user_search",
    "autofill_user",
    "database_refresh",
    "sync",
    "case_open",
    "followup_case",
    "section_confirm",
    "finalization",
    "save_sheet",
    "ui_error",
)


class UserMessageContractTests(unittest.TestCase):
    def test_supported_contexts_match_the_canonical_catalog(self) -> None:
        self.assertEqual(user_messages.SUPPORTED_USER_MESSAGE_CONTEXTS, CANONICAL_CONTEXTS)

    def test_all_log_user_error_contexts_are_supported(self) -> None:
        source = Path(__file__).resolve().parents[1].joinpath("app.py").read_text(
            encoding="utf-8",
            errors="replace",
        )
        used_contexts = set(re.findall(r'_log_user_error\("([^"]+)"', source))

        self.assertEqual(used_contexts - set(CANONICAL_CONTEXTS), set())

    def test_database_refresh_returns_specific_copy(self) -> None:
        message = user_messages.map_exception_to_user_message(
            "database_refresh",
            RuntimeError("Supabase unavailable"),
        )

        self.assertIn("base de datos", message.lower())

    def test_linked_user_search_returns_specific_permission_copy(self) -> None:
        message = user_messages.map_exception_to_user_message(
            "linked_user_search",
            PermissionError("permission denied"),
        )

        self.assertIn("vinculado", message.lower())
        self.assertIn("permisos", message.lower())

    def test_ui_error_returns_controlled_default_copy(self) -> None:
        message = user_messages.map_exception_to_user_message("ui_error", RuntimeError("boom"))

        self.assertEqual(message, user_messages._default_message("ui_error"))

    def test_login_permission_denied_returns_controlled_copy(self) -> None:
        message = user_messages.map_exception_to_user_message(
            "login",
            RuntimeError("No fue posible validar tu perfil con los permisos actuales."),
        )

        self.assertIn("perfil", message.lower())


if __name__ == "__main__":
    unittest.main()
