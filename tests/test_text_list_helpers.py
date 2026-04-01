from __future__ import annotations

import unittest

import app
import text_review


class AppTextListHelperTests(unittest.TestCase):
    def test_get_text_list_continuation_for_dash_bullet(self) -> None:
        self.assertEqual(app._get_text_list_continuation("- Primer punto"), "  - ")

    def test_get_text_list_continuation_for_numbered_list(self) -> None:
        self.assertEqual(app._get_text_list_continuation("3. Tercer punto"), "  4. ")
        self.assertEqual(app._get_text_list_continuation("2) Segundo punto"), "  3) ")

    def test_get_text_list_continuation_returns_none_for_plain_text(self) -> None:
        self.assertIsNone(app._get_text_list_continuation("Esto es un parrafo normal."))

    def test_normalize_text_list_blocks_promotes_intro_plus_short_lines(self) -> None:
        text = (
            "A continuacion, algunas estrategias utiles:\n\n"
            "Establecer prioridades claras\n"
            "dividir tareas grandes en pasos pequeños\n"
            "evitar distracciones innecesarias\n"
            "tomar descansos regulares\n\n"
            "Aplicar estos habitos mejora el enfoque."
        )

        self.assertEqual(
            app._normalize_text_list_blocks(text),
            (
                "A continuacion, algunas estrategias utiles:\n\n"
                "  - Establecer prioridades claras\n"
                "  - dividir tareas grandes en pasos pequeños\n"
                "  - evitar distracciones innecesarias\n"
                "  - tomar descansos regulares\n\n"
                "Aplicar estos habitos mejora el enfoque."
            ),
        )


class TextReviewListFormattingTests(unittest.TestCase):
    def test_maybe_format_reviewed_list_from_semicolons(self) -> None:
        text = "Ajuste visual; ajuste auditivo; ajuste cognitivo"
        reviewed = text_review._maybe_format_reviewed_list(text, text)

        self.assertEqual(
            reviewed,
            "  - Ajuste visual\n  - ajuste auditivo\n  - ajuste cognitivo",
        )

    def test_maybe_format_reviewed_list_from_inline_numbers(self) -> None:
        text = "1. Primer paso 2. Segundo paso 3. Tercer paso"
        reviewed = text_review._maybe_format_reviewed_list(text, text)

        self.assertEqual(
            reviewed,
            "  - Primer paso\n  - Segundo paso\n  - Tercer paso",
        )

    def test_maybe_format_reviewed_list_for_intro_plus_plain_lines_block(self) -> None:
        text = (
            "A continuacion, algunas estrategias utiles:\n\n"
            "Establecer prioridades claras\n"
            "dividir tareas grandes en pasos pequeños\n"
            "evitar distracciones innecesarias\n"
            "tomar descansos regulares\n\n"
            "Aplicar estos habitos puede mejorar el enfoque."
        )

        self.assertEqual(
            text_review._maybe_format_reviewed_list(text, text),
            (
                "A continuacion, algunas estrategias utiles:\n\n"
                "  - Establecer prioridades claras\n"
                "  - dividir tareas grandes en pasos pequeños\n"
                "  - evitar distracciones innecesarias\n"
                "  - tomar descansos regulares\n\n"
                "Aplicar estos habitos puede mejorar el enfoque."
            ),
        )

    def test_maybe_format_reviewed_list_preserves_regular_paragraph(self) -> None:
        text = "La empresa cuenta con accesos adecuados y el personal fue sensibilizado."

        self.assertEqual(text_review._maybe_format_reviewed_list(text, text), text)


if __name__ == "__main__":
    unittest.main()
