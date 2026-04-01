from __future__ import annotations

import unittest

from formularios import common


class CommonNumberFormattingTests(unittest.TestCase):
    def test_coerce_excel_decimal_value_keeps_general_numbers_unscaled(self) -> None:
        self.assertEqual(common._coerce_excel_decimal_value("22.25"), 22.25)
        self.assertEqual(common._coerce_excel_decimal_value("1"), 1.0)

    def test_coerce_excel_decimal_value_scales_percent_formatted_numbers(self) -> None:
        self.assertEqual(common._coerce_excel_decimal_value("22.25", number_format="0%"), 0.2225)
        self.assertEqual(common._coerce_excel_decimal_value("1", number_format="0%"), 0.01)

    def test_coerce_excel_decimal_value_accepts_comma_decimal_for_percent(self) -> None:
        self.assertEqual(common._coerce_excel_decimal_value("18,7", number_format="0%"), 0.187)
        self.assertEqual(common._coerce_excel_decimal_value("25%", number_format="0%"), 0.25)


if __name__ == "__main__":
    unittest.main()
