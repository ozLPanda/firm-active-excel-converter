from __future__ import annotations

import unittest

from converter import name_without_availability_marker


class ConverterNameNormalizationTests(unittest.TestCase):
    def test_removes_unavailable_marker_from_product_name_case_insensitive(self) -> None:
        source = (
            'НЕТ В НАЛИЧИИ Группа безопасности компакт 3 бар 1" (Д 25) '
            "(AQUALINK АК)( арт-1158) установку данного узла производится только "
            "высококвалифицированным специалистом, уточните у продавца рекомендации по эксплуатации"
        )
        expected = (
            'Группа безопасности компакт 3 бар 1" (Д 25) '
            "(AQUALINK АК)( арт-1158) установку данного узла производится только "
            "высококвалифицированным специалистом, уточните у продавца рекомендации по эксплуатации"
        )

        cleaned, marker_found = name_without_availability_marker(source)

        self.assertTrue(marker_found)
        self.assertEqual(cleaned, expected)

    def test_removes_unavailable_marker_with_mixed_case(self) -> None:
        cleaned, marker_found = name_without_availability_marker("Нет в Наличии Товар")

        self.assertTrue(marker_found)
        self.assertEqual(cleaned, "Товар")


if __name__ == "__main__":
    unittest.main()
