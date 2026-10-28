from __future__ import annotations

import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from config import TP_ZHD_LIST_T1_OLD  # noqa: E402


class LegacyTable1LayoutTests(unittest.TestCase):
    def test_old_method_uses_the_approved_word_order(self) -> None:
        self.assertEqual(TP_ZHD_LIST_T1_OLD, [
            "Алаколь", "Казыгурт", "Нур жолы", "Бахты", "Калжат",
            "Тажен", "Атамекен", "Б. конысбаева", "Майкапчагай",
            "Темир-баба", "Капланбек", "Морпорт актау", "Курык порт",
            "Курык жд", "Болашак", "Оазис", "Мактаарал", "Сарыагаш",
            "Алтынколь", "Темиржол",
        ])


if __name__ == "__main__":
    unittest.main()
