from __future__ import annotations

import sys
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
sys.path.insert(0, str(SRC_DIR))

from builders import build_rail_registered_count  # noqa: E402


class RailRegisteredCountTests(unittest.TestCase):
    def test_counts_only_stage2_dates_with_existing_filters(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.append(["Дата", "Владелец записи", "Статус", "company_id"])
        ws.append([datetime(2026, 8, 31, 23, 59, 59), "КЕДЕН KEDEN", "activated", 10])
        ws.append([datetime(2026, 9, 1, 0, 0, 0), "КЕДЕН KEDEN", "activated", 10])
        ws.append(["2026-09-04 12:30:00", "ЭСФ ESF", "finished", 20])
        ws.append(["2026-09-04 12:30:00", "Другой владелец", "activated", 20])
        ws.append(["2026-09-04 12:30:00", "КЕДЕН KEDEN", "preinfo", 20])
        ws.append(["2026-09-04 12:30:00", "КЕДЕН KEDEN", "activated", 1])

        result = build_rail_registered_count({"ws": ws})

        self.assertEqual(result, 2)

    def test_requires_date_column(self) -> None:
        wb = Workbook()
        ws = wb.active
        ws.append(["Владелец записи", "Статус", "company_id"])

        self.assertIsNone(build_rail_registered_count({"ws": ws}))


if __name__ == "__main__":
    unittest.main()
