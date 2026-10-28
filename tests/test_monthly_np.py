from __future__ import annotations

import datetime as dt
import sys
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
sys.path.insert(0, str(SRC_DIR))

from monthly_np import (  # noqa: E402
    calculate_monthly_rows,
    load_previous_month_snapshot,
    make_snapshot,
    parse_report_date,
    save_snapshot,
)
from utils import monthly_np_storage_path  # noqa: E402


class MonthlyNpTests(unittest.TestCase):
    def test_parses_report_date_before_processed_suffixes(self) -> None:
        path = Path("report_01_09_2026__processed_20260901_1442.xlsx")
        self.assertEqual(parse_report_date(path), dt.date(2026, 9, 1))

    def test_september_uses_august_end_and_calculates_only_new_usage(self) -> None:
        baseline = {
            "snapshot_date": "2026-08-31",
            "used_np_rows": {
                "тп «Тажен»": 1832,
                "Навешивания на территории РК": 2138,
                "Итого": 3970,
            },
        }
        current = [
            {"Пункт": "тп «Тажен»", "Количество": 1848},
            {"Пункт": "тп «Болашак»", "Количество": "-"},
            {"Пункт": "Навешивания на территории РК", "Количество": 2146},
            {"Пункт": "Итого", "Количество": 3994},
        ]

        rows = calculate_monthly_rows(current, baseline)
        values = {row["Пункт"]: row["Количество"] for row in rows}

        self.assertEqual(values["тп «Тажен»"], 16)
        self.assertEqual(values["тп «Болашак»"], "-")
        self.assertEqual(values["Навешивания на территории РК"], 8)
        self.assertEqual(values["Итого"], 24)

    def test_new_month_uses_exact_or_latest_previous_month_snapshot(self) -> None:
        with tempfile.TemporaryDirectory() as runtime_raw, tempfile.TemporaryDirectory() as bundled_raw:
            runtime_dir = Path(runtime_raw)
            bundled_dir = Path(bundled_raw)

            august_30 = make_snapshot([], dt.date(2026, 8, 30), "report_30_08.xlsx")
            save_snapshot(runtime_dir, august_30)
            loaded, exact = load_previous_month_snapshot(
                dt.date(2026, 9, 2), runtime_dir, bundled_dir
            )
            self.assertEqual(loaded["snapshot_date"], "2026-08-30")
            self.assertFalse(exact)

            august_31 = make_snapshot([], dt.date(2026, 8, 31), "report_31_08.xlsx")
            save_snapshot(runtime_dir, august_31)
            loaded, exact = load_previous_month_snapshot(
                dt.date(2026, 9, 2), runtime_dir, bundled_dir
            )
            self.assertEqual(loaded["snapshot_date"], "2026-08-31")
            self.assertTrue(exact)

    def test_october_uses_embedded_september_30_baseline(self) -> None:
        bundled_dir = SRC_DIR / "data"
        with tempfile.TemporaryDirectory() as runtime_raw:
            loaded, exact = load_previous_month_snapshot(
                dt.date(2026, 10, 7), Path(runtime_raw), bundled_dir
            )

        self.assertIsNotNone(loaded)
        self.assertEqual(loaded["snapshot_date"], "2026-09-30")
        self.assertEqual(loaded["used_np_rows"]["Итого"], 55564)
        self.assertEqual(loaded["_baseline_storage"], "embedded")
        self.assertTrue(exact)

    def test_damaged_user_baseline_falls_back_to_embedded_copy(self) -> None:
        bundled_dir = SRC_DIR / "data"
        with tempfile.TemporaryDirectory() as runtime_raw:
            runtime_dir = Path(runtime_raw)
            (runtime_dir / "np_snapshot_20260930.json").write_text(
                "not valid json", encoding="utf-8"
            )
            loaded, exact = load_previous_month_snapshot(
                dt.date(2026, 10, 7), runtime_dir, bundled_dir
            )

        self.assertIsNotNone(loaded)
        self.assertEqual(loaded["used_np_rows"]["Итого"], 55564)
        self.assertEqual(loaded["_baseline_storage"], "embedded")
        self.assertTrue(exact)

    def test_snapshot_storage_is_next_to_program_folders(self) -> None:
        with tempfile.TemporaryDirectory() as program_raw:
            program_dir = Path(program_raw)
            with patch("utils.app_base_path", return_value=program_dir):
                storage = monthly_np_storage_path()

            self.assertEqual(storage, program_dir / "monthly_baselines")
            self.assertTrue(storage.is_dir())


if __name__ == "__main__":
    unittest.main()
