from __future__ import annotations

import sys
import unittest
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from builders import build_rail_registered_count, build_table3
from config import load_config


class RailOwnerAliasTests(unittest.TestCase):
    def rail_source(self, owner):
        ws = Workbook().active
        ws.append(["Дата", "Владелец записи", "Статус", "company_id",
                   "Start waypoint", "End waypoint"])
        for date, status, company, end in [
            ("2026-09-10", "finished", 10, "Себеж"),
            ("2026-09-11", "finished", 10, "БРЕСТ-СЕВЕРНЫЙ"),
            ("2026-09-12", "activated", 10, "Себеж"),
            ("2026-09-13", "finished", 1, "Себеж"),
            ("2026-09-14", "canceled", 10, "Себеж"),
            ("2026-08-30", "finished", 10, "Себеж"),
        ]:
            ws.append([date, owner, status, company, 'Т/П "ТЕМИРЖОЛ"', end])
        return {"ws": ws}

    def report_source(self):
        ws = Workbook().active
        ws.append(["Показатель", "Значение"])
        ws.append(["From", "To", "Деактивированная перевозка", "Завершенная перевозка"])
        ws.append(["KZ", "RU", 0, 10])
        ws.append(["Перевозки по типам кортежей"])
        return {"ws": ws}

    def test_old_and_new_owner_names_preserve_existing_filters(self):
        base = "ТОО «Институт космической техники и технологий»"
        for owner in (base, base + " (оператор)",
                      'ТОО "Институт космической техники и технологий" (оператор)'):
            with self.subTest(owner=owner):
                rail = self.rail_source(owner)
                # Registration includes activated, but excludes pre-stage dates.
                self.assertEqual(build_rail_registered_count(rail), 3)
                table = build_table3(self.report_source(), load_config().normalization, rail)
                # Таблица3 учитывает только завершенные ЖД-перевозки второго этапа.
                self.assertEqual([r["Количество"] for r in table["rail_out_rows"]],
                                 [1, 1, 0, 0, 2])
                self.assertEqual(table["rail_in_total"], 0)

    def test_unrelated_owner_is_not_admitted(self):
        rail = self.rail_source("Другой владелец (оператор)")
        self.assertEqual(build_rail_registered_count(rail), 0)
        table = build_table3(self.report_source(), load_config().normalization, rail)
        self.assertEqual(table["rail_out_total"], 0)


if __name__ == "__main__":
    unittest.main()
