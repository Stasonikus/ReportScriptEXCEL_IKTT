from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
sys.path.insert(0, str(SRC_DIR))

from stage2 import apply_stage2_baseline  # noqa: E402


class Stage2BaselineTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        baseline_path = SRC_DIR / "data" / "stage2_baseline_20260831.json"
        cls.baseline = json.loads(baseline_path.read_text(encoding="utf-8"))

    def test_table1_subtracts_dashes_as_zero_and_recalculates_totals(self) -> None:
        tables = [
            {
                "table": "Таблица1",
                "rail_registered_total": 209,
                "rows": [
                    {
                        "ТП/ЖД": "Алаколь",
                        "РФ": 2209,
                        "РБ": 2128,
                        "КР": 6,
                        "РА": "-",
                        "Итого": 4343,
                    },
                    {
                        "ТП/ЖД": "Итого",
                        "РФ": 2209,
                        "РБ": 2128,
                        "КР": 6,
                        "РА": "-",
                        "Итого": 4343,
                    },
                ],
            }
        ]

        result = apply_stage2_baseline(tables, self.baseline)[0]

        self.assertEqual(
            result["rows"][0],
            {
                "ТП/ЖД": "Алаколь",
                "РФ": 477,
                "РБ": 236,
                "КР": 3,
                "РА": 0,
                "Итого": 716,
            },
        )
        self.assertEqual(result["rows"][-1]["Итого"], 716)
        self.assertEqual(result["rail_registered_total"], 209)

    def test_table4_keeps_stock_snapshot_and_subtracts_only_used_np(self) -> None:
        tables = [
            {
                "table": "Таблица5",
                "rows": [
                    {"Точка": "тп «Тажен»", "Количество НП": 145},
                ],
                "used_np_rows": [
                    {"Пункт": "тп «Тажен»", "Количество": 2255},
                    {"Пункт": "тп «Болашак»", "Количество": "-"},
                    {
                        "Пункт": "Навешивания на территории РК",
                        "Количество": 2156,
                    },
                    {"Пункт": "Итого", "Количество": 4411},
                ],
            }
        ]

        result = apply_stage2_baseline(tables, self.baseline)[0]

        self.assertEqual(result["rows"][0]["Количество НП"], 145)
        self.assertEqual(result["used_np_rows"][0]["Количество"], 423)
        self.assertEqual(result["used_np_rows"][1]["Количество"], 0)
        self.assertEqual(result["used_np_rows"][2]["Количество"], 18)
        self.assertEqual(result["used_np_rows"][3]["Количество"], 441)
        self.assertEqual(result["used_np_total"], 441)

    def test_table2_matches_manual_stage2_example(self) -> None:
        def section(values: list[int]) -> list[dict[str, object]]:
            names = [
                "Российская Федерация",
                "Республика Беларусь",
                "Кыргызская Республика",
                "Республика Армения",
                "Итого",
            ]
            return [
                {"№": index, "Страна": name, "Количество перевозок": value}
                for index, (name, value) in enumerate(zip(names, values), start=1)
            ]

        table = {
            "table": "Таблица2",
            "export_rows": section([358, 8, 6, 0, 372]),
            "trade_rows": section([569, 37, 473, 0, 1079]),
        }

        result = apply_stage2_baseline([table], self.baseline)[0]

        self.assertEqual(
            [row["Количество перевозок"] for row in result["export_rows"]],
            [2, 0, 0, 0, 2],
        )
        self.assertEqual(
            [row["Количество перевозок"] for row in result["trade_rows"]],
            [5, 0, 2, 0, 7],
        )

    def test_new_table3_is_already_stage2_direct_and_is_not_subtracted(self) -> None:
        table = {
            "table": "Таблица3",
            "stage": "stage2-direct",
            "rows": [
                {
                    "Тип транспорта": "Авто",
                    "РБ - КР": 10,
                    "РФ - КР": 20,
                    "КР - РФ": 30,
                    "КР - РБ": 40,
                    "Итого": 100,
                }
            ],
        }

        self.assertEqual(apply_stage2_baseline([table], self.baseline)[0], table)

    def test_table3_matches_manual_stage2_example(self) -> None:
        def section(names: list[str], values: list[int]) -> list[dict[str, object]]:
            return [
                {"Направление": name, "Количество": value}
                for name, value in zip(names + ["Итого"], values)
            ]

        out_names = [
            "В Российскую Федерацию",
            "В Республику Беларусь",
            "В Кыргызскую Республику",
            "В Республику Армения",
        ]
        in_names = [
            "Из Российской Федерации",
            "Из Республики Беларусь",
            "Из Кыргызской Республики",
            "Из Республики Армения",
        ]
        table = {
            "table": "Таблица4",
            "auto_out_rows": section(out_names, [14634, 4507, 841, 0, 19982]),
            "auto_in_rows": section(in_names, [11710, 12803, 204, 0, 24717]),
            "rail_out_rows": section(out_names, [0, 2, 1, 0, 3]),
            "rail_in_rows": section(in_names, [12, 42, 0, 0, 54]),
        }

        result = apply_stage2_baseline([table], self.baseline)[0]

        self.assertEqual(result["auto_out_total"], 191)
        self.assertEqual(result["auto_in_total"], 216)
        self.assertEqual(result["rail_out_total"], 3)
        self.assertEqual(result["rail_in_total"], 54)

    def test_table3_keeps_direct_stage2_rail_without_second_subtraction(self) -> None:
        out_names = [
            "В Российскую Федерацию", "В Республику Беларусь",
            "В Кыргызскую Республику", "В Республику Армения",
        ]
        in_names = [
            "Из Российской Федерации", "Из Республики Беларусь",
            "Из Кыргызской Республики", "Из Республики Армения",
        ]

        def section(names: list[str], values: list[int]) -> list[dict[str, object]]:
            return [{"Направление": name, "Количество": value}
                    for name, value in zip(names + ["Итого"], values)]

        table = {
            "table": "Таблица4",
            "rail_values_are_stage2_direct": True,
            "auto_out_rows": section(out_names, [15136, 4532, 892, 0, 20560]),
            "auto_in_rows": section(in_names, [13653, 12996, 210, 0, 26859]),
            "rail_out_rows": section(out_names, [77, 128, 2, 0, 207]),
            "rail_in_rows": section(in_names, [74, 366, 0, 0, 440]),
        }

        result = apply_stage2_baseline([table], self.baseline)[0]

        self.assertEqual(result["rail_out_total"], 207)
        self.assertEqual(result["rail_in_total"], 440)
        self.assertEqual(result["rail_total"], 647)
        self.assertEqual(result["auto_out_total"], 769)
        self.assertEqual(result["auto_in_total"], 2358)

    def test_table5_keeps_current_values_without_subtraction(self) -> None:
        rows = [
            {"Страна начала перевозки": "Российская Федерация", "Количество перевозок": 786},
            {"Страна начала перевозки": "Республика Беларусь", "Количество перевозок": 703},
            {"Страна начала перевозки": "Кыргызская Республика", "Количество перевозок": 3},
            {"Страна начала перевозки": "Республика Армения", "Количество перевозок": 0},
            {"Страна начала перевозки": "Итого", "Количество перевозок": 1492},
        ]

        table = {"table": "Таблица6", "rows": rows, "total": 1492}

        result = apply_stage2_baseline([table], self.baseline)[0]

        self.assertEqual(result, table)


if __name__ == "__main__":
    unittest.main()
