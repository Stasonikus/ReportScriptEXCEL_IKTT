from __future__ import annotations

import datetime as dt
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from xml.sax.saxutils import escape

from openpyxl import Workbook, load_workbook


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from big_export import (
    REQUIRED_HEADERS,
    _transport,
    build_rail_export_and_trade,
    build_rail_transit_by_point,
    build_tables_from_big_export,
    detect_transport_by_vehicle_numbers,
    is_big_export_file,
    load_t1_transport_groups,
)
from config import load_config
from io_excel import write_output
from main import (
    _choose_method,
    _make_full_period_output_filename,
    _make_output_filename,
    _merge_hybrid_tables,
    _write_report,
)
from power_bi import write_power_bi_output


class BigExportTests(unittest.TestCase):
    def test_power_bi_output_contains_flat_facts_and_reference_tables(self) -> None:
        tables = [
            {
                "table": "Таблица1", "stage": "stage2-direct", "transport_split": True,
                "auto_rows": [{"ТП/ЖД": "Алаколь", "РФ": 3, "РБ": 0, "КР": 0, "РА": 0, "Итого": 3}, {"ТП/ЖД": "Итого", "РФ": 3, "РБ": 0, "КР": 0, "РА": 0, "Итого": 3}],
                "rail_rows": [{"ТП/ЖД": "Темиржол", "РФ": 0, "РБ": 2, "КР": 0, "РА": 0, "Итого": 2}, {"ТП/ЖД": "Итого", "РФ": 0, "РБ": 2, "КР": 0, "РА": 0, "Итого": 2}],
            },
            {
                "table": "Таблица2", "stage": "stage2-direct", "transport_split": True,
                "export_split_rows": [{"№": 1, "Страна": "Российская Федерация", "Авто": 4, "ЖД": 1, "Итого": 5}] * 4,
                "trade_split_rows": [{"№": 1, "Страна": "Республика Беларусь", "Авто": 2, "ЖД": 3, "Итого": 5}] * 4,
            },
        ]
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "power-bi.xlsx"
            write_power_bi_output(
                output, tables, dt.date(2026, 9, 7), "4 - Смешанный метод",
                ROOT / "src" / "data" / "border_crossing_points.json",
            )
            workbook = load_workbook(output, data_only=True)
            self.assertEqual(workbook.sheetnames, [
                "01_Перевозки", "02_Пломбы", "03_Страны",
                "04_Места (3)", "Количество перевозок",
            ])
            self.assertEqual(
                [cell.value for cell in workbook["01_Перевозки"][1]],
                [
                    "ID", "Дата_среза", "Этап", "Показатель", "Направление",
                    "Транспорт", "Пункт_пропуска_место", "Код_страны",
                    "Страна", "Количество_перевозок", "Источник",
                ],
            )
            self.assertEqual(workbook["01_Перевозки"].cell(2, 4).value, "Транзит (наложение НП)")
            self.assertEqual(workbook["01_Перевозки"].cell(2, 10).value, 3)
            self.assertEqual(workbook["02_Пломбы"].max_row, 29)
            self.assertEqual(workbook["04_Места (3)"].cell(2, 1).value, "тп «Аксу»")

    def test_report_and_power_bi_files_are_written_to_separate_paths(self) -> None:
        tables = [{"table": "Таблица6", "stage": "stage2-direct", "rows": [
            {"Страна начала перевозки": "Российская Федерация", "Количество перевозок": 1},
            {"Страна начала перевозки": "Итого", "Количество перевозок": 1},
        ]}]
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            report_path = root / "results" / "stage 2 report_20260907_1200.xlsx"
            power_bi_path = root / "power bi data" / "stage 2 report_20260907_1200.xlsx"
            self.assertEqual(_write_report(report_path, power_bi_path, tables, "1 - Старый метод"), 0)
            self.assertTrue(report_path.exists())
            self.assertTrue(power_bi_path.exists())

    def test_vehicle_number_registry_prioritizes_auto_and_recognizes_rail(self) -> None:
        registry = ROOT / "src" / "data" / "vehicle_plate_formats.json"
        self.assertEqual(
            detect_transport_by_vehicle_numbers("08KG445PK, 58010315", registry),
            "auto",
        )
        self.assertEqual(
            detect_transport_by_vehicle_numbers("CICU8735744 / 58010315", registry),
            "rail",
        )
        self.assertIsNone(detect_transport_by_vehicle_numbers("не указан", registry))

    def test_export_and_trade_rail_values_come_from_rail_source(self) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append([
            "Дата",
            "Владелец записи",
            "Уникальный номер перевозки",
            "Статус",
            "Тип",
            "company_id",
            "End waypoint",
        ])
        worksheet.append([
            "2026-09-01 10:00:00", "КЕДЕН KEDEN", "R-EXPORT", "finished",
            "Таможенная процедура экспорта", 2, "БРЕСТ-ВОСТОЧНЫЙ",
        ])
        worksheet.append([
            "2026-09-02 10:00:00", "ЭСФ ESF", "R-TRADE", "activated",
            "Перевозка товаров в рамках взаимной торговли", 2, 'МТО "Аламедин"',
        ])
        worksheet.append([
            "2026-09-02 10:00:00", "КЕДЕН KEDEN", "R-IGNORE", "planned",
            "Таможенная процедура экспорта", 2, "БРЕСТ-ВОСТОЧНЫЙ",
        ])

        export_values, trade_values, stats = build_rail_export_and_trade(
            {"ws": worksheet}, ROOT / "src" / "data"
        )

        self.assertEqual(export_values["BY"], 1)
        self.assertEqual(trade_values["KG"], 1)
        self.assertEqual(stats["rows_used"], 2)

    def test_console_choice_and_method_filenames(self) -> None:
        with patch("builtins.input", return_value="1"):
            self.assertEqual(_choose_method(), "6")
        with patch("builtins.input", side_effect=["2", "3"]):
            self.assertEqual(_choose_method(), "3")
        with patch("builtins.input", side_effect=["x", "2", "9", "4"]):
            self.assertEqual(_choose_method(), "4")
        with patch("builtins.input", side_effect=["2", "5"]):
            self.assertEqual(_choose_method(), "5")
        with patch("builtins.input", side_effect=["2", "0", "3"]):
            self.assertEqual(_choose_method(), "0")

        out_dir = Path("out")
        stamp = "20260904_1200"
        old_path = _make_output_filename(out_dir, "old method", stamp)
        new_path = _make_output_filename(out_dir, "new method", stamp)
        self.assertNotEqual(old_path, new_path)
        self.assertIn("old method", old_path.name)
        self.assertIn("new method", new_path.name)
        self.assertEqual(
            _make_full_period_output_filename(out_dir, stamp).name,
            "full period report_20260904_1200.xlsx",
        )

    def test_hybrid_uses_first_tables_from_big_and_remaining_from_old(self) -> None:
        big_tables = [
            {"table": "Таблица1", "source": "big"},
            {"table": "Таблица2", "source": "big"},
            {"table": "Таблица3", "source": "big"},
            {"table": "Таблица4", "source": "big"},
            {"table": "Таблица5", "source": "big"},
            {"table": "Таблица6", "source": "big"},
        ]
        old_tables = [
            {"table": "Таблица1", "source": "old"},
            {"table": "Таблица2", "source": "old"},
            {"table": "Таблица3", "source": "old"},
            {"table": "Таблица4", "source": "old"},
            {"table": "Таблица5", "source": "old"},
            {"table": "Таблица6", "source": "old"},
        ]

        hybrid = _merge_hybrid_tables(big_tables, old_tables)

        self.assertEqual(
            [table["source"] for table in hybrid],
            ["big", "big", "big", "old", "old", "old", "old"],
        )

    def test_declared_rail_transport_accepts_source_typo(self) -> None:
        self.assertEqual(_transport("Железнодорожный траспорт"), "rail")

    def test_alakol_is_strictly_automobile_post(self) -> None:
        registry = ROOT / "src" / "data" / "border_crossing_points.json"
        groups = load_t1_transport_groups(registry, ["Алаколь"])
        self.assertEqual(groups["Алаколь"], {"auto"})

    def test_external_points_follow_the_map_transport_legend(self) -> None:
        registry = ROOT / "src" / "data" / "border_crossing_points.json"
        groups = load_t1_transport_groups(
            registry,
            ["Бахты", "Капланбек", "Курык порт", "Морпорт актау"],
        )
        self.assertEqual(groups["Бахты"], {"auto"})
        self.assertEqual(groups["Капланбек"], {"auto"})
        self.assertEqual(groups["Курык порт"], {"rail"})
        self.assertEqual(groups["Морпорт актау"], {"rail"})

    def test_table1_rows_match_the_approved_template(self) -> None:
        config = load_config()
        groups = load_t1_transport_groups(
            ROOT / "src" / "data" / "border_crossing_points.json",
            config.tp_zhd_list_t1,
        )
        auto_points = [
            point for point in config.tp_zhd_list_t1
            if groups.get(point) == {"auto"}
        ]
        rail_points = [
            point for point in config.tp_zhd_list_t1
            if groups.get(point) == {"rail"}
        ]
        self.assertEqual(auto_points, [
            "Алаколь", "Казыгурт", "Нур жолы", "Калжат",
            "Бахты", "Тажен", "Атамекен", "Б. конысбаева",
            "Майкапчагай", "Темир-баба", "Капланбек",
        ])
        self.assertEqual(rail_points, [
            "Морпорт актау", "Курык порт", "Болашак", "Оазис",
            "Сарыагаш", "Алтынколь", "Темиржол",
        ])

    def test_rail_transit_is_grouped_by_start_waypoint_aliases(self) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append([
            "Дата",
            "Владелец записи",
            "Уникальный номер перевозки",
            "Статус",
            "company_id",
            "Start waypoint",
            "End waypoint",
        ])
        worksheet.append([
            "2026-09-01 10:00:00",
            "КЕДЕН KEDEN",
            "R-1",
            "finished",
            2,
            'Т/П "АЛТЫНКОЛЬ-ЖОЛ"',
            "БРЕСТ-ВОСТОЧНЫЙ",
        ])
        worksheet.append([
            "2026-09-02 10:00:00",
            "ЭСФ ESF",
            "R-2",
            "activated",
            2,
            "ТП «КУРЫК» ЖД и авто порт",
            'МТО "Аламедин"',
        ])
        worksheet.append([
            "2026-09-03 10:00:00",
            "КЕДЕН KEDEN",
            "R-3",
            "finished",
            2,
            'Т/П "ТЕМИРЖОЛ"',
            "БРЕСТ-ВОСТОЧНЫЙ",
        ])

        config = load_config()
        values, stats = build_rail_transit_by_point(
            {"ws": worksheet},
            ROOT / "src" / "data" / "border_crossing_points.json",
            config.tp_zhd_list_t1,
            ROOT / "src" / "data",
        )

        self.assertEqual(values[("Алтынколь", "BY")], 1)
        self.assertEqual(values[("Курык порт", "KG")], 1)
        self.assertEqual(values[("Темиржол", "BY")], 1)
        self.assertEqual(sum(values.values()), 3)
        self.assertEqual(stats["rows_used"], 3)

    def _make_source(self, path: Path) -> None:
        headers = list(REQUIRED_HEADERS)
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Лист1"
        worksheet.append(headers)

        def add(**values):
            defaults = {header: "NULL" for header in headers}
            defaults.update(
                {
                    "Владелец": "КЕДЕН KEDEN",
                    "Номер перевозки": len(worksheet["A"]) + 1,
                    "Компания": "Тест",
                    "БИН": "123",
                    "Номера документов": "NULL",
                    "Номера ТС": "TEST",
                    "Коды ТН ВЭД": "NULL",
                    "Кол. путевых точек": "2",
                    "Кол. водителей": "1",
                    "Кол. операторов": "2",
                    "Промежуточные точки": "",
                    "Промежуточные страны": "",
                }
            )
            defaults.update(values)
            worksheet.append([defaults[header] for header in headers])

        add(
            **{
                "Дата создания": "2026-07-01 10:00:00",
                "Номер перевозки": "PRE-STAGE2",
                "Дата обновления": "2026-09-01 10:00:00",
                "Статус": "Завершенная перевозка",
                "Тип перевозки": "Таможенная процедура таможенного транзита",
                "Тип транспорта": "Перевозка автомобильным транспортном",
                "Начальный пункт": 'Т/П "ТАЖЕН"',
                "Конечный пункт": "Россия",
                "Страна начала": "Республика Казахстан",
                "Страна завершения": "Российская Федерация",
                "Номера НП": "NP-1",
            }
        )
        add(
            **{
                "Дата создания": "2026-09-03 10:00:00",
                "Номер перевозки": "RAIL-1",
                "Дата обновления": "2026-09-04 10:00:00",
                "Статус": "Завершенная перевозка",
                "Тип перевозки": "Таможенная процедура таможенного транзита",
                "Тип транспорта": "Перевозка железнодорожным транспортном",
                "Начальный пункт": 'Т/П "ТЕМИРЖОЛ"',
                "Конечный пункт": "Брест",
                "Страна начала": "Республика Казахстан",
                "Страна завершения": "Республика Беларусь",
                "Номера НП": "NP-2, NP-3",
            }
        )
        add(
            **{
                "Дата создания": "2026-09-05 10:00:00",
                "Номер перевозки": "EXPORT-1",
                "Дата обновления": "2026-09-05 18:00:00",
                "Статус": "Завершенная перевозка",
                "Тип перевозки": "Таможенная процедура экспорта",
                "Тип транспорта": "Перевозка автомобильным транспортном",
                "Начальный пункт": "Астана",
                "Конечный пункт": "Россия",
                "Страна начала": "Республика Казахстан",
                "Страна завершения": "Российская Федерация",
                "Номера НП": "NULL",
            }
        )
        add(
            **{
                "Дата создания": "2026-09-06 10:00:00",
                "Номер перевозки": "TRADE-1",
                "Дата обновления": "2026-09-06 11:00:00",
                "Статус": "Активированная перевозка",
                "Тип перевозки": "Перевозка товаров в рамках взаимной торговли",
                "Тип транспорта": "Перевозка автомобильным транспортном",
                "Начальный пункт": "Астана",
                "Конечный пункт": "Бишкек",
                "Страна начала": "Республика Казахстан",
                "Страна завершения": "Кыргызская Республика",
                "Номера НП": "NP-4",
            }
        )
        add(
            **{
                "Дата создания": "2026-09-01 10:00:00",
                "Номер перевозки": "INCOMING-1",
                "Дата обновления": "2026-09-01 11:00:00",
                "Статус": "Активированная перевозка",
                "Тип перевозки": "Таможенная процедура экспорта",
                "Тип транспорта": "Перевозка автомобильным транспортном",
                "Начальный пункт": "Россия",
                "Конечный пункт": "Астана",
                "Страна начала": "Российская Федерация",
                "Страна завершения": "Республика Казахстан",
                "Номера НП": "NP-5",
            }
        )

        for number, country_from, country_to, transport, status in (
            ("TRACK-BY-KG", "Республика Беларусь", "Кыргызская Республика", "Автомобильный транспорт", "Активированная перевозка"),
            ("TRACK-RU-KG", "Российская Федерация", "Кыргызская Республика", "Железнодорожный траспорт", "Завершенная перевозка"),
            ("TRACK-KG-RU", "Кыргызская Республика", "Российская Федерация", "Автомобильный транспорт", "Завершенная перевозка"),
            ("TRACK-KG-BY", "Кыргызская Республика", "Республика Беларусь", "Железнодорожный траспорт", "Активированная перевозка"),
        ):
            add(**{
                "Дата создания": "2026-09-01 00:00:00",
                "Номер перевозки": number,
                "Дата обновления": "2026-09-02 10:00:00",
                "Статус": status,
                "Тип перевозки": "Таможенная процедура таможенного транзита",
                "Тип транспорта": transport,
                "Начальный пункт": country_from,
                "Конечный пункт": country_to,
                "Страна начала": country_from,
                "Страна завершения": country_to,
                "Номера НП": "NULL",
            })

        for number, created, status, trip_type, country_to in (
            ("TRACK-BEFORE", "2026-08-31 23:59:59", "Активированная перевозка", "Таможенная процедура таможенного транзита", "Кыргызская Республика"),
            ("TRACK-STATUS", "2026-09-01 00:00:00", "Ожидание навешивания", "Таможенная процедура таможенного транзита", "Кыргызская Республика"),
            ("TRACK-EXPORT", "2026-09-01 00:00:00", "Завершенная перевозка", "Таможенная процедура экспорта", "Кыргызская Республика"),
            ("TRACK-WRONG-PAIR", "2026-09-01 00:00:00", "Завершенная перевозка", "Таможенная процедура таможенного транзита", "Российская Федерация"),
        ):
            add(**{
                "Дата создания": created,
                "Номер перевозки": number,
                "Дата обновления": "2026-09-02 10:00:00",
                "Статус": status,
                "Тип перевозки": trip_type,
                "Тип транспорта": "Автомобильный транспорт",
                "Начальный пункт": "Республика Беларусь",
                "Конечный пункт": country_to,
                "Страна начала": "Республика Беларусь",
                "Страна завершения": country_to,
                "Номера НП": "NULL",
            })

        # Одна и та же перевозка может повторяться в выгрузке после обновления.
        duplicate = [cell.value for cell in worksheet[4]]
        duplicate[headers.index("Дата обновления")] = "2026-09-05 19:00:00"
        worksheet.append(duplicate)
        workbook.save(path)

    def test_accepts_html_xls_with_current_column_names(self) -> None:
        """The 10.09 large export is an HTML table with an .xls suffix."""
        with tempfile.TemporaryDirectory() as temp_dir:
            source_xlsx = Path(temp_dir) / "source.xlsx"
            source_html_xls = Path(temp_dir) / "file.xls"
            self._make_source(source_xlsx)

            workbook = load_workbook(source_xlsx, data_only=True)
            worksheet = workbook.active
            renamed = {
                "Кол. водителей": "Водителей",
                "Кол. путевых точек": "Путевых точек",
                "Кол. операторов": "Операторов",
                "Промежуточные точки": "Промежуточные пункты",
            }
            html_rows = []
            for row in worksheet.iter_rows(values_only=True):
                html_rows.append(
                    "<tr>" + "".join(
                        f"<td>{escape(str(renamed.get(value, value) if value is not None else ''))}</td>"
                        for value in row
                    ) + "</tr>"
                )
            workbook.close()
            source_html_xls.write_text(
                "<html><body><table>" + "".join(html_rows) + "</table></body></html>",
                encoding="utf-8",
            )

            config = load_config()
            combined = dict(config.normalization)
            combined["tp_zhd_list_t1"] = config.tp_zhd_list_t1
            combined["t1_transport_groups"] = load_t1_transport_groups(
                ROOT / "src" / "data" / "border_crossing_points.json",
                config.tp_zhd_list_t1,
            )
            tables = build_tables_from_big_export(
                source_html_xls,
                None,
                combined,
                report_date=dt.date(2026, 9, 6),
                split_by_transport=True,
                point_registry_path=ROOT / "src" / "data" / "border_crossing_points.json",
                vehicle_registry_path=ROOT / "src" / "data" / "vehicle_plate_formats.json",
            )

            self.assertTrue(is_big_export_file(source_html_xls))
            self.assertEqual(tables[0]["rows"][-1]["Итого"], 1)
            self.assertEqual(tables[1]["export_rows"][-1]["Количество перевозок"], 1)

    def test_builds_six_output_tables_from_dates_and_statuses(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "big.xlsx"
            output_path = Path(temp_dir) / "result.xlsx"
            self._make_source(source_path)

            config = load_config()
            combined = dict(config.normalization)
            combined["tp_zhd_list_t1"] = config.tp_zhd_list_t1
            tables = build_tables_from_big_export(
                source_path,
                None,
                combined,
                report_date=dt.date(2026, 9, 6),
            )

            self.assertTrue(is_big_export_file(source_path))
            self.assertEqual([table["table"] for table in tables], [
                "Таблица1",
                "Таблица2",
                "Таблица3",
                "Таблица4",
                "Таблица5",
                "Таблица6",
                "Таблица7",
            ])
            self.assertEqual(tables[0]["rows"][-1]["Итого"], 1)
            self.assertEqual(tables[0]["rail_registered_total"], 1)
            self.assertEqual(tables[1]["export_rows"][-1]["Количество перевозок"], 1)
            self.assertEqual(tables[1]["trade_rows"][-1]["Количество перевозок"], 1)
            self.assertEqual(tables[2]["rows"][0]["РБ - КР"], 1)
            self.assertEqual(tables[2]["rows"][0]["КР - РФ"], 1)
            self.assertEqual(tables[2]["rows"][1]["РФ - КР"], 1)
            self.assertEqual(tables[2]["rows"][1]["КР - РБ"], 1)
            self.assertEqual(tables[2]["rows"][-1]["Итого"], 4)
            # Первый блок Т4 считает завершенные исходящие перевозки нужных
            # операторов независимо от ошибочной отметки типа транспорта.
            self.assertEqual(tables[3]["auto_out_total"], 2)
            self.assertEqual(tables[4]["used_np_total"], 3)
            self.assertEqual(tables[5]["rows"][-1]["Количество перевозок"], 1)

            combined["t1_transport_groups"] = {"Темиржол": {"rail"}}
            split_tables = build_tables_from_big_export(
                source_path,
                None,
                combined,
                report_date=dt.date(2026, 9, 6),
                rail_transport_ids={"RAIL-1"},
            )
            self.assertTrue(split_tables[0]["transport_split"])
            self.assertEqual(split_tables[0]["rail_rows"][-1]["Итого"], 1)
            self.assertEqual(split_tables[0]["rail_registered_total"], 1)
            self.assertEqual(split_tables[1]["export_split_rows"][-1]["Авто"], 1)
            self.assertEqual(split_tables[1]["export_split_rows"][-1]["ЖД"], 0)
            self.assertEqual(split_tables[3]["rail_out_total"], 1)

            # Для транзита классификация самого поста главнее признака записи
            # в ЖД-реестре. Это не дает пропасть перевозкам авто-постов.
            combined["t1_transport_groups"] = {"Темиржол": {"auto"}}
            auto_post_tables = build_tables_from_big_export(
                source_path,
                None,
                combined,
                report_date=dt.date(2026, 9, 6),
                rail_transport_ids={"RAIL-1"},
            )
            self.assertEqual(auto_post_tables[0]["auto_rows"][-1]["Итого"], 1)
            self.assertEqual(auto_post_tables[0]["rail_rows"][-1]["Итого"], 0)

            write_output(output_path, tables)
            output = load_workbook(output_path, data_only=False)
            self.assertEqual(output.sheetnames, [
                "Таблица1",
                "Таблица2",
                "Таблица3",
                "Таблица4",
                "Таблица5",
                "Таблица6",
                "Таблица7",
            ])
            tracking_sheet = output["Таблица3"]
            self.assertEqual(tracking_sheet["A1"].value, "Услуги прослеживания")
            self.assertEqual(tracking_sheet["A2"].value, "Направления")
            self.assertEqual(
                [tracking_sheet.cell(3, column).value for column in range(1, 7)],
                ["Тип транспорта", "РБ - КР", "РФ - КР", "КР - РФ", "КР - РБ", "Итого"],
            )
            self.assertEqual(tracking_sheet["A4"].value, "Авто")
            self.assertEqual(tracking_sheet["A5"].value, "ЖД")
            self.assertEqual(tracking_sheet["A6"].value, "Итого")

    def test_seal_application_table_uses_post_aliases_and_unique_np_numbers(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "big.xlsx"
            self._make_source(source_path)
            config = load_config()
            combined = dict(config.normalization)
            combined["tp_zhd_list_t1"] = config.tp_zhd_list_t1
            combined["t1_transport_groups"] = load_t1_transport_groups(
                ROOT / "src" / "data" / "border_crossing_points.json",
                config.tp_zhd_list_t1,
            )
            tables = build_tables_from_big_export(
                source_path,
                None,
                combined,
                report_date=dt.date(2026, 9, 6),
                split_by_transport=True,
                point_registry_path=ROOT / "src" / "data" / "border_crossing_points.json",
                vehicle_registry_path=ROOT / "src" / "data" / "vehicle_plate_formats.json",
            )
            table7 = tables[6]
            auto = {row["ТП/ЖД"]: row for row in table7["auto_rows"]}
            rail = {row["ТП/ЖД"]: row for row in table7["rail_rows"]}
            self.assertEqual(auto["Тажен"]["РФ"], 0)
            self.assertEqual(rail["Темиржол"]["РБ"], 2)
            self.assertEqual(table7["rail_total"], 2)
            self.assertEqual(table7["total"], 2)

    def test_full_period_override_starts_exactly_on_11_february(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "big.xlsx"
            self._make_source(source_path)

            workbook = load_workbook(source_path)
            worksheet = workbook.active
            headers = [cell.value for cell in worksheet[1]]
            created_col = headers.index("Дата создания") + 1
            number_col = headers.index("Номер перевозки") + 1

            # PRE-STAGE2 becomes the exact inclusive boundary record.
            worksheet.cell(2, created_col).value = "2026-02-11 00:00:00"
            before = [cell.value for cell in worksheet[2]]
            before[created_col - 1] = "2026-02-10 23:59:59"
            before[number_col - 1] = "BEFORE-FULL-PERIOD"
            worksheet.append(before)
            workbook.save(source_path)
            workbook.close()

            config = load_config()
            combined = dict(config.normalization)
            combined["tp_zhd_list_t1"] = config.tp_zhd_list_t1
            combined["t1_transport_groups"] = load_t1_transport_groups(
                ROOT / "src" / "data" / "border_crossing_points.json",
                config.tp_zhd_list_t1,
            )

            stage2_tables = build_tables_from_big_export(
                source_path,
                None,
                combined,
                report_date=dt.date(2026, 9, 6),
            )
            full_period_tables = build_tables_from_big_export(
                source_path,
                None,
                combined,
                report_date=dt.date(2026, 9, 6),
                period_start=dt.datetime(2026, 2, 11, 0, 0, 0),
                split_by_transport=True,
                point_registry_path=ROOT / "src" / "data" / "border_crossing_points.json",
                vehicle_registry_path=ROOT / "src" / "data" / "vehicle_plate_formats.json",
            )

            self.assertEqual(stage2_tables[0]["rows"][-1]["Итого"], 1)
            self.assertEqual(full_period_tables[0]["rows"][-1]["Итого"], 2)
            self.assertEqual(full_period_tables[0]["stage"], "full-period")
            self.assertEqual(full_period_tables[6]["auto_total"], 1)
            self.assertEqual(full_period_tables[6]["rail_total"], 2)

    def test_split_mode_uses_only_big_export_for_auto_and_rail(self) -> None:
        """Режим 3 не должен зависеть от отдельной ЖД-выгрузки."""
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "big.xlsx"
            self._make_source(source_path)

            config = load_config()
            split_config = dict(config.normalization)
            split_config["tp_zhd_list_t1"] = config.tp_zhd_list_t1
            split_config["t1_transport_groups"] = load_t1_transport_groups(
                ROOT / "src" / "data" / "border_crossing_points.json",
                config.tp_zhd_list_t1,
            )
            tables = build_tables_from_big_export(
                source_path,
                None,
                split_config,
                report_date=dt.date(2026, 9, 6),
                split_by_transport=True,
                point_registry_path=ROOT / "src" / "data" / "border_crossing_points.json",
                vehicle_registry_path=ROOT / "src" / "data" / "vehicle_plate_formats.json",
            )

            self.assertTrue(tables[0]["transport_split"])
            self.assertEqual(tables[0]["rail_rows"][-1]["Итого"], 1)
            self.assertEqual(tables[0]["rail_registered_total"], 1)
            self.assertEqual(tables[1]["export_split_rows"][-1]["Авто"], 1)
            self.assertEqual(tables[1]["export_split_rows"][-1]["ЖД"], 0)
            self.assertEqual(tables[1]["trade_split_rows"][-1]["Авто"], 1)
            self.assertEqual(tables[3]["rail_out_total"], 1)

    def test_table4_uses_completed_status_and_not_owner_filter(self) -> None:
        """Т4 включает завершенные независимо от оператора, но не деактивированные."""
        with tempfile.TemporaryDirectory() as temp_dir:
            source_path = Path(temp_dir) / "big.xlsx"
            self._make_source(source_path)
            config = load_config()
            split_config = dict(config.normalization)
            split_config["tp_zhd_list_t1"] = config.tp_zhd_list_t1
            split_config["t1_transport_groups"] = load_t1_transport_groups(
                ROOT / "src" / "data" / "border_crossing_points.json",
                config.tp_zhd_list_t1,
            )
            build_kwargs = {
                "report_date": dt.date(2026, 9, 6),
                "split_by_transport": True,
                "point_registry_path": ROOT / "src" / "data" / "border_crossing_points.json",
                "vehicle_registry_path": ROOT / "src" / "data" / "vehicle_plate_formats.json",
            }
            before = build_tables_from_big_export(
                source_path, None, split_config, **build_kwargs
            )

            workbook = load_workbook(source_path)
            worksheet = workbook.active
            headers = [cell.value for cell in worksheet[1]]
            source_row = [cell.value for cell in worksheet[4]]
            values = dict(zip(headers, source_row))
            values.update({
                "Номер перевозки": "OUT-ANY-OWNER",
                "Владелец": "Сторонний оператор",
                "Дата создания": "2026-09-06 10:00:00",
                "Дата обновления": "2026-09-06 12:00:00",
                "Статус": "Завершенная перевозка",
                "Страна начала": "Республика Казахстан",
                "Страна завершения": "Российская Федерация",
                "Тип транспорта": "Автомобильный транспорт",
            })
            worksheet.append([values[header] for header in headers])
            values.update({
                "Номер перевозки": "IN-DEACTIVATED",
                "Статус": "Деактивированная перевозка",
                "Страна начала": "Российская Федерация",
                "Страна завершения": "Республика Казахстан",
            })
            worksheet.append([values[header] for header in headers])
            workbook.save(source_path)
            workbook.close()

            after = build_tables_from_big_export(
                source_path, None, split_config, **build_kwargs
            )
            self.assertEqual(
                after[3]["auto_out_total"], before[3]["auto_out_total"] + 1
            )
            self.assertEqual(
                after[3]["auto_in_total"], before[3]["auto_in_total"]
            )


if __name__ == "__main__":
    unittest.main()
