from __future__ import annotations

import datetime as dt
import re
import shutil
import time
from copy import deepcopy
from pathlib import Path

from big_export import (
    STAGE2_START,
    build_tracking_table,
    build_tables_from_big_export,
    extract_stage2_rail_ids,
    find_big_export_file,
    is_big_export_file,
    load_t1_transport_groups,
)
from daily_raw04 import build_daily_raw04_tables, find_raw04_export, is_raw04_export
from builders import (
    build_rail_registered_count,
    build_table1,
    build_table2,
    build_table3,
    build_table4,
    build_table5,
)

from config import TP_ZHD_LIST_T1_OLD, load_config

from io_excel import (
    read_rail_report,
    read_report_1,
    read_report_2,
    write_output,
)

from normalize import normalize_report
from monthly_np import (
    calculate_monthly_rows,
    load_previous_month_snapshot,
    make_snapshot,
    parse_report_date,
    save_snapshot,
)
from power_bi import write_power_bi_output
from stage2 import STAGE2_BASELINE_FILE, apply_stage2_baseline, load_stage2_baseline
from utils import app_path, monthly_np_storage_path, resource_path


REPORT_PATTERN = re.compile(r"^report_.*\.xlsx$", re.IGNORECASE)
APP_VERSION = "2.7.0"
FULL_PERIOD_START = dt.datetime(2026, 2, 11, 0, 0, 0)

TABLE4_PATTERN = re.compile(
    r"^готовность.*апп.*территория.*\.xlsx$",
    re.IGNORECASE,
)

RAIL_PATTERN = re.compile(
    r"^(жд\s*выгрузка.*\.xlsx|.*\.xls)$",
    re.IGNORECASE,
)
#вот тут добавить разные форматы входа файла, xlm формат.

def _make_output_filename(
    out_dir: Path,
    method: str = "",
    timestamp: str | None = None,
) -> Path:
    ts = timestamp or dt.datetime.now().strftime("%Y%m%d_%H%M")
    suffix = f" {method}" if method else ""
    return out_dir / f"stage 2 report{suffix}_{ts}.xlsx"


def _make_full_period_output_filename(
    out_dir: Path,
    timestamp: str | None = None,
) -> Path:
    ts = timestamp or dt.datetime.now().strftime("%Y%m%d_%H%M")
    return out_dir / f"full period report_{ts}.xlsx"


def _make_daily_raw04_output_filename(
    out_dir: Path,
    timestamp: str | None = None,
) -> Path:
    ts = timestamp or dt.datetime.now().strftime("%Y%m%d_%H%M")
    return out_dir / f"daily NP report RAW04_{ts}.xlsx"


def _pause():
    input("\nНажмите Enter для выхода...")


def _find_input_file(in_dir: Path) -> Path | None:

    files = []

    for f in in_dir.glob("*.xlsx"):

        name = f.name

        if name.startswith("~$"):
            continue

        if REPORT_PATTERN.match(name) and not is_big_export_file(f):
            files.append(f)

    if len(files) == 0:
        print("Ошибка: файл источника не найден.")
        print("В папке in должен находиться один файл report_*.xlsx")
        return None

    if len(files) > 1:
        print("Ошибка: найдено несколько файлов источника.")
        print("В папке in должен находиться только один файл report_*.xlsx\n")

        for f in files:
            print(" -", f.name)

        return None

    return files[0]


def _find_optional_input_file(in_dir: Path) -> Path | None:
    """Find the usual report without turning its absence into an error."""
    files = [
        file
        for file in in_dir.glob("*.xlsx")
        if not file.name.startswith("~$")
        and REPORT_PATTERN.match(file.name)
        and not is_big_export_file(file)
    ]
    if len(files) == 1:
        print(f"[raw04] дополнительный report найден: {files[0].name}")
        return files[0]
    if not files:
        print(
            "[raw04] ВНИМАНИЕ: report_*.xlsx отсутствует. "
            "Таблица8 с остатками и месячным расходом НП будет помечена "
            "как недоступная."
        )
        return None
    print("[raw04] ВНИМАНИЕ: найдено несколько report_*.xlsx; Таблица8 недоступна.")
    for file in files:
        print(" -", file.name)
    return None


def _find_table4_file(in_dir: Path) -> Path | None:

    found = []

    for f in in_dir.glob("*.xlsx"):

        name = f.name

        if name.startswith("~$"):
            continue

        if TABLE4_PATTERN.match(name):
            found.append(f)

    if len(found) == 0:
        print("[main] table4 source file NOT FOUND")
        return None

    if len(found) > 1:
        print("[main] table4 source: found multiple files")
        for f in found:
            print(" -", f.name)

        print("[main] Таблица5 будет пропущена")
        return None

    print(f"[main] table4 source FOUND: {found[0].name}")

    return found[0]


def _find_rail_file(in_dir: Path) -> Path | None:

    found = []

    for f in in_dir.iterdir():

        if not f.is_file():
            continue

        name = f.name

        if name.startswith("~$"):
            continue

        # RAW04 files are often delivered as a large HTML-XLS named simply
        # ``file.xls``.  The old broad rail pattern also matched that name,
        # so the program parsed the 100+ MB RAW04 export once as rail data and
        # then once more as RAW04.  Apart from being very slow, this was not a
        # valid rail source.  A real rail file remains eligible below.
        if is_raw04_export(f):
            print(f"[main] rail source skipped (RAW04 export): {name}")
            continue

        if RAIL_PATTERN.match(name):
            found.append(f)

    if len(found) == 0:
        print("[main] ЖД Выгрузка.xlsx NOT FOUND")
        print("[main] Таблица4 будет создана без ЖД-уточнения")
        return None

    if len(found) > 1:
        print("[main] rail source: found multiple files")
        for f in found:
            print(" -", f.name)

        print("[main] Таблица4 будет создана без ЖД-уточнения")
        return None

    print(f"[main] rail source FOUND: {found[0].name}")

    return found[0]


def _confirm_continue_without(missing_names: list[str]) -> bool:

    if not missing_names:
        return True

    print("\n[main] Не найдены дополнительные файлы:")
    for name in missing_names:
        print(" -", name)

    while True:
        answer = input("Продолжить без этих файлов? (y/n): ").strip().lower()
        if answer in {"y", "yes", "д", "да"}:
            return True
        if answer in {"n", "no", "н", "нет"}:
            return False

        print("Введите y или n.")


def _move_processed(src: Path, in_dir: Path):

    processed_dir = in_dir / "processed"
    processed_dir.mkdir(exist_ok=True)

    ts = dt.datetime.now().strftime("%Y%m%d_%H%M")

    base_name = src.stem + f"__processed_{ts}"
    dst = processed_dir / f"{base_name}{src.suffix}"

    counter = 1
    while dst.exists():
        dst = processed_dir / f"{base_name}(+{counter}){src.suffix}"
        counter += 1

    shutil.move(str(src), str(dst))


def _choose_other_method() -> str | None:
    print("\nОстальные режимы:")
    print("0 - Назад")
    print("1 - Старый метод, три привычных файла")
    print("2 - Новый метод, большая выгрузка и файл готовности")
    print(
        f"3 - Режим v{APP_VERSION}: с АПП КГД и большой выгрузкой"
    )
    print(
        "4 - Смешанный отчет: авто из большой выгрузки, все ЖД-данные "
        "из ЖД выгрузки, остальные таблицы по старому методу"
    )
    print(
        "5 - Единый период с 11.02.2026: расчет как в Основном режиме, "
        "без разделения на этапы"
    )
    while True:
        answer = input("Введите 0, 1, 2, 3, 4 или 5: ").strip()
        if answer == "0":
            return None
        if answer in {"1", "2", "3", "4", "5"}:
            return answer
        print("Нужно ввести 0, 1, 2, 3, 4 или 5.")


def _choose_method() -> str:
    while True:
        print("\nГлавное меню:")
        print("1 - Основной режим: Ежедневная справка RAW04")
        print("2 - Остальные режимы")
        print("3 - Выход")

        answer = input("Введите 1, 2 или 3: ").strip()
        if answer == "1":
            return "6"
        if answer == "2":
            method = _choose_other_method()
            if method is not None:
                return method
            continue
        if answer == "3":
            return "0"
        print("Нужно ввести 1, 2 или 3.")


def _read_table4_source(table4_file: Path | None):
    if table4_file is None:
        return None
    try:
        source = read_report_2(table4_file)
        print("[main] table4 report read: OK")
        return source
    except Exception as exc:
        print("[main] table4 report read: FAIL")
        print(exc)
        return None


def _build_old_tables(
    src_file: Path,
    table4_source,
    rail_file: Path | None,
):
    cfg = load_config()
    legacy_t1_points = TP_ZHD_LIST_T1_OLD

    try:
        r1 = read_report_1(src_file, sheet_name="Отчёт")
        print("[check] report read: OK")
    except Exception as exc:
        print("Ошибка: неверный файл источника.")
        print("Проверьте структуру файла report_*.xlsx")
        print(exc)
        return None, None, 2

    rail_source = None
    if rail_file is not None:
        try:
            rail_source = read_rail_report(rail_file)
            print("[main] rail report read: OK")
        except Exception as exc:
            print("[main] rail report read: FAIL")
            print(exc)
            print("[main] Таблица4 будет создана без ЖД-уточнения")

    r1_norm = normalize_report(
        r1,
        cfg.normalization,
        allowed_pp=legacy_t1_points,
    )
    print("[main] normalize step: OK")

    print("[raw04] Старая Таблица1: расчет...")
    started = time.perf_counter()
    t1 = build_table1(r1_norm, cfg.normalization, legacy_t1_points)
    t1["legacy_layout"] = True
    t1["rail_registered_total"] = build_rail_registered_count(rail_source)
    print(f"[raw04] Старая Таблица1 готова ({time.perf_counter() - started:.1f} сек.)")

    print("[raw04] Старая Таблица2: расчет...")
    started = time.perf_counter()
    t2 = build_table2(r1_norm, cfg.normalization)
    print(f"[raw04] Старая Таблица2 готова ({time.perf_counter() - started:.1f} сек.)")

    print("[raw04] Старая Таблица4: расчет...")
    started = time.perf_counter()
    t3 = build_table3(
        r1_norm,
        cfg.normalization,
        rail_raw=rail_source,
        data_dir=resource_path("data"),
    )
    print(f"[raw04] Старая Таблица4 готова ({time.perf_counter() - started:.1f} сек.)")

    print("[raw04] Старая Таблица5: расчет...")
    started = time.perf_counter()
    t5 = build_table5(r1_norm, cfg.normalization)
    print(f"[raw04] Старая Таблица5 готова ({time.perf_counter() - started:.1f} сек.)")
    # Старый режим не получает большую выгрузку, поэтому новая Таблица3
    # присутствует в отчёте как честный пустой шаблон. Прежние Т3-Т5
    # сдвигаются на Т4-Т6 без изменения своей расчётной логики.
    tables = [t1, t2, build_tracking_table(), t3]
    tables.append({
        "table": "Таблица7",
        "stage": "legacy-empty",
        "auto_rows": [],
        "rail_rows": [],
        "total": 0,
    })
    monthly_np_snapshot = None

    try:
        print("[raw04] Старая Таблица8: расчет расхода пломб...")
        started = time.perf_counter()
        t4 = build_table4(r1_norm, table4_source, cfg.normalization)
        print(f"[raw04] Старая Таблица8 готова ({time.perf_counter() - started:.1f} сек.)")
        report_date = parse_report_date(src_file)
        monthly_np_snapshot = make_snapshot(
            t4.get("used_np_rows", []) or [],
            report_date,
            src_file.name,
        )
        monthly_baseline, monthly_baseline_exact = load_previous_month_snapshot(
            report_date,
            monthly_np_storage_path(),
            resource_path("data"),
        )
        t4["monthly_used_np_rows"] = calculate_monthly_rows(
            t4.get("used_np_rows", []) or [],
            monthly_baseline,
        )
        t4["monthly_report_date"] = report_date.isoformat()
        t4["monthly_elapsed_days"] = report_date.day
        t4["monthly_baseline_date"] = (
            str(monthly_baseline.get("snapshot_date"))
            if monthly_baseline is not None
            else ""
        )
        t4["monthly_baseline_exact"] = monthly_baseline_exact

        if monthly_baseline is None:
            print(
                "[main] ВНИМАНИЕ: база расхода пломб за предыдущий месяц "
                "не найдена; месячный расход не рассчитан"
            )
        elif not monthly_baseline_exact:
            print(
                "[main] ВНИМАНИЕ: используется последний доступный снимок "
                f"предыдущего месяца от {monthly_baseline.get('snapshot_date')}"
            )
        else:
            print(
                "[main] monthly NP baseline applied: "
                f"date={monthly_baseline.get('snapshot_date')} "
                f"source={monthly_baseline.get('_baseline_storage', 'unknown')}"
            )

        tables.append(t4)
        print("[main] Таблица5 created")
    except Exception as exc:
        print("[main] Таблица5 build FAIL")
        print(exc)

    tables.append(t5)

    try:
        baseline_path = resource_path(f"data/{STAGE2_BASELINE_FILE}")
        baseline = load_stage2_baseline(baseline_path)
        tables = apply_stage2_baseline(tables, baseline)
        print(
            "[main] stage 2 baseline applied: "
            f"cutoff={baseline.get('cutoff')}"
        )
    except Exception as exc:
        print("Ошибка: не удалось применить базу второго этапа.")
        print(exc)
        return None, None, 4

    return tables, monthly_np_snapshot, 0


def _build_new_tables(big_export_file: Path, table4_source):
    cfg = load_config()
    combined_config = dict(cfg.normalization)
    combined_config["tp_zhd_list_t1"] = list(cfg.tp_zhd_list_t1)
    try:
        tables = build_tables_from_big_export(
            big_export_file,
            table4_source,
            combined_config,
            report_date=dt.date.today(),
        )
    except Exception as exc:
        print("Ошибка: не удалось обработать большую выгрузку.")
        print(exc)
        return None, 5
    return tables, 0


def _build_daily_raw04_tables(big_export_file: Path):
    """Separate daily-certificate mode. Legacy table logic is not involved."""
    try:
        tables = build_daily_raw04_tables(
            big_export_file,
            resource_path("data/border_crossing_points.json"),
        )
    except Exception as exc:
        print("Ошибка: не удалось сформировать ежедневную справку RAW04.")
        print(exc)
        return None, 8
    return tables, 0


def _build_raw04_primary_without_report(big_export_file: Path):
    """Build the big-export tables when the optional stock report is absent."""
    cfg = load_config()
    split_config = dict(cfg.normalization)
    split_config["tp_zhd_list_t1"] = list(cfg.tp_zhd_list_t1)
    point_registry_path = resource_path("data/border_crossing_points.json")
    split_config["t1_transport_groups"] = load_t1_transport_groups(
        point_registry_path,
        cfg.tp_zhd_list_t1,
    )
    try:
        tables = build_tables_from_big_export(
            big_export_file,
            None,
            split_config,
            report_date=dt.date.today(),
            period_start=FULL_PERIOD_START,
            split_by_transport=True,
            point_registry_path=point_registry_path,
            vehicle_registry_path=resource_path("data/vehicle_plate_formats.json"),
            include_big_export_table4=False,
        )
    except Exception as exc:
        print("Ошибка: не удалось рассчитать RAW04 без дополнительного report.")
        print(exc)
        return None, 6

    # Таблица5 базового набора становится Таблицей8 в ежедневной справке.
    # Не подставляем нулевые или пилотные данные, если исходного report нет.
    tables[4].update({
        "stock_unavailable": True,
        "stock_unavailable_message": (
            "Нет файла report_*.xlsx: остатки и месячный расход НП не рассчитаны."
        ),
    })
    return tables, 0


def _merge_primary_and_raw04_tables(old_tables, raw04_tables):
    """Use the current primary mode for existing tables and RAW04 for new ones.

    The main mode owns tables 1, 2, 4, 5 (stock), and 6 (inbound).
    The certificate uses the latter two as tables 8 and 5, respectively.
    """
    old_by_name = {table.get("table"): table for table in old_tables}
    raw_by_name = {table.get("table"): table for table in raw04_tables}
    required_old = {"Таблица1", "Таблица2", "Таблица4", "Таблица5", "Таблица6"}
    required_raw = {"Таблица3", "Таблица6", "Таблица7", "Диагностика"}
    missing = sorted(
        [name for name in required_old if name not in old_by_name]
        + [name for name in required_raw if name not in raw_by_name]
    )
    if missing:
        raise ValueError("Не удалось собрать ежедневную справку: " + ", ".join(missing))

    # The main-mode inbound table becomes certificate Table5.
    legacy_inbound = deepcopy(old_by_name["Таблица6"])
    legacy_inbound["table"] = "Таблица5"
    legacy_inbound["daily_raw04"] = True

    # The main-mode stock and monthly usage table stays as Table8.
    legacy_stock = deepcopy(old_by_name["Таблица5"])
    legacy_stock["table"] = "Таблица8"
    legacy_stock["legacy_stock_layout"] = True

    base_tables = [
        old_by_name["Таблица1"],
        old_by_name["Таблица2"],
        raw_by_name["Таблица3"],
        old_by_name["Таблица4"],
        legacy_inbound,
        raw_by_name["Таблица6"],
        raw_by_name["Таблица7"],
        legacy_stock,
        raw_by_name["Диагностика"],
    ]
    tables = [_build_daily_raw04_header_table(base_tables), *base_tables]
    for table in tables:
        table["raw04_mode"] = True
    return tables


def _build_daily_raw04_header_table(tables):
    """Build the numeric certificate header from its output tables.

    The values stay live: no amount from the certificate prose is stored here.
    """
    by_name = {table.get("table"): table for table in tables}
    table1 = by_name.get("Таблица1", {})
    table2 = by_name.get("Таблица2", {})
    table3 = by_name.get("Таблица3", {})
    table4 = by_name.get("Таблица4", {})

    def row_total(rows, field="Итого"):
        return int((rows or [{}])[-1].get(field, 0) or 0)

    def t1_by_country(country):
        return sum(
            int((rows or [{}])[-1].get(country, 0) or 0)
            for rows in (table1.get("auto_rows"), table1.get("rail_rows"))
        )

    def t2_by_country(rows, country):
        for row in rows or []:
            if row.get("Страна") == country:
                return int(row.get("Итого", 0) or 0)
        return 0

    export_rows = table2.get("export_split_rows", []) or []
    trade_rows = table2.get("trade_split_rows", []) or []
    transit_total = sum(t1_by_country(code) for code in ("РФ", "РБ", "КР", "РА"))
    export_total = row_total(export_rows)
    trade_total = row_total(trade_rows)
    completed_in_rk = row_total(table4.get("auto_in_rows"), "Количество") + row_total(
        table4.get("rail_in_rows"), "Количество"
    )

    rows = [
        ("Национальным оператором оформлено перевозок", "Таблица 1 + Таблица 2", transit_total + export_total + trade_total),
        ("Осуществлен мониторинг перевозок", "Таблица 3: Итого", row_total(table3.get("rows"))),
        ("На территории РК завершено перевозок", "Таблица 4: в РК, авто + ЖД", completed_in_rk),
        ("Таможенный транзит", "Таблица 1: Итого (Сумма)", transit_total),
        ("Транзит в Республику Беларусь", "Таблица 1: РБ", t1_by_country("РБ")),
        ("Транзит в Российскую Федерацию", "Таблица 1: РФ", t1_by_country("РФ")),
        ("Транзит в Кыргызскую Республику", "Таблица 1: КР", t1_by_country("КР")),
        ("Экспорт", "Таблица 2: Итого (сумма)", export_total),
        ("Экспорт в Республику Беларусь", "Таблица 2: экспорт, РБ", t2_by_country(export_rows, "Республика Беларусь")),
        ("Экспорт в Российскую Федерацию", "Таблица 2: экспорт, РФ", t2_by_country(export_rows, "Российская Федерация")),
        ("Экспорт в Кыргызскую Республику", "Таблица 2: экспорт, КР", t2_by_country(export_rows, "Кыргызская Республика")),
        ("Взаимная торговля", "Таблица 2: Итого (сумма)", trade_total),
        ("Взаимная торговля в Кыргызскую Республику", "Таблица 2: взаимка, КР", t2_by_country(trade_rows, "Кыргызская Республика")),
        ("Взаимная торговля в Российскую Федерацию", "Таблица 2: взаимка, РФ", t2_by_country(trade_rows, "Российская Федерация")),
        ("Взаимная торговля в Республику Беларусь", "Таблица 2: взаимка, РБ", t2_by_country(trade_rows, "Республика Беларусь")),
    ]
    return {
        "table": "Шапка",
        "daily_raw04_header": True,
        "rows": [
            {"№": number, "Название": title, "Данные": source, "Число": value}
            for number, (title, source, value) in enumerate(rows, start=1)
        ],
    }


def _build_combined_tables(
    big_export_file: Path,
    table4_source,
    rail_file: Path,
):
    cfg = load_config()
    combined_config = dict(cfg.normalization)
    combined_config["tp_zhd_list_t1"] = list(cfg.tp_zhd_list_t1)
    try:
        rail_source = read_rail_report(rail_file)
        rail_ids = extract_stage2_rail_ids(rail_source)
        point_registry_path = resource_path("data/border_crossing_points.json")
        combined_config["t1_transport_groups"] = load_t1_transport_groups(
            point_registry_path,
            cfg.tp_zhd_list_t1,
        )
        tables = build_tables_from_big_export(
            big_export_file,
            table4_source,
            combined_config,
            report_date=dt.date.today(),
            rail_transport_ids=rail_ids,
            rail_raw=rail_source,
            point_registry_path=point_registry_path,
            data_dir=resource_path("data"),
            vehicle_registry_path=resource_path("data/vehicle_plate_formats.json"),
        )
        print(
            "[combined method] ЖД-перевозок второго этапа: ",
            len(rail_ids),
            sep="",
        )
        print(
            "[combined method] ЖД-транзит по начальным пунктам: ",
            tables[0].get("rail_rows", [{}])[-1].get("Итого", 0),
            sep="",
        )
    except Exception as exc:
        print("Ошибка: не удалось выполнить комбинированный расчет.")
        print(exc)
        return None, 6
    return tables, 0


def _build_split_big_export_tables(
    big_export_file: Path,
    table4_source,
    report_file: Path,
    period_start: dt.datetime | None = None,
    subtract_stage2_baseline: bool = True,
):
    """Режимы 3/5: Т1-Т4 и Т6 из большой выгрузки.

    Т5 собирается исключительно из обычного report и файла готовности
    АПП/ЖД. Отдельная ЖД-выгрузка здесь намеренно не читается: авто/ЖД
    определяются по типу транспорта, номерам ТС и реестру пунктов. Режим 5
    передает начало 11.02.2026 и не вычитает базу на 31.08.2026.
    """
    cfg = load_config()
    split_config = dict(cfg.normalization)
    split_config["tp_zhd_list_t1"] = list(cfg.tp_zhd_list_t1)
    point_registry_path = resource_path("data/border_crossing_points.json")
    split_config["t1_transport_groups"] = load_t1_transport_groups(
        point_registry_path,
        cfg.tp_zhd_list_t1,
    )
    try:
        report_raw = read_report_1(report_file, sheet_name="Отчёт")
        report_normalized = normalize_report(
            report_raw,
            cfg.normalization,
            allowed_pp=cfg.tp_zhd_list_t1,
        )
        table4 = build_table4(report_normalized, table4_source, cfg.normalization)
        # Правая часть Т5 — отдельный месячный показатель. Его считаем из
        # исходных накопительных значений: текущий report минус снимок на
        # последний день предыдущего календарного месяца. Это обязательно
        # делаем до вычитания базы второго этапа (31.08.2026).
        report_date = parse_report_date(report_file)
        monthly_np_snapshot = make_snapshot(
            table4.get("used_np_rows", []) or [],
            report_date,
            report_file.name,
        )
        monthly_baseline, monthly_baseline_exact = load_previous_month_snapshot(
            report_date,
            monthly_np_storage_path(),
            resource_path("data"),
        )
        table4["monthly_used_np_rows"] = calculate_monthly_rows(
            table4.get("used_np_rows", []) or [],
            monthly_baseline,
        )
        table4["monthly_report_date"] = report_date.isoformat()
        table4["monthly_elapsed_days"] = report_date.day
        table4["monthly_baseline_date"] = (
            str(monthly_baseline.get("snapshot_date"))
            if monthly_baseline is not None
            else ""
        )
        table4["monthly_baseline_exact"] = monthly_baseline_exact
        if monthly_baseline is None:
            print(
                "[split big export] ВНИМАНИЕ: база расхода пломб за "
                "предыдущий месяц не найдена; месячный расход не рассчитан"
            )
        elif not monthly_baseline_exact:
            print(
                "[split big export] ВНИМАНИЕ: используется последний "
                "доступный снимок предыдущего месяца от "
                f"{monthly_baseline.get('snapshot_date')}"
            )
        else:
            table4["stage"] = "full-period"
            print(
                "[split big export] monthly NP baseline applied: "
                f"date={monthly_baseline.get('snapshot_date')} "
                f"source={monthly_baseline.get('_baseline_storage', 'unknown')}"
            )
        if subtract_stage2_baseline:
            # Значения использованных НП в report накопительные. Для второго
            # этапа вычитаем фиксированный снимок на 31.08.2026; остатки НП в
            # левой части Т5 остаются текущим снимком и не вычитаются.
            baseline = load_stage2_baseline(
                resource_path(f"data/{STAGE2_BASELINE_FILE}")
            )
            table4 = apply_stage2_baseline([table4], baseline)[0]
        else:
            print(
                "[split big export] единый период: накопительные значения "
                "report используются без вычитания базы 31.08.2026"
            )
        tables = build_tables_from_big_export(
            big_export_file,
            None,
            split_config,
            report_date=dt.date.today(),
            period_start=period_start or STAGE2_START,
            split_by_transport=True,
            point_registry_path=point_registry_path,
            vehicle_registry_path=resource_path("data/vehicle_plate_formats.json"),
            include_big_export_table4=False,
        )
        # Не допускаем смешивания источников: Т5 режима 3 — только report и
        # готовность АПП/ЖД; Т1-Т4 и Т6 — только большая выгрузка.
        tables[4] = table4
    except Exception as exc:
        print("Ошибка: не удалось выполнить расчет с разделением авто/ЖД.")
        print(exc)
        return None, None, 6
    print(
        "[split big export] ЖД-перевозок расчетного периода: ",
        tables[0].get("rail_registered_total", 0),
        sep="",
    )
    print(
        "[split big export] ЖД-транзит по начальным пунктам: ",
        tables[0].get("rail_rows", [{}])[-1].get("Итого", 0),
        sep="",
    )
    print(
        "[split big export] Использованных НП расчетного периода: ",
        table4.get("used_np_total", 0),
        sep="",
    )
    return tables, monthly_np_snapshot, 0


def _merge_hybrid_tables(big_tables, old_tables):
    """Собрать режим 4: Т1-Т3 из большой выгрузки, Т4-Т6 старым методом."""
    big_by_name = {table.get("table"): table for table in big_tables}
    old_by_name = {table.get("table"): table for table in old_tables}
    required = (
        "Таблица1", "Таблица2", "Таблица3",
        "Таблица4", "Таблица5", "Таблица6",
    )

    missing = [
        name
        for name in required
        if name not in (
            big_by_name
            if name in {"Таблица1", "Таблица2", "Таблица3"}
            else old_by_name
        )
    ]
    if missing:
        raise ValueError(
            "Не удалось собрать смешанный отчет. Нет таблиц: "
            + ", ".join(missing)
        )

    return [
        big_by_name["Таблица1"],
        big_by_name["Таблица2"],
        big_by_name["Таблица3"],
        old_by_name["Таблица4"],
        old_by_name["Таблица5"],
        old_by_name["Таблица6"],
        big_by_name.get(
            "Таблица7",
            old_by_name.get(
                "Таблица7",
                {
                    "table": "Таблица7",
                    "stage": "legacy-empty",
                    "source": "old",
                    "auto_rows": [],
                    "rail_rows": [],
                },
            ),
        ),
    ]


def _write_report(
    out_path: Path,
    power_bi_out_path: Path,
    tables,
    method_name: str,
) -> int:
    try:
        write_output(out_path, tables=tables)
    except Exception as exc:
        print("Ошибка: не удалось записать файл результата.")
        print(exc)
        return 3
    try:
        write_power_bi_output(
            power_bi_out_path,
            tables,
            report_date=dt.date.today(),
            method=method_name,
            point_registry_path=resource_path("data/border_crossing_points.json"),
        )
    except Exception as exc:
        print("Ошибка: не удалось записать данные для Power BI.")
        print(exc)
        return 3
    print("[main] output created:", out_path.name)
    return 0


def _move_unique_processed(paths: list[Path | None], in_dir: Path) -> None:
    seen: set[Path] = set()
    for path in paths:
        if path is None:
            continue
        resolved = path.resolve()
        if resolved in seen or not path.exists():
            continue
        seen.add(resolved)
        _move_processed(path, in_dir)


def main() -> int:

    in_dir = app_path("in")
    out_dir = app_path("out")
    reports_dir = out_dir / "results"
    power_bi_dir = out_dir / "power bi data"

    in_dir.mkdir(exist_ok=True)
    out_dir.mkdir(exist_ok=True)
    reports_dir.mkdir(exist_ok=True)
    power_bi_dir.mkdir(exist_ok=True)

    print(f"[main] start version {APP_VERSION}")
    method = _choose_method()
    if method == "0":
        print("[main] Выход из программы.")
        return 0
    table4_file = _find_table4_file(in_dir)
    old_src_file = (
        _find_optional_input_file(in_dir)
        if method == "6"
        else _find_input_file(in_dir) if method in {"1", "3", "4", "5"} else None
    )
    rail_file = _find_rail_file(in_dir) if method in {"1", "4"} else None
    big_export_file = (
        find_raw04_export(in_dir)
        if method == "6"
        else find_big_export_file(in_dir) if method in {"2", "3", "4", "5"} else None
    )

    if method in {"1", "3", "4", "5"} and old_src_file is None:
        return 1
    if method in {"2", "3", "4", "5"} and big_export_file is None:
        return 1
    if method == "6" and big_export_file is None:
        return 1
    if method == "4" and rail_file is None:
        return 1
    if method in {"3", "4", "5"} and table4_file is None:
        return 1

    missing_optional: list[str] = []
    if table4_file is None and method != "6":
        missing_optional.append("Готовность АПП_ЖД_Территория.xlsx")
    if method in {"1", "4"} and rail_file is None:
        missing_optional.append("ЖД Выгрузка.xlsx")
    if method == "6" and table4_file is None:
        print(
            "[raw04] ВНИМАНИЕ: файл готовности АПП/ЖД отсутствует. "
            "Если report есть, Таблица8 будет сформирована без части данных."
        )
    if not _confirm_continue_without(missing_optional):
        print("[main] execution cancelled by user")
        return 1
    table4_source = _read_table4_source(table4_file)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    processed_files: list[Path | None] = []
    old_snapshot = None

    if method == "1":
        print("\n[main] Запуск старого метода")
        old_tables, old_snapshot, code = _build_old_tables(
            old_src_file,
            table4_source,
            rail_file,
        )
        if code != 0 or old_tables is None:
            return code
        code = _write_report(
            _make_output_filename(reports_dir, "", timestamp),
            _make_output_filename(power_bi_dir, "", timestamp),
            old_tables,
            "1 - Старый метод",
        )
        if code != 0:
            return code
        processed_files.extend([old_src_file, rail_file])

    if method == "2":
        print("\n[main] Запуск нового метода")
        new_tables, code = _build_new_tables(big_export_file, table4_source)
        if code != 0 or new_tables is None:
            return code
        code = _write_report(
            _make_output_filename(reports_dir, "new method", timestamp),
            _make_output_filename(power_bi_dir, "new method", timestamp),
            new_tables,
            "2 - Новый метод",
        )
        if code != 0:
            return code
        processed_files.append(big_export_file)

    if method == "3":
        print("\n[main] Запуск метода с разделением авто/ЖД")
        combined_tables, old_snapshot, code = _build_split_big_export_tables(
            big_export_file,
            table4_source,
            old_src_file,
        )
        if code != 0 or combined_tables is None:
            return code
        code = _write_report(
            _make_output_filename(reports_dir, "combined method", timestamp),
            _make_output_filename(power_bi_dir, "combined method", timestamp),
            combined_tables,
            "3 - Большая выгрузка с разделением авто/ЖД",
        )
        if code != 0:
            return code
        processed_files.extend([big_export_file, old_src_file])

    if method == "4":
        print("\n[main] Запуск смешанного метода")
        old_tables, old_snapshot, code = _build_old_tables(
            old_src_file,
            table4_source,
            rail_file,
        )
        if code != 0 or old_tables is None:
            return code

        big_tables, code = _build_combined_tables(
            big_export_file,
            table4_source,
            rail_file,
        )
        if code != 0 or big_tables is None:
            return code

        try:
            hybrid_tables = _merge_hybrid_tables(big_tables, old_tables)
        except Exception as exc:
            print("Ошибка: не удалось собрать смешанный отчет.")
            print(exc)
            return 7

        code = _write_report(
            _make_output_filename(reports_dir, "hybrid method", timestamp),
            _make_output_filename(power_bi_dir, "hybrid method", timestamp),
            hybrid_tables,
            "4 - Смешанный метод",
        )
        if code != 0:
            return code
        processed_files.extend([old_src_file, big_export_file, rail_file])

    if method == "5":
        print("\n[main] Запуск единого периода с 11.02.2026")
        full_period_tables, old_snapshot, code = _build_split_big_export_tables(
            big_export_file,
            table4_source,
            old_src_file,
            period_start=FULL_PERIOD_START,
            subtract_stage2_baseline=False,
        )
        if code != 0 or full_period_tables is None:
            return code
        code = _write_report(
            _make_full_period_output_filename(reports_dir, timestamp),
            _make_full_period_output_filename(power_bi_dir, timestamp),
            full_period_tables,
            "5 - Единый период с 11.02.2026",
        )
        if code != 0:
            return code
        processed_files.extend([big_export_file, old_src_file])

    if method == "6":
        print("\n[main] Запуск ежедневной справки RAW04")
        print("[raw04] Таблицы 1, 2, 4, 5: основной расчет с 11.02.2026 (авто/ЖД)")
        if old_src_file is not None:
            base_tables, old_snapshot, code = _build_split_big_export_tables(
                big_export_file,
                table4_source,
                old_src_file,
                period_start=FULL_PERIOD_START,
                # Таблица8 сохраняет прежний расчет второго этапа и месяца.
                subtract_stage2_baseline=True,
            )
        else:
            base_tables, code = _build_raw04_primary_without_report(
                big_export_file,
            )
        if code != 0 or base_tables is None:
            return code
        raw04_tables, code = _build_daily_raw04_tables(big_export_file)
        if code != 0 or raw04_tables is None:
            return code
        try:
            daily_tables = _merge_primary_and_raw04_tables(base_tables, raw04_tables)
        except Exception as exc:
            print("Ошибка: не удалось объединить старые и новые таблицы RAW04.")
            print(exc)
            return 8
        daily_output = _make_daily_raw04_output_filename(reports_dir, timestamp)
        try:
            # Power BI is intentionally not generated here: the user-facing
            # certificate combines legacy and RAW04-only sheets.
            print("[raw04] Шаг 3 из 3: записываются листы Excel")
            write_output(daily_output, tables=daily_tables)
        except Exception as exc:
            print("Ошибка: не удалось записать ежедневную справку RAW04.")
            print(exc)
            return 3
        print("[main] RAW04 output created:", daily_output.name)
        processed_files.extend([old_src_file, big_export_file])

    if old_snapshot is not None:
        try:
            snapshot_path = save_snapshot(
                monthly_np_storage_path(),
                old_snapshot,
            )
            print("[main] monthly NP snapshot saved:", snapshot_path.name)
        except Exception as exc:
            print("Предупреждение: не удалось сохранить месячный снимок пломб.")
            print(exc)

    processed_files.append(table4_file)
    try:
        _move_unique_processed(processed_files, in_dir)
        print("[main] source files moved to in/processed")
    except Exception as exc:
        print("Предупреждение: не удалось переместить исходный файл.")
        print(exc)

    print("[main] done")
    return 0


if __name__ == "__main__":

    try:

        code = main()

        if code != 0:
            _pause()

    except Exception:

        print("\nОшибка выполнения программы.")
        print("Проверьте выбранный режим и входные файлы в папке in")
        _pause()
