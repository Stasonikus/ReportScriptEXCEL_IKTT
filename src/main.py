from __future__ import annotations

import datetime as dt
import re
import shutil
from pathlib import Path

from builders import (
    build_table1,
    build_table2,
    build_table3,
    build_table4,
    build_table5,
)

from config import load_config

from io_excel import (
    read_rail_report,
    read_report_1,
    read_report_2,
    write_output,
)

from normalize import normalize_report
from utils import app_path, resource_path


REPORT_PATTERN = re.compile(r"^report_.*\.xlsx$", re.IGNORECASE)

TABLE4_PATTERN = re.compile(
    r"^готовность.*апп.*территория.*\.xlsx$",
    re.IGNORECASE,
)

RAIL_PATTERN = re.compile(
    r"^жд\s*выгрузка.*\.xlsx$",
    re.IGNORECASE,
)


def _make_output_filename(out_dir: Path) -> Path:
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M")
    return out_dir / f"report_{ts}.xlsx"


def _pause():
    input("\nНажмите Enter для выхода...")


def _find_input_file(in_dir: Path) -> Path | None:

    files = []

    for f in in_dir.glob("*.xlsx"):

        name = f.name

        if name.startswith("~$"):
            continue

        if REPORT_PATTERN.match(name):
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

        print("[main] Таблица4 будет пропущена")
        return None

    print(f"[main] table4 source FOUND: {found[0].name}")

    return found[0]


def _find_rail_file(in_dir: Path) -> Path | None:

    found = []

    for f in in_dir.glob("*.xlsx"):

        name = f.name

        if name.startswith("~$"):
            continue

        if RAIL_PATTERN.match(name):
            found.append(f)

    if len(found) == 0:
        print("[main] ЖД Выгрузка.xlsx NOT FOUND")
        print("[main] Таблица3 будет создана без ЖД-уточнения")
        return None

    if len(found) > 1:
        print("[main] rail source: found multiple files")
        for f in found:
            print(" -", f.name)

        print("[main] Таблица3 будет создана без ЖД-уточнения")
        return None

    print(f"[main] rail source FOUND: {found[0].name}")

    return found[0]


def _move_processed(src: Path, in_dir: Path):

    processed_dir = in_dir / "processed"
    processed_dir.mkdir(exist_ok=True)

    ts = dt.datetime.now().strftime("%Y%m%d_%H%M")

    new_name = src.stem + f"__processed_{ts}" + src.suffix
    dst = processed_dir / new_name

    shutil.move(str(src), str(dst))


def main() -> int:

    in_dir = app_path("in")
    out_dir = app_path("out")

    in_dir.mkdir(exist_ok=True)
    out_dir.mkdir(exist_ok=True)

    print("[main] start")

    src_file = _find_input_file(in_dir)

    if src_file is None:
        return 1

    print("[main] source =", src_file.name)

    table4_file = _find_table4_file(in_dir)
    rail_file = _find_rail_file(in_dir)

    out_path = _make_output_filename(out_dir)

    cfg = load_config()

    # -----------------------
    # read source
    # -----------------------

    try:

        r1 = read_report_1(src_file, sheet_name="Отчёт")
        print("[check] report read: OK")

    except Exception:

        print("Ошибка: неверный файл источника.")
        print("Проверьте структуру файла report_*.xlsx")
        return 2

    table4_source = None

    if table4_file is not None:

        try:

            table4_source = read_report_2(table4_file)

            print("[main] table4 report read: OK")

        except Exception:

            print("[main] table4 report read: FAIL")
            table4_source = None

    rail_source = None

    if rail_file is not None:

        try:

            rail_source = read_rail_report(rail_file)

            print("[main] rail report read: OK")

        except Exception as e:

            print("[main] rail report read: FAIL")
            print(e)
            print("[main] Таблица3 будет создана без ЖД-уточнения")
            rail_source = None

    # -----------------------
    # normalize
    # -----------------------

    r1_norm = normalize_report(
        r1,
        cfg.normalization,
        allowed_pp=cfg.tp_zhd_list_t1,
    )

    print("[main] normalize step: OK")

    # -----------------------
    # build tables
    # -----------------------

    t1 = build_table1(
        r1_norm,
        cfg.normalization,
        cfg.tp_zhd_list_t1,
    )

    print("[main] build_table1 ->")

    for row in t1["rows"]:
        print(row)

    t2 = build_table2(r1_norm, cfg.normalization)

    t3 = build_table3(
        r1_norm,
        cfg.normalization,
        rail_raw=rail_source,
        data_dir=resource_path("data"),
    )

    t5 = build_table5(r1_norm, cfg.normalization)

    tables = [t1, t2, t3]

    # -----------------------
    # table 4
    # -----------------------

    if table4_source is not None:

        try:

            t4 = build_table4(
                r1_norm,
                table4_source,
                cfg.normalization,
            )

            tables.append(t4)

            print("[main] Таблица4 created")

        except Exception as e:

            print("[main] Таблица4 build FAIL")
            print(e)

    else:

        print("[main] Таблица4 skipped")

    tables.append(t5)

    # -----------------------
    # write output
    # -----------------------

    try:

        write_output(out_path, tables=tables)

    except Exception:

        print("Ошибка: не удалось записать файл результата.")
        return 3

    print("[main] output created:", out_path.name)

    print(f"[main] rows written to Таблица1: {len(t1['rows'])}")

    print(
        "[main] rows written to Таблица3 "
        f"(auto_out/auto_in/rail_out/rail_in): "
        f"{len(t3.get('auto_out_rows', []))}/"
        f"{len(t3.get('auto_in_rows', []))}/"
        f"{len(t3.get('rail_out_rows', []))}/"
        f"{len(t3.get('rail_in_rows', []))}"
    )

    # -----------------------
    # move processed
    # -----------------------

    try:

        _move_processed(src_file, in_dir)

        if table4_file is not None:
            _move_processed(table4_file, in_dir)

        if rail_file is not None:
            _move_processed(rail_file, in_dir)

        print("[main] source moved to in/processed")

    except Exception:

        print("Предупреждение: не удалось переместить исходный файл.")

    print("[main] done")

    return 0


if __name__ == "__main__":

    try:

        code = main()

        if code != 0:
            _pause()

    except Exception:

        print("\nОшибка выполнения программы.")
        print("Проверьте файл источника report_*.xlsx")
        _pause()
