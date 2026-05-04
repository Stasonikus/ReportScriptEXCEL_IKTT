from __future__ import annotations  # позволяет использовать аннотации типов как строки (удобно для Python <3.10)

import datetime as dt  # работа с датой и временем
import re  # регулярные выражения (поиск по шаблону)
import shutil  # операции с файлами (перемещение и т.д.)
from pathlib import Path  # удобная работа с путями (вместо os.path)

# импорт функций построения таблиц
from builders import build_table1, build_table2, build_table3, build_table5
# загрузка конфигурации (списки, настройки нормализации)
from config import load_config
# функции чтения и записи Excel
from io_excel import read_report_1, read_report_2, write_output
# нормализация данных (приведение к единому виду)
from normalize import normalize_report


# регулярка: ищем файлы вида report_*.xlsx (без учета регистра)
REPORT_PATTERN = re.compile(r"^report_.*\.xlsx$", re.IGNORECASE)


def _make_output_filename(out_dir: Path) -> Path:
    # создаём текущую дату/время в формате YYYYMMDD_HHMM
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M")
    # формируем путь: out/report_дата.xlsx
    return out_dir / f"report_{ts}.xlsx"


def _pause():
    # пауза в консоли (чтобы окно не закрылось сразу)
    input("\nНажмите Enter для выхода...")


def _find_input_file(in_dir: Path) -> Path | None:

    files = []  # сюда будем складывать найденные файлы

    # перебираем все Excel файлы в папке in
    for f in in_dir.glob("*.xlsx"):
        name = f.name  # имя файла

        # игнорируем временные файлы Excel (~$...)
        if name.startswith("~$"):
            continue

        # проверяем подходит ли под шаблон report_*.xlsx
        if REPORT_PATTERN.match(name):
            files.append(f)

    # если не нашли ни одного файла
    if len(files) == 0:
        print("Ошибка: файл источника не найден.")
        print("В папке in должен находиться один файл report_*.xlsx")
        return None

    # если нашли больше одного файла
    if len(files) > 1:
        print("Ошибка: найдено несколько файлов источника.")
        print("В папке in должен находиться только один файл report_*.xlsx\n")

        # выводим список найденных файлов
        for f in files:
            print(" -", f.name)

        return None

    # если всё ок — возвращаем единственный файл
    return files[0]


def _move_processed(src: Path, in_dir: Path):

    # создаём папку in/processed если её нет
    processed_dir = in_dir / "processed"
    processed_dir.mkdir(exist_ok=True)

    # текущая дата/время
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M")

    # новое имя файла с пометкой processed
    new_name = src.stem + f"__processed_{ts}" + src.suffix
    dst = processed_dir / new_name

    # перемещаем файл
    shutil.move(str(src), str(dst))


def main() -> int:

    # папка входящих файлов
    in_dir = Path("in")
    # папка для результата
    out_dir = Path("out")

    # создаём папки если их нет
    in_dir.mkdir(exist_ok=True)
    out_dir.mkdir(exist_ok=True)

    print("[main] start")  # старт программы

    # ищем входной файл
    src_file = _find_input_file(in_dir)

    # если не нашли — ошибка
    if src_file is None:
        return 1

    print("[main] source =", src_file.name)  # выводим имя файла

    # создаём имя выходного файла
    out_path = _make_output_filename(out_dir)

    # загружаем конфиг (списки ПП, правила нормализации)
    cfg = load_config()

    # -----------------------
    # read source (чтение Excel)
    # -----------------------

    try:
        # читаем первый отчёт (основной)
        r1 = read_report_1(src_file, sheet_name="Отчёт")
        print("[check] report read: OK")
    except Exception:
        # если ошибка — файл не подходит
        print("Ошибка: неверный файл источника.")
        print("Проверьте структуру файла report_*.xlsx")
        return 2

    try:
        # второй отчёт (сейчас заглушка)
        _ = read_report_2(src_file)
    except Exception:
        pass  # игнорируем ошибки (не критично)

    # -----------------------
    # normalize (нормализация данных)
    # -----------------------

    r1_norm = normalize_report(
        r1,
        cfg.normalization,           # правила нормализации
        allowed_pp=cfg.tp_zhd_list_t1,  # список допустимых ПП
    )

    print("[main] normalize step: OK")

    # -----------------------
    # build tables (формирование таблиц)
    # -----------------------

    # таблица 1 (основная по ПП)
    t1 = build_table1(
        r1_norm,
        cfg.normalization,
        cfg.tp_zhd_list_t1,
    )

    print("[main] build_table1 ->")
    # вывод строк таблицы в консоль (для дебага)
    for row in t1["rows"]:
        print(row)

    # остальные таблицы
    t2 = build_table2(r1_norm, cfg.normalization)
    t3 = build_table3(r1_norm, cfg.normalization)
    t5 = build_table5(r1_norm, cfg.normalization)

    # -----------------------
    # write output (запись Excel)
    # -----------------------

    try:
        # записываем все таблицы в файл
        write_output(out_path, tables=[t1, t2, t3, t5])
    except Exception:
        print("Ошибка: не удалось записать файл результата.")
        return 3

    print("[main] output created:", out_path.name)

    # вывод статистики по таблице 1
    print(f"[main] rows written to Таблица1: {len(t1['rows'])}")

    # вывод статистики по таблице 3
    print(
        "[main] rows written to Таблица3 "
        f"(auto_out/auto_in/rail_out/rail_in): "
        f"{len(t3.get('auto_out_rows', []))}/"
        f"{len(t3.get('auto_in_rows', []))}/"
        f"{len(t3.get('rail_out_rows', []))}/"
        f"{len(t3.get('rail_in_rows', []))}"
    )

    # -----------------------
    # move processed file (перемещение исходника)
    # -----------------------

    try:
        # переносим файл в папку processed
        _move_processed(src_file, in_dir)
        print("[main] source moved to in/processed")
    except Exception:
        print("Предупреждение: не удалось переместить исходный файл.")

    print("[main] done")  # завершение

    return 0  # успешное завершение


if __name__ == "__main__":

    try:
        # запускаем main
        code = main()

        # если код ошибки — ждём Enter
        if code != 0:
            _pause()

    except Exception:
        # глобальный перехват ошибок
        print("\nОшибка выполнения программы.")
        print("Проверьте файл источника report_*.xlsx")
        _pause()