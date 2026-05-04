from __future__ import annotations  # позволяет использовать аннотации типов как строки

from pathlib import Path  # удобная работа с путями
from typing import Any, Dict, List  # типизация

# библиотека для работы с Excel
from openpyxl import Workbook, load_workbook
# стили для ячеек
from openpyxl.styles import Font, Alignment
# утилита для перевода номера колонки в букву (1 -> A)
from openpyxl.utils import get_column_letter


def read_report_1(path: Path, sheet_name: str = "Отчёт") -> Dict[str, Any]:
    # проверяем существует ли файл
    if not path.exists():
        raise FileNotFoundError(f"report_1 not found: {path}")

    # загружаем Excel, data_only=True — берём значения формул, а не сами формулы
    wb = load_workbook(path, data_only=True)

    # проверяем есть ли нужный лист
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"sheet not found: {sheet_name}")

    # берём лист
    ws = wb[sheet_name]

    # возвращаем структуру с данными
    return {
        "path": str(path),          # путь к файлу
        "workbook": wb,            # сам workbook
        "ws": ws,                  # лист
        "sheet_name": sheet_name,  # имя листа
        "sheetnames": wb.sheetnames,  # список всех листов
    }


def read_report_2(path: Path) -> Dict[str, Any]:
    # проверка файла
    if not path.exists():
        raise FileNotFoundError(f"report_2 not found: {path}")

    # пока заглушка (данные не читаются)
    return {"path": str(path), "data": None}


def write_output_test(out_path: Path, sheet_name: str = "Таблица1") -> None:
    # создаём папку если её нет
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # создаём новый Excel файл
    wb = Workbook()
    ws = wb.active  # берём активный лист
    ws.title = sheet_name  # переименовываем лист
    ws["A1"] = "OK"  # пишем тестовое значение
    wb.save(out_path)  # сохраняем файл


def _safe_sheet_title(name: str) -> str:
    # запрещённые символы для Excel
    bad = [":", "\\", "/", "?", "*", "[", "]"]

    # заменяем их на пробел
    for ch in bad:
        name = name.replace(ch, " ")

    name = name.strip()  # убираем пробелы по краям

    # Excel ограничивает имя листа 31 символом
    if len(name) > 31:
        name = name[:31]

    # если пусто — ставим Sheet
    return name or "Sheet"


def _delete_default_sheet_if_needed(wb: Workbook) -> None:
    # если есть только один лист и он называется "Sheet"
    if len(wb.worksheets) == 1 and wb.worksheets[0].title == "Sheet":
        # удаляем его
        wb.remove(wb.worksheets[0])


def _set_col_widths(ws, headers: List[str], start_col: int = 1, max_rows_scan: int = 200) -> None:
    # перебираем колонки
    for i, h in enumerate(headers, start=start_col):
        letter = get_column_letter(i)  # получаем букву колонки (A, B, C...)

        max_len = len(str(h))  # начальная длина = длина заголовка

        # смотрим значения в колонке (до max_rows_scan строк)
        for r in range(1, min(ws.max_row, max_rows_scan) + 1):
            v = ws.cell(row=r, column=i).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))  # ищем максимальную длину

        # устанавливаем ширину (минимум 10, максимум 55)
        ws.column_dimensions[letter].width = min(max(10, max_len + 2), 55)


def _write_rows_as_table(ws, headers: List[str], rows: List[Dict[str, Any]], start_row: int = 1, start_col: int = 1) -> int:
    # записываем заголовки
    for j, h in enumerate(headers, start=start_col):
        c = ws.cell(row=start_row, column=j, value=h)
        c.font = Font(bold=True)  # жирный текст
        c.alignment = Alignment(horizontal="center", vertical="center")  # центрирование

    r_out = start_row + 1  # следующая строка после заголовка

    # записываем строки данных
    for row in rows:
        for j, h in enumerate(headers, start=start_col):
            ws.cell(row=r_out, column=j, value=row.get(h, None))  # берём значение по ключу
        r_out += 1

    return r_out  # возвращаем номер следующей свободной строки


def _write_section_title(ws, title: str, row: int, start_col: int = 1, span_cols: int = 3) -> int:
    # записываем заголовок секции
    c = ws.cell(row=row, column=start_col, value=title)
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal="center", vertical="center")

    # если нужно растянуть на несколько колонок
    if span_cols > 1:
        ws.merge_cells(
            start_row=row,
            start_column=start_col,
            end_row=row,
            end_column=start_col + span_cols - 1,
        )

    return row + 1  # возвращаем следующую строку


def _write_table2(ws, t: Dict[str, Any]) -> None:
    # берём данные или пустые списки
    export_rows = t.get("export_rows", []) or []
    trade_rows = t.get("trade_rows", []) or []

    headers = ["№", "Страна", "Количество перевозок"]

    row = 1

    # основной заголовок
    row = _write_section_title(ws, "Статистические данные экспорта и взаимной торговли из РК", row, span_cols=3)

    # блок "Экспорт"
    row = _write_section_title(ws, "Экспорт", row, span_cols=3)
    row = _write_rows_as_table(ws, headers=headers, rows=export_rows, start_row=row)

    row += 1  # пустая строка

    # блок "Взаимная торговля"
    row = _write_section_title(ws, "Взаимная торговля", row, span_cols=3)
    row = _write_rows_as_table(ws, headers=headers, rows=trade_rows, start_row=row)

    ws.freeze_panes = "A4"  # закрепляем верхние строки
    _set_col_widths(ws, headers)  # автоширина колонок


def _write_table3(ws, t: Dict[str, Any]) -> None:
    # разные типы данных
    auto_out_rows = t.get("auto_out_rows", []) or []
    auto_in_rows = t.get("auto_in_rows", []) or []

    rail_out_rows = t.get("rail_out_rows", []) or []
    rail_in_rows = t.get("rail_in_rows", []) or []

    headers = ["Направление", "Количество"]

    row = 1

    # основной заголовок
    row = _write_section_title(
        ws,
        "Статистические данные завершенных перевозок",
        row,
        span_cols=len(headers),
    )

    # Авто из РК
    row = _write_section_title(
        ws,
        "Завершённые автомобильные перевозки из Республики Казахстан",
        row,
        span_cols=len(headers),
    )
    row = _write_rows_as_table(ws, headers, auto_out_rows, row)

    row += 1

    # Авто в РК
    row = _write_section_title(
        ws,
        "Завершенные автомобильные перевозки в Республику Казахстан",
        row,
        span_cols=len(headers),
    )
    row = _write_rows_as_table(ws, headers, auto_in_rows, row)

    row += 1

    # ЖД из РК
    row = _write_section_title(
        ws,
        "Завершённые железнодорожные перевозки из Республики Казахстан",
        row,
        span_cols=len(headers),
    )
    row = _write_rows_as_table(ws, headers, rail_out_rows, row)

    row += 1

    # ЖД в РК
    row = _write_section_title(
        ws,
        "Завершенные железнодорожные перевозки в Республику Казахстан",
        row,
        span_cols=len(headers),
    )
    row = _write_rows_as_table(ws, headers, rail_in_rows, row)

    ws.freeze_panes = "A4"
    _set_col_widths(ws, headers)


def _write_table5(ws, t: Dict[str, Any]) -> None:
    rows = t.get("rows", []) or []

    headers = ["Страна начала перевозки", "Количество перевозок"]

    row = 1

    # заголовок
    row = _write_section_title(
        ws,
        "Статистика по перевозкам, направленным в сторону Республики Казахстан",
        row,
        span_cols=2,
    )

    # таблица
    _write_rows_as_table(ws, headers=headers, rows=rows, start_row=row)

    ws.freeze_panes = "A3"
    _set_col_widths(ws, headers)


def write_output(out_path: Path, tables: List[Any]) -> None:

    # создаём папку
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # если таблиц нет — создаём тестовый файл
    if not tables:
        write_output_test(out_path, sheet_name="Таблица1")
        print(f"[io_excel.write_output] tables_stub_count=0 (created {out_path} with Таблица1:A1='OK')")
        return

    # создаём Excel
    wb = Workbook()
    _delete_default_sheet_if_needed(wb)

    # перебираем таблицы
    for t in tables:

        if not isinstance(t, dict):
            continue

        # имя листа
        sheet_title = _safe_sheet_title(str(t.get("table", "Sheet")))
        ws = wb.create_sheet(title=sheet_title)

        # спец обработка таблиц
        if sheet_title == "Таблица2":
            _write_table2(ws, t)
            continue

        if sheet_title == "Таблица3":
            _write_table3(ws, t)
            continue

        if sheet_title == "Таблица5":
            _write_table5(ws, t)
            continue

        rows = t.get("rows", [])

        # если формат не тот — просто пишем как строку
        if not isinstance(rows, list) or (rows and not isinstance(rows[0], dict)):
            ws["A1"] = str(t)
            continue

        # заголовки
        if sheet_title == "Таблица1":
            headers = ["ТП/ЖД", "РФ", "РБ", "КР", "РА", "Итого"]
        else:
            headers = list(rows[0].keys()) if rows else ["data"]

        # пишем таблицу
        _write_rows_as_table(ws, headers, rows)

        ws.freeze_panes = "A2"
        _set_col_widths(ws, headers)

    # сохраняем файл
    wb.save(out_path)

    print(f"[io_excel.write_output] wrote {len(tables)} table(s) to: {out_path}")