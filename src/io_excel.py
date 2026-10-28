from __future__ import annotations

import datetime as dt
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Dict, List

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, PatternFill, Side
from openpyxl.utils import get_column_letter


OUTPUT_FONT_NAME = "Times New Roman"
OUTPUT_FONT_SIZE = 14
MIN_ROW_HEIGHT = 21
MAX_AUTOFIT_WIDTH = 90
PILOT_KARASU_NP = 391
STAGE2_START_DATE = dt.date(2026, 9, 1)


def _stage2_elapsed_days(current_date: dt.date | None = None) -> int:
    """Return the monthly day counter used for the daily NP calculation."""
    report_date = current_date or dt.date.today()
    if report_date <= STAGE2_START_DATE:
        return 1
    if (report_date.year, report_date.month) == (
        STAGE2_START_DATE.year,
        STAGE2_START_DATE.month,
    ):
        return max(1, (report_date - STAGE2_START_DATE).days + 1)
    return report_date.day


def _output_font(bold: bool = False) -> Font:
    return Font(
        name=OUTPUT_FONT_NAME,
        size=OUTPUT_FONT_SIZE,
        bold=bold,
    )


def read_report_1(path: Path, sheet_name: str = "Отчёт") -> Dict[str, Any]:

    if not path.exists():
        raise FileNotFoundError(f"report_1 not found: {path}")

    wb = load_workbook(path, data_only=True)

    if sheet_name not in wb.sheetnames:
        raise KeyError(f"sheet not found: {sheet_name}")

    ws = wb[sheet_name]

    return {
        "path": str(path),
        "workbook": wb,
        "ws": ws,
        "sheet_name": sheet_name,
        "sheetnames": wb.sheetnames,
    }


def read_report_2(path: Path) -> Dict[str, Any]:

    if not path.exists():
        raise FileNotFoundError(f"report_2 not found: {path}")

    wb = load_workbook(path, data_only=True)

    target_sheet = None

    for s in wb.sheetnames:

        low = s.strip().lower()

        if "распределение" in low and "нп" in low:
            target_sheet = s
            break

    if target_sheet is None:
        target_sheet = wb.sheetnames[0]

    ws = wb[target_sheet]

    return {
        "path": str(path),
        "workbook": wb,
        "ws": ws,
        "sheet_name": target_sheet,
        "sheetnames": wb.sheetnames,
        "worksheets": wb.worksheets,
    }


def read_rail_report(path: Path) -> Dict[str, Any]:

    if not path.exists():
        raise FileNotFoundError(f"rail report not found: {path}")

    if path.suffix.lower() == ".xls":
        return _read_rail_xls_skipping_logs(path)

    wb = load_workbook(path, data_only=True)
    ws = wb.active

    return {
        "path": str(path),
        "workbook": wb,
        "ws": ws,
        "sheet_name": ws.title,
        "sheetnames": wb.sheetnames,
    }


def _read_rail_xls_skipping_logs(path: Path) -> Dict[str, Any]:

    raw = path.read_bytes()

    if raw.lstrip()[:5].lower() == b"<html":
        rows = _read_html_table_rows(raw)
        return _rail_rows_to_workbook(path, rows, sheet_name="Rail")

    import xlrd

    source_book = xlrd.open_workbook(str(path))
    source_sheet = source_book.sheet_by_index(0)

    rows = [
        [
            source_sheet.cell_value(source_row, source_col)
            for source_col in range(source_sheet.ncols)
        ]
        for source_row in range(source_sheet.nrows)
    ]

    return _rail_rows_to_workbook(path, rows, sheet_name=source_sheet.name)


class _HtmlTableParser(HTMLParser):

    def __init__(self) -> None:
        super().__init__()
        self.rows: List[List[str]] = []
        self._in_row = False
        self._in_cell = False
        self._cell_parts: List[str] = []
        self._row: List[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()

        if tag == "tr":
            self._in_row = True
            self._row = []
            return

        if tag in {"td", "th"} and self._in_row:
            self._in_cell = True
            self._cell_parts = []

    def handle_data(self, data: str) -> None:
        if self._in_cell:
            self._cell_parts.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()

        if tag in {"td", "th"} and self._in_cell:
            value = " ".join("".join(self._cell_parts).split())
            self._row.append(value)
            self._in_cell = False
            self._cell_parts = []
            return

        if tag == "tr" and self._in_row:
            if self._row:
                self.rows.append(self._row)

            self._in_row = False
            self._row = []


def _read_html_table_rows(raw: bytes) -> List[List[str]]:

    parser = _HtmlTableParser()
    parser.feed(raw.decode("utf-8", errors="replace"))

    return parser.rows


def _rail_rows_to_workbook(
    path: Path,
    rows: List[List[Any]],
    sheet_name: str,
) -> Dict[str, Any]:

    wb = Workbook()
    ws = wb.active
    ws.title = _safe_sheet_title(sheet_name)

    source_start_row = _find_rail_header_row_index(rows)

    if source_start_row is None:
        source_start_row = 16

    for source_row in range(source_start_row, len(rows)):
        target_row = source_row - source_start_row + 1
        row = rows[source_row]

        for source_col, value in enumerate(row):
            ws.cell(
                row=target_row,
                column=source_col + 1,
                value=value,
            )

    return {
        "path": str(path),
        "workbook": wb,
        "ws": ws,
        "sheet_name": ws.title,
        "sheetnames": [ws.title],
    }


def _find_rail_header_row_index(rows: List[List[Any]]) -> int | None:

    required = {
        "владелец записи",
        "статус",
        "company_id",
        "start waypoint",
        "end waypoint",
    }

    for idx, row in enumerate(rows):
        values = {
            str(value).strip().lower()
            for value in row
            if value is not None and str(value).strip()
        }

        if required.issubset(values):
            return idx

    return None


def write_output_test(out_path: Path, sheet_name: str = "Таблица1") -> None:

    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()

    ws = wb.active

    ws.title = sheet_name

    ws["A1"] = "OK"

    wb.save(out_path)


def _safe_sheet_title(name: str) -> str:

    bad = [":", "\\", "/", "?", "*", "[", "]"]

    for ch in bad:
        name = name.replace(ch, " ")

    name = name.strip()

    if len(name) > 31:
        name = name[:31]

    return name or "Sheet"


def _delete_default_sheet_if_needed(wb: Workbook) -> None:

    if len(wb.worksheets) == 1 and wb.worksheets[0].title == "Sheet":
        wb.remove(wb.worksheets[0])


def _apply_output_font_to_used_range(ws, force_bold: bool | None = None) -> None:

    for row in ws.iter_rows(
        min_row=1,
        max_row=ws.max_row,
        min_col=1,
        max_col=ws.max_column,
    ):
        for cell in row:
            is_bold = bool(cell.font and cell.font.bold)

            if force_bold is not None:
                is_bold = force_bold

            cell.font = _output_font(bold=is_bold)


def _is_total_value(value: Any) -> bool:

    return str(value or "").strip().lower() in {"итого", "все"}


def _bold_total_rows(ws) -> None:

    for row in ws.iter_rows(
        min_row=1,
        max_row=ws.max_row,
        min_col=1,
        max_col=ws.max_column,
    ):
        if any(_is_total_value(cell.value) for cell in row):
            for cell in row:
                cell.font = _output_font(bold=True)


def _set_column_bold(ws, col_idx: int) -> None:

    for row_idx in range(1, ws.max_row + 1):
        ws.cell(row=row_idx, column=col_idx).font = _output_font(bold=True)


def _fit_used_range(ws, max_width: int = MAX_AUTOFIT_WIDTH) -> None:

    for col_idx in range(1, ws.max_column + 1):
        max_line_len = 0

        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value

            if value is None:
                continue

            lines = str(value).splitlines() or [str(value)]
            max_line_len = max(max_line_len, *(len(line) for line in lines))

        if max_line_len == 0:
            continue

        letter = get_column_letter(col_idx)
        width = max(10, int(max_line_len * 1.35) + 3)
        ws.column_dimensions[letter].width = min(width, max_width)

    for row_idx in range(1, ws.max_row + 1):
        max_lines = 1

        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            value = cell.value

            if value is None:
                continue

            text = str(value)
            explicit_lines = max(1, len(text.splitlines()))
            col_width = ws.column_dimensions[get_column_letter(col_idx)].width or 10
            wrapped_lines = max(1, int((len(text) / max(col_width, 1)) + 0.95))
            max_lines = max(max_lines, explicit_lines, wrapped_lines)

            cell.alignment = Alignment(
                horizontal=cell.alignment.horizontal,
                vertical=cell.alignment.vertical or "center",
                wrap_text=True,
            )

        ws.row_dimensions[row_idx].height = max(
            ws.row_dimensions[row_idx].height or 0,
            MIN_ROW_HEIGHT * max_lines,
        )


def _set_col_widths(
    ws,
    headers: List[str],
    start_col: int = 1,
    max_rows_scan: int = 200,
) -> None:

    for i, h in enumerate(headers, start=start_col):

        letter = get_column_letter(i)

        max_len = len(str(h))

        for r in range(1, min(ws.max_row, max_rows_scan) + 1):

            v = ws.cell(row=r, column=i).value

            if v is None:
                continue

            max_len = max(max_len, len(str(v)))

        ws.column_dimensions[letter].width = min(
            max(10, max_len + 2),
            55,
        )


def _write_rows_as_table(
    ws,
    headers: List[str],
    rows: List[Dict[str, Any]],
    start_row: int = 1,
    start_col: int = 1,
    dash_zero: bool = False,
) -> int:

    for j, h in enumerate(headers, start=start_col):

        c = ws.cell(row=start_row, column=j, value=h)

        c.font = _output_font(bold=True)

        c.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    r_out = start_row + 1

    for row in rows:

        is_summary = any(_is_total_value(value) for value in row.values())

        for j, h in enumerate(headers, start=start_col):

            value = row.get(h, None)
            if dash_zero and isinstance(value, (int, float)) and value == 0:
                value = "-"

            cell = ws.cell(
                row=r_out,
                column=j,
                value=value,
            )

            if value == "-" or isinstance(value, (int, float)):
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                )

            if is_summary:
                cell.font = _output_font(bold=True)

        r_out += 1

    return r_out


def _write_section_title(
    ws,
    title: str,
    row: int,
    start_col: int = 1,
    span_cols: int = 3,
) -> int:

    c = ws.cell(row=row, column=start_col, value=title)

    c.font = _output_font(bold=True)

    c.alignment = Alignment(
        horizontal="center",
        vertical="center",
    )

    if span_cols > 1:

        ws.merge_cells(
            start_row=row,
            start_column=start_col,
            end_row=row,
            end_column=start_col + span_cols - 1,
        )

    return row + 1


def _write_table2(ws, t: Dict[str, Any]) -> None:

    if t.get("transport_split"):
        def write_transport_section(
            title: str,
            rows: List[Dict[str, Any]],
            start_row: int,
        ) -> int:
            data_rows = rows[:4]
            total_row = rows[-1] if rows else {}
            current_row = start_row

            ws.merge_cells(
                start_row=current_row,
                start_column=1,
                end_row=current_row,
                end_column=3,
            )
            ws.cell(current_row, 1, title)
            current_row += 1

            for value_key, group_title, total_title in (
                ("Авто", "Автомобильные перевозки", "Итого (авто)"),
                ("ЖД", "Железнодорожные перевозки", "Итого (жд)"),
            ):
                ws.merge_cells(
                    start_row=current_row,
                    start_column=1,
                    end_row=current_row,
                    end_column=3,
                )
                ws.cell(current_row, 1, group_title)
                current_row += 1

                for item in data_rows:
                    ws.cell(current_row, 1, item.get("№"))
                    ws.cell(current_row, 2, item.get("Страна"))
                    ws.cell(current_row, 3, int(item.get(value_key, 0) or 0))
                    current_row += 1

                ws.merge_cells(
                    start_row=current_row,
                    start_column=1,
                    end_row=current_row,
                    end_column=2,
                )
                ws.cell(current_row, 1, total_title)
                ws.cell(current_row, 3, int(total_row.get(value_key, 0) or 0))
                current_row += 1

            ws.merge_cells(
                start_row=current_row,
                start_column=1,
                end_row=current_row,
                end_column=2,
            )
            ws.cell(current_row, 1, "Итого (сумма)")
            ws.cell(current_row, 3, int(total_row.get("Итого", 0) or 0))
            return current_row + 1

        row = write_transport_section(
            "Экспорт из РК",
            t.get("export_split_rows", []) or [],
            1,
        )
        last_row = write_transport_section(
            "Взаимная торговля из РК",
            t.get("trade_split_rows", []) or [],
            row,
        ) - 1

        thin = Side(style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        title_rows = {1, 2, 8, 15, 16, 22}
        total_rows = {7, 13, 14, 21, 27, 28}
        for row_cells in ws.iter_rows(
            min_row=1,
            max_row=last_row,
            min_col=1,
            max_col=3,
        ):
            for cell in row_cells:
                cell.border = border
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )
                if cell.row in title_rows | total_rows:
                    cell.font = _output_font(bold=True)

        ws.column_dimensions["A"].width = 5
        ws.column_dimensions["B"].width = 42
        ws.column_dimensions["C"].width = 20
        ws.freeze_panes = "A3"
        return

    export_rows = t.get("export_rows", []) or []
    trade_rows = t.get("trade_rows", []) or []

    headers = ["№", "Страна", "Количество перевозок"]

    row = 1

    row = _write_section_title(
        ws,
        "Статистические данные экспорта и взаимной торговли из РК",
        row,
        span_cols=3,
    )

    row = _write_section_title(
        ws,
        "Экспорт",
        row,
        span_cols=3,
    )

    row = _write_rows_as_table(
        ws,
        headers=headers,
        rows=export_rows,
        start_row=row,
    )

    row += 1

    row = _write_section_title(
        ws,
        "Взаимная торговля",
        row,
        span_cols=3,
    )

    row = _write_rows_as_table(
        ws,
        headers=headers,
        rows=trade_rows,
        start_row=row,
    )

    ws.freeze_panes = "A4"

    _set_col_widths(ws, headers)


def _write_table1_split(ws, t: Dict[str, Any]) -> None:
    # Поле данных называется «ТП/ЖД», но в утвержденном шаблоне его заголовок
    # визуально пустой. Поэтому ключ данных и подпись ячейки храним отдельно.
    headers = ["ТП/ЖД", "РФ", "РБ", "КР", "РА", "Итого"]
    auto_rows = [dict(item) for item in (t.get("auto_rows", []) or [])]
    rail_rows = [dict(item) for item in (t.get("rail_rows", []) or [])]
    if auto_rows:
        auto_rows[-1]["ТП/ЖД"] = "Итого (Авто)"
    if rail_rows:
        rail_rows[-1]["ТП/ЖД"] = "Итого (ЖД)"
    sum_row: Dict[str, Any] = {"ТП/ЖД": "Итого (Сумма)"}
    for header in headers[1:]:
        auto_value = auto_rows[-1].get(header, 0) if auto_rows else 0
        rail_value = rail_rows[-1].get(header, 0) if rail_rows else 0
        sum_row[header] = int(auto_value or 0) + int(rail_value or 0)
    rail_rows.append(sum_row)
    row = 1
    row = _write_section_title(
        ws,
        "Транзитные перевозки (наложение НП)",
        row,
        span_cols=len(headers),
    )
    row = _write_section_title(
        ws,
        "Автомобильные пункты пропуска",
        row,
        span_cols=len(headers),
    )
    auto_header_start = row
    row = _write_rows_as_table(
        ws,
        headers,
        auto_rows,
        row,
        dash_zero=True,
    )
    ws.cell(row=auto_header_start, column=1).value = None
    row += 1
    row = _write_section_title(
        ws,
        "Железнодорожные пункты пропуска",
        row,
        span_cols=len(headers),
    )
    rail_header_start = row
    _write_rows_as_table(
        ws,
        headers,
        rail_rows,
        row,
        dash_zero=True,
    )
    ws.cell(row=rail_header_start, column=1).value = None

    title_cell = ws.cell(row=1, column=8, value="Количество оформленных ЖД")
    value_cell = ws.cell(
        row=2,
        column=8,
        value=t.get("rail_registered_total", ""),
    )
    for cell in (title_cell, value_cell):
        cell.font = _output_font(bold=True)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
    ws.column_dimensions["H"].width = 24
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    auto_header_row = 3
    auto_total_row = auto_header_row + len(auto_rows)
    spacer_row = auto_total_row + 1
    rail_title_row = spacer_row + 1
    rail_header_row = rail_title_row + 1
    rail_sum_row = rail_header_row + len(rail_rows)
    rail_total_row = rail_sum_row - 1
    centered_rows = {
        1,
        2,
        auto_header_row,
        auto_total_row,
        rail_title_row,
        rail_header_row,
        rail_total_row,
        rail_sum_row,
    }
    for row_cells in ws.iter_rows(
        min_row=1,
        max_row=ws.max_row,
        min_col=1,
        max_col=6,
    ):
        for cell in row_cells:
            if cell.row == spacer_row:
                continue
            cell.border = border
            cell.alignment = Alignment(
                horizontal=(
                    "center"
                    if cell.row in centered_rows or cell.column != 1
                    else "left"
                ),
                vertical="center",
                wrap_text=True,
            )
    for cell in (title_cell, value_cell):
        cell.border = border
    ws.freeze_panes = "A4"
    _set_col_widths(ws, headers)
    _set_column_bold(ws, 6)


def _write_seal_post_table7(ws, t: Dict[str, Any]) -> None:
    """Write seal applications by normalized customs post."""
    headers = ["ТП/ЖД", "РФ", "РБ", "КР", "РА", "Итого"]
    auto_rows = [dict(item) for item in (t.get("auto_rows", []) or [])]
    rail_rows = [dict(item) for item in (t.get("rail_rows", []) or [])]
    if auto_rows:
        auto_rows[-1]["ТП/ЖД"] = "Итого (Авто)"
    if rail_rows:
        rail_rows[-1]["ТП/ЖД"] = "Итого (ЖД)"
    combined: Dict[str, Any] = {"ТП/ЖД": "Итого (Сумма)"}
    for header in headers[1:]:
        combined[header] = int(
            (auto_rows[-1].get(header, 0) if auto_rows else 0) or 0
        ) + int((rail_rows[-1].get(header, 0) if rail_rows else 0) or 0)
    rail_rows.append(combined)

    row = _write_section_title(
        ws,
        "Наложение навигационных пломб по таможенным постам",
        1,
        span_cols=len(headers),
    )
    row = _write_section_title(
        ws,
        "Автомобильные пункты пропуска",
        row,
        span_cols=len(headers),
    )
    auto_header_row = row
    row = _write_rows_as_table(ws, headers, auto_rows, row, dash_zero=True)
    ws.cell(row=auto_header_row, column=1).value = None
    row += 1
    row = _write_section_title(
        ws,
        "Железнодорожные пункты пропуска",
        row,
        span_cols=len(headers),
    )
    rail_header_row = row
    _write_rows_as_table(ws, headers, rail_rows, row, dash_zero=True)
    ws.cell(row=rail_header_row, column=1).value = None

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    total_fill = PatternFill(fill_type="solid", fgColor="D9E2F3")
    total_rows = {
        auto_header_row + len(auto_rows),
        rail_header_row + len(rail_rows) - 1,
        rail_header_row + len(rail_rows),
    }
    for row_cells in ws.iter_rows(
        min_row=1,
        max_row=ws.max_row,
        min_col=1,
        max_col=len(headers),
    ):
        for cell in row_cells:
            cell.border = border
            cell.alignment = Alignment(
                horizontal=("left" if cell.column == 1 and cell.row > 3 else "center"),
                vertical="center",
                wrap_text=True,
            )
            if cell.row in total_rows or cell.column == len(headers):
                cell.font = _output_font(bold=True)
            if cell.row in total_rows:
                cell.fill = total_fill
    ws.freeze_panes = "A4"
    _set_col_widths(ws, headers)
    _set_column_bold(ws, len(headers))


def _write_table1_legacy(ws, t: Dict[str, Any]) -> None:
    """Единая Т1 метода 1 в утверждённом Word-формате."""
    headers = ["ТП/ЖД", "РФ", "РБ", "КР", "РА", "Итого"]
    rows = [dict(item) for item in (t.get("rows", []) or [])]
    row = _write_section_title(
        ws,
        "Транзитные перевозки (наложение НП)",
        1,
        span_cols=len(headers),
    )
    _write_rows_as_table(ws, headers, rows, start_row=row, dash_zero=True)

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_row = 2
    total_row = header_row + len(rows)
    total_fill = PatternFill(fill_type="solid", fgColor="D9E2F3")
    for row_cells in ws.iter_rows(
        min_row=1,
        max_row=total_row,
        min_col=1,
        max_col=len(headers),
    ):
        for cell in row_cells:
            cell.border = border
            cell.alignment = Alignment(
                horizontal=(
                    "left" if cell.column == 1 and cell.row > header_row else "center"
                ),
                vertical="center",
                wrap_text=True,
            )
            if cell.row in {1, header_row, total_row} or cell.column == 6:
                cell.font = _output_font(bold=True)
            if cell.row == total_row or cell.column == 6:
                cell.fill = total_fill

    # The daily RAW04 certificate follows the Word table exactly and has no
    # separate "registered rail" indicator at the right.  Older modes keep it.
    if not t.get("daily_raw04"):
        title_cell = ws.cell(row=1, column=8, value="Количество оформленных ЖД")
        value_cell = ws.cell(
            row=2,
            column=8,
            value=t.get("rail_registered_total", ""),
        )
        for cell in (title_cell, value_cell):
            cell.font = _output_font(bold=True)
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )
    ws.freeze_panes = "A3"
    _set_col_widths(ws, headers)
    if not t.get("daily_raw04"):
        ws.column_dimensions["H"].width = 24


def _write_tracking_table3(ws, t: Dict[str, Any]) -> None:
    """Новая Т3: проследование через РК по направлениям и транспорту."""
    headers = [
        "Тип транспорта",
        "РБ - КР",
        "РФ - КР",
        "КР - РФ",
        "КР - РБ",
        "Итого",
    ]
    rows = t.get("rows", []) or []
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    total_fill = PatternFill(fill_type="solid", fgColor="D9E2F3")

    ws.merge_cells("A1:F1")
    ws["A1"] = "Услуги прослеживания"
    ws.merge_cells("A2:F2")
    ws["A2"] = "Направления"

    for row_index in (1, 2):
        cell = ws.cell(row=row_index, column=1)
        cell.font = _output_font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for column, header in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=column, value=header)
        cell.font = _output_font(bold=True)
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    for row_index, item in enumerate(rows, start=4):
        is_total = str(item.get("Тип транспорта", "")).strip() == "Итого"
        for column, header in enumerate(headers, start=1):
            value = item.get(header)
            if header != "Тип транспорта" and isinstance(value, (int, float)) and value == 0:
                value = "-"
            cell = ws.cell(row=row_index, column=column, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_total or column == len(headers):
                cell.font = _output_font(bold=True)
            if is_total or column == len(headers):
                cell.fill = total_fill

    last_row = max(3, 3 + len(rows))
    for row_cells in ws.iter_rows(
        min_row=1, max_row=last_row, min_col=1, max_col=len(headers)
    ):
        for cell in row_cells:
            cell.border = border

    ws.column_dimensions["A"].width = 23
    for letter in ("B", "C", "D", "E", "F"):
        ws.column_dimensions[letter].width = 14
    ws.row_dimensions[3].height = 42
    ws.freeze_panes = "A4"


def _write_table3(ws, t: Dict[str, Any]) -> None:

    auto_out_rows = t.get("auto_out_rows", []) or []
    auto_in_rows = t.get("auto_in_rows", []) or []

    rail_out_rows = t.get("rail_out_rows", []) or []
    rail_in_rows = t.get("rail_in_rows", []) or []

    headers = ["Направление", "Количество"]

    row = 1

    row = _write_section_title(
        ws,
        "Статистические данные завершенных перевозок",
        row,
        span_cols=len(headers),
    )

    row = _write_section_title(
        ws,
        "Завершённые автомобильные перевозки из Республики Казахстан",
        row,
        span_cols=len(headers),
    )

    row = _write_rows_as_table(
        ws,
        headers,
        auto_out_rows,
        row,
    )

    row += 1

    row = _write_section_title(
        ws,
        "Завершенные автомобильные перевозки в Республику Казахстан",
        row,
        span_cols=len(headers),
    )

    row = _write_rows_as_table(
        ws,
        headers,
        auto_in_rows,
        row,
    )

    row += 1

    row = _write_section_title(
        ws,
        "Завершённые железнодорожные перевозки из Республики Казахстан",
        row,
        span_cols=len(headers),
    )

    row = _write_rows_as_table(
        ws,
        headers,
        rail_out_rows,
        row,
    )

    row += 1

    row = _write_section_title(
        ws,
        "Завершенные железнодорожные перевозки в Республику Казахстан",
        row,
        span_cols=len(headers),
    )

    row = _write_rows_as_table(
        ws,
        headers,
        rail_in_rows,
        row,
    )

    ws.freeze_panes = "A4"

    _set_col_widths(ws, headers)


def _write_table4(ws, t: Dict[str, Any]) -> None:

    rows = t.get("rows", []) or []
    date_label = str(
        t.get("date_label_short") or t.get("date_label", "") or ""
    ).strip()

    headers = [
        "Пункты\nпропуска/региональ\nные склады",
        f"НП на {date_label}" if date_label else "НП",
    ]
    data_rows = [
        {
            headers[0]: row.get("Точка"),
            headers[1]: row.get("Количество НП"),
        }
        for row in rows
    ]
    first_np_row = 2
    last_np_row = first_np_row + len(data_rows) - 1
    data_rows.append(
        {
            headers[0]: "Итого(с учетом пилот Карасу)",
            headers[1]: f"=SUM(B{first_np_row}:B{last_np_row})+{PILOT_KARASU_NP}",
        }
    )

    row = _write_rows_as_table(
        ws,
        headers=headers,
        rows=data_rows,
        start_row=1,
        dash_zero=True,
    )

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row_cells in ws.iter_rows(min_row=1, max_row=row - 1, min_col=1, max_col=2):
        for cell in row_cells:
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center" if cell.column == 2 else "left",
                vertical="center",
                wrap_text=True,
            )

            if cell.row == 1 or cell.column == 2:
                cell.font = _output_font(bold=True)

    used_np_rows = t.get("used_np_rows", []) or []
    monthly_used_np_rows = t.get("monthly_used_np_rows", []) or used_np_rows
    used_headers = [
        "Пункты пропуска/региональные склады",
        "",
        "Кол-во\nиспольз\nованных НП",
    ]

    for col, value in zip(range(4, 7), used_headers):
        cell = ws.cell(row=1, column=col, value=value)
        cell.border = border
        cell.font = _output_font(bold=True)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    used_row = 2
    for item in used_np_rows:
        point_name = item.get("Пункт")
        is_total = point_name in {"Навешивания на территории РК", "Итого"}
        used_value = item.get("Количество")
        if point_name != "Итого" and isinstance(used_value, (int, float)) and used_value == 0:
            used_value = "-"

        ws.cell(row=used_row, column=4, value=point_name)
        ws.cell(row=used_row, column=5, value=None)
        ws.cell(row=used_row, column=6, value=used_value)

        for col in range(4, 7):
            cell = ws.cell(row=used_row, column=col)
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center" if col == 6 else "left",
                vertical="center",
                wrap_text=True,
            )

            if col == 6 or is_total:
                cell.font = _output_font(bold=True)

        used_row += 1

    # Version 2.3: average daily seal consumption.  The user-facing day
    # The counter restarts on the first day of every calendar month.
    # Examples: 31.08 -> 31, 01.09 -> 1, 02.09 -> 2.
    # Keep column G empty as a visual gutter between the second and third tables.
    daily_start_col = 8  # H
    daily_title_row = 1
    daily_days_row = 2
    daily_header_row = 3
    daily_first_row = 4
    elapsed_days = int(t.get("monthly_elapsed_days") or _stage2_elapsed_days())
    daily_value_col = daily_start_col + 2
    daily_value_col_letter = get_column_letter(daily_value_col)

    ws.merge_cells(
        start_row=daily_title_row,
        start_column=daily_start_col,
        end_row=daily_title_row,
        end_column=daily_start_col + 2,
    )
    ws.cell(
        row=daily_title_row,
        column=daily_start_col,
        value="Ежедневный расход пломб",
    )

    ws.merge_cells(
        start_row=daily_days_row,
        start_column=daily_start_col,
        end_row=daily_days_row,
        end_column=daily_start_col + 1,
    )
    ws.cell(
        row=daily_days_row,
        column=daily_start_col,
        value="Количество прошедших дней текущего месяца",
    )
    ws.cell(
        row=daily_days_row,
        column=daily_value_col,
        value=elapsed_days,
    )

    daily_headers = [
        "Пункты пропуска/региональные склады",
        "Использовано НП\nв текущем месяце",
        "Средний расход\nпломб в день",
    ]
    for col, value in zip(range(daily_start_col, daily_start_col + 3), daily_headers):
        ws.cell(row=daily_header_row, column=col, value=value)

    for index, item in enumerate(monthly_used_np_rows):
        daily_row = daily_first_row + index
        point_name = item.get("Пункт")
        is_total = point_name in {"Навешивания на территории РК", "Итого"}
        monthly_value = item.get("Количество")
        if point_name != "Итого" and isinstance(monthly_value, (int, float)) and monthly_value == 0:
            monthly_value = "-"

        ws.cell(row=daily_row, column=daily_start_col, value=point_name)
        ws.cell(
            row=daily_row,
            column=daily_start_col + 1,
            value=monthly_value,
        )
        ws.cell(
            row=daily_row,
            column=daily_value_col,
            value=(
                f'=IF(ISNUMBER(I{daily_row}),ROUND(I{daily_row}/${daily_value_col_letter}${daily_days_row},0),"-")'
            ),
        )

        if is_total:
            for col in range(daily_start_col, daily_start_col + 3):
                ws.cell(row=daily_row, column=col).font = _output_font(bold=True)

    daily_last_row = daily_first_row + len(monthly_used_np_rows) - 1
    for row_idx in range(daily_title_row, daily_last_row + 1):
        for col in range(daily_start_col, daily_start_col + 3):
            cell = ws.cell(row=row_idx, column=col)
            cell.border = border
            cell.alignment = Alignment(
                horizontal=(
                    "left"
                    if col == daily_start_col and row_idx >= daily_first_row
                    else "center"
                ),
                vertical="center",
                wrap_text=True,
            )

            if row_idx in {daily_title_row, daily_days_row, daily_header_row} or col == daily_value_col:
                cell.font = _output_font(bold=True)

    ws.row_dimensions[1].height = 42
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["D"].width = 38
    ws.column_dimensions["E"].width = 8
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 38
    ws.column_dimensions["I"].width = 8
    ws.column_dimensions["J"].width = 22
    ws.row_dimensions[daily_title_row].height = max(
        ws.row_dimensions[daily_title_row].height or 0,
        28,
    )
    ws.row_dimensions[daily_header_row].height = max(
        ws.row_dimensions[daily_header_row].height or 0,
        42,
    )

    ws.freeze_panes = "A2"



def _write_table5(ws, t: Dict[str, Any]) -> None:

    rows = t.get("rows", []) or []

    headers = [
        "Страна начала перевозки",
        "Количество перевозок",
    ]

    row = 1

    row = _write_section_title(
        ws,
        "Статистика по перевозкам, направленным в сторону Республики Казахстан",
        row,
        span_cols=2,
    )

    _write_rows_as_table(
        ws,
        headers=headers,
        rows=rows,
        start_row=row,
    )

    ws.freeze_panes = "A3"

    _set_col_widths(ws, headers)


def _write_daily_np_direction_table6(ws, t: Dict[str, Any]) -> None:
    """RAW-04 Table 6: seal applications, not transportation counts."""
    headers = [
        "Направление", "Транзит", "Экспорт", "Взаимная торговля", "Итого",
    ]
    rows = t.get("rows", []) or []
    row = _write_section_title(
        ws,
        "Наложение навигационных пломб по направлениям",
        1,
        span_cols=len(headers),
    )
    _write_rows_as_table(ws, headers, rows, row, dash_zero=True)
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    total_fill = PatternFill(fill_type="solid", fgColor="D9E2F3")
    total_row = row + len(rows)
    for cells in ws.iter_rows(min_row=1, max_row=total_row, min_col=1, max_col=len(headers)):
        for cell in cells:
            cell.border = border
            cell.alignment = Alignment(
                horizontal="left" if cell.column == 1 and cell.row > row else "center",
                vertical="center",
                wrap_text=True,
            )
            if cell.row in {1, row, total_row} or cell.column == len(headers):
                cell.font = _output_font(bold=True)
            if cell.row == total_row:
                cell.fill = total_fill
    ws.freeze_panes = "A3"
    _set_col_widths(ws, headers)


def _write_daily_diagnostics(ws, t: Dict[str, Any]) -> None:
    headers = ["Показатель", "Значение"]
    rows = [
        {"Показатель": "Период с", "Значение": t.get("period_start", "-")},
        {"Показатель": "Последнее время статуса", "Значение": t.get("last_status", "-")},
        *(t.get("rows", []) or []),
    ]
    _write_rows_as_table(ws, headers, rows, 1, dash_zero=False)
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cells in ws.iter_rows(min_row=1, max_row=1 + len(rows), min_col=1, max_col=2):
        for cell in cells:
            cell.border = border
            cell.alignment = Alignment(horizontal="left" if cell.column == 1 else "center", vertical="center", wrap_text=True)
            if cell.row == 1:
                cell.font = _output_font(bold=True)
    ws.column_dimensions["A"].width = 64
    ws.column_dimensions["B"].width = 24
    ws.freeze_panes = "A2"


def _write_daily_raw04_header(ws, t: Dict[str, Any]) -> None:
    """Write the numeric certificate header preceding the report tables."""
    headers = ["№", "Название", "Данные", "Число"]
    rows = t.get("rows", []) or []
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="D9E2F3")

    for column, title in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=column, value=title)
        cell.font = _output_font(bold=True)
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    for row_index, item in enumerate(rows, start=2):
        values = [item.get(header, "") for header in headers]
        for column, value in enumerate(values, start=1):
            cell = ws.cell(row=row_index, column=column, value=value)
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center" if column in {1, 4} else "left",
                vertical="center",
                wrap_text=True,
            )
            if column == 4:
                cell.font = _output_font(bold=True)
                cell.number_format = "# ##0"
            else:
                cell.font = _output_font()

    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 48
    ws.column_dimensions["C"].width = 36
    ws.column_dimensions["D"].width = 18
    ws.freeze_panes = "A2"


def _write_stock_unavailable(ws, t: Dict[str, Any]) -> None:
    """Keep Table8 visible without inventing stock values when report is absent."""
    message = str(t.get("stock_unavailable_message") or "Данные недоступны.")
    ws.merge_cells("A1:F1")
    ws["A1"] = "Остатки и расход навигационных пломб"
    ws["A1"].font = _output_font(bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("A2:F2")
    ws["A2"] = message
    ws["A2"].font = _output_font(bold=True)
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=6):
        for cell in row:
            cell.border = Border(
                left=Side(style="thin", color="000000"),
                right=Side(style="thin", color="000000"),
                top=Side(style="thin", color="000000"),
                bottom=Side(style="thin", color="000000"),
            )
    ws.column_dimensions["A"].width = 32
    for column in "BCDEF":
        ws.column_dimensions[column].width = 18


def write_output(out_path: Path, tables: List[Any]) -> None:

    out_path.parent.mkdir(parents=True, exist_ok=True)

    if not tables:

        write_output_test(
            out_path,
            sheet_name="Таблица1",
        )

        print(
            f"[io_excel.write_output] "
            f"tables_stub_count=0 "
            f"(created {out_path} with Таблица1:A1='OK')"
        )

        return

    wb = Workbook()

    _delete_default_sheet_if_needed(wb)

    for t in tables:

        if not isinstance(t, dict):
            continue

        sheet_title = _safe_sheet_title(
            str(t.get("table", "Sheet"))
        )

        ws = wb.create_sheet(title=sheet_title)

        if t.get("raw04_mode"):
            print(f"[raw04] записывается лист: {sheet_title}")

        # The RAW04 combined certificate keeps the legacy stock sheet after
        # the seven certificate tables, while preserving its old layout.
        if t.get("legacy_stock_layout"):
            if t.get("stock_unavailable"):
                _write_stock_unavailable(ws, t)
                continue
            _write_table4(ws, t)
            _apply_output_font_to_used_range(ws, force_bold=True)
            _fit_used_range(ws)
            continue

        if sheet_title == "Шапка" and t.get("daily_raw04_header"):
            _write_daily_raw04_header(ws, t)
            continue

        if sheet_title == "Таблица1" and t.get("transport_split"):
            _write_table1_split(ws, t)
            _apply_output_font_to_used_range(ws)
            _bold_total_rows(ws)
            _fit_used_range(ws)
            continue

        if sheet_title == "Таблица1" and t.get("legacy_layout"):
            _write_table1_legacy(ws, t)
            _apply_output_font_to_used_range(ws)
            _bold_total_rows(ws)
            _fit_used_range(ws)
            continue

        if sheet_title == "Таблица2":

            _write_table2(ws, t)
            _apply_output_font_to_used_range(ws)
            _bold_total_rows(ws)
            _fit_used_range(ws)

            continue

        if sheet_title == "Таблица3":

            _write_tracking_table3(ws, t)
            _apply_output_font_to_used_range(ws)
            _bold_total_rows(ws)
            _fit_used_range(ws)

            continue

        if sheet_title == "Таблица4":

            _write_table3(ws, t)
            _apply_output_font_to_used_range(ws)
            _bold_total_rows(ws)
            _fit_used_range(ws)

            continue

        if sheet_title == "Таблица5":

            if t.get("daily_raw04"):
                _write_table5(ws, t)
                _apply_output_font_to_used_range(ws)
                _bold_total_rows(ws)
                _fit_used_range(ws)
                continue

            _write_table4(ws, t)
            _apply_output_font_to_used_range(ws, force_bold=True)
            _fit_used_range(ws)

            continue

        if sheet_title == "Таблица6":

            if t.get("daily_raw04"):
                _write_daily_np_direction_table6(ws, t)
                _apply_output_font_to_used_range(ws)
                _bold_total_rows(ws)
                _fit_used_range(ws)
                continue

            _write_table5(ws, t)
            _apply_output_font_to_used_range(ws)
            _bold_total_rows(ws)
            _fit_used_range(ws)

            continue

        if sheet_title == "Таблица7":
            _write_seal_post_table7(ws, t)
            _apply_output_font_to_used_range(ws)
            _bold_total_rows(ws)
            _fit_used_range(ws)
            continue

        if sheet_title == "Диагностика" and t.get("daily_raw04"):
            _write_daily_diagnostics(ws, t)
            _apply_output_font_to_used_range(ws)
            _fit_used_range(ws)
            continue

        rows = t.get("rows", [])

        if not isinstance(rows, list) or (
            rows and not isinstance(rows[0], dict)
        ):

            ws["A1"] = str(t)

            continue

        if sheet_title == "Таблица1":

            headers = [
                "ТП/ЖД",
                "РФ",
                "РБ",
                "КР",
                "РА",
                "Итого",
            ]

        else:

            headers = list(rows[0].keys()) if rows else ["data"]

        _write_rows_as_table(
            ws,
            headers,
            rows,
            dash_zero=(sheet_title == "Таблица1"),
        )

        if sheet_title == "Таблица1":
            rail_registered_total = t.get("rail_registered_total")

            title_cell = ws.cell(row=1, column=8, value="Количество оформленных ЖД")
            value_cell = ws.cell(
                row=2,
                column=8,
                value="" if rail_registered_total is None else rail_registered_total,
            )

            for cell in (title_cell, value_cell):
                cell.font = _output_font(bold=True)
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )

            ws.column_dimensions["H"].width = 24
            _set_column_bold(ws, 6)

        ws.freeze_panes = "A2"

        _set_col_widths(ws, headers)
        _apply_output_font_to_used_range(ws)
        _bold_total_rows(ws)
        _fit_used_range(ws)

    wb.save(out_path)

    print(
        f"[io_excel.write_output] "
        f"wrote {len(tables)} table(s) to: {out_path}"
    )
