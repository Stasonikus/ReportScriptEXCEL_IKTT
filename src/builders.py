# src/builders.py
from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path
from typing import Any, Dict


# Для отладки Т1 можно поставить, например: "Атамекен"
DEBUG_PP: str | None = None


def build_table1(report_1_norm: Any, cfg: Dict[str, Any], tp_zhd_list_t1: list[str]) -> Any:
    """
    Т1 — реализовано ранее (v2):
    - секция по t1_block_range (из normalize.py)
    - блок ПП: после строки ПП разрешены только строки, начинающиеся с "→"
      как только строка НЕ начинается с "→" — блок ПП завершается
    - повтор ПП: суммируем по направлениям
    - "→ 0" исключаем
    """
    ws = report_1_norm["ws"]

    start_row, end_row = report_1_norm.get("t1_block_range", (1, ws.max_row + 1))
    scan_start = start_row + 1
    scan_end = end_row

    markers = [m.lower() for m in cfg.get("pp_markers", ["т/п", "тп", "таможенный пост"])]
    quotes = cfg.get("quotes_to_strip", ['"', "«", "»"])
    alias_map: Dict[str, str] = cfg.get("pp_aliases", {}) or {}

    countries = cfg.get("countries", {})
    valid_arrows = {"RU", "BY", "KG", "ARM"}
    target_cols = {"РФ", "РБ", "КР", "РА"}

    allowed = set(tp_zhd_list_t1)

    max_scan_cols = min(ws.max_column, 50)

    def clean_quotes(s: str) -> str:
        for q in quotes:
            s = s.replace(q, "")
        return s.strip()

    def norm_spaces_dashes(s: str) -> str:
        s = s.replace("–", "-").replace("—", "-")
        while "  " in s:
            s = s.replace("  ", " ")
        return s.strip()

    def title_simple(s: str) -> str:
        s = s.strip()
        if not s:
            return s
        return s[:1].upper() + s[1:].lower()

    def parse_int(v: Any) -> int:
        if isinstance(v, (int, float)):
            return int(v)
        if isinstance(v, str):
            s = v.strip().replace(" ", "")
            if s.isdigit():
                return int(s)
        return 0

    def extract_pp_name(text: str) -> str | None:
        low = text.lower()
        hit_pos = -1
        hit_len = 0
        for m in markers:
            pos = low.find(m)
            if pos != -1:
                hit_pos = pos
                hit_len = len(m)
                break
        if hit_pos == -1:
            return None

        after = text[hit_pos + hit_len :].strip()
        if not after:
            return None

        name = clean_quotes(after)
        name = norm_spaces_dashes(name)
        name = title_simple(name)
        name = alias_map.get(name, name)
        return name

    def extract_arrow_code(text: str) -> str | None:
        m = re.match(r"^\s*→\s*([A-Za-z0-9]{1,4})\s*$", text)
        if not m:
            return None
        return m.group(1).upper()

    def first_nonempty_cell(row_idx: int) -> tuple[int, Any] | None:
        for col in range(1, max_scan_cols + 1):
            v = ws.cell(row=row_idx, column=col).value
            if v is None:
                continue
            if isinstance(v, str) and not v.strip():
                continue
            return col, v
        return None

    def dump_row(row_idx: int, cols: int = 14) -> None:
        if not DEBUG_PP:
            return
        vals = []
        for c in range(1, cols + 1):
            v = ws.cell(row=row_idx, column=c).value
            vals.append("" if v is None else str(v))
        letters = "ABCDEFGHIJKLMN"[:cols]
        print(f"    row={row_idx}: " + " | ".join([f"{letters[i]}={vals[i]}" for i in range(cols)]))

    pp_vals: Dict[str, Dict[str, int]] = {}
    current_pp: str | None = None

    for row_idx in range(scan_start, scan_end):
        first = first_nonempty_cell(row_idx)
        if first is None:
            current_pp = None
            continue

        col0, val0 = first

        # строка ПП
        if isinstance(val0, str):
            pp = extract_pp_name(val0)
        else:
            pp = None

        if pp is not None:
            if pp in allowed:
                current_pp = pp
                pp_vals.setdefault(current_pp, {"РФ": 0, "РБ": 0, "КР": 0, "РА": 0})
                if DEBUG_PP and current_pp == DEBUG_PP:
                    print(f"[DEBUG] PP start '{current_pp}' at row={row_idx}, first_col={col0}")
                    dump_row(row_idx)
            else:
                current_pp = None
            continue

        if current_pp is None:
            continue

        if not isinstance(val0, str):
            current_pp = None
            continue

        arrow_code = extract_arrow_code(val0)
        if arrow_code is None:
            current_pp = None
            continue

        value_col = col0 + 1
        value = parse_int(ws.cell(row=row_idx, column=value_col).value)

        if DEBUG_PP and current_pp == DEBUG_PP:
            print(f"[DEBUG] arrow row={row_idx} arrow='{arrow_code}' value={value}")
            dump_row(row_idx)

        if arrow_code == "0":
            continue

        if arrow_code not in valid_arrows:
            continue

        mapped = cfg.get("countries", {}).get(arrow_code, arrow_code)
        if mapped not in target_cols:
            continue

        pp_vals[current_pp][mapped] += value

    rows: list[dict[str, Any]] = []
    for pp in tp_zhd_list_t1:
        vals = pp_vals.get(pp, {"РФ": 0, "РБ": 0, "КР": 0, "РА": 0})
        total = int(vals["РФ"]) + int(vals["РБ"]) + int(vals["КР"]) + int(vals["РА"])
        rows.append({"ТП/ЖД": pp, "РФ": int(vals["РФ"]), "РБ": int(vals["РБ"]), "КР": int(vals["КР"]), "РА": int(vals["РА"]), "Итого": total})

    print(f"[builders.build_table1] built rows={len(rows)}")
    return {"table": "Таблица1", "stage": "data-only", "rows": rows}


def build_table2(report_1_norm: Any, cfg: Dict[str, Any]) -> Any:
    """
    Т2: "Статистические данные экспорта и взаимной торговли из РК"

    Блок источника:
      начало: строка содержит "Перевозки по типам кортежей"
      конец:   строка содержит "Аналитика по пунктам пропуска"

    Правила:
      - строка header: первая ячейка = "From" -> пропускаем
      - KZ->KZ: не обрабатывается и не суммируется
      - KZ->ВСЕ: не выводится и не суммируется, но используется для проверок
      - KZ->?  : не выводится и не суммируется, но используется в проверке взаимной торговли
      - направления: RU, BY, KG, RA (если RA нет -> 0)
      - Экспорт: берём из колонки "Таможенная процедура экспорта"
        проверка: sum(RU,BY,KG,RA) == (KZ->ВСЕ).export
      - Взаимная торговля: берём из колонки "Перевозка товаров в рамках взаимной торговли"
        проверка (твоя финальная формула):
           sum(RU,BY,KG,RA) == (KZ->ВСЕ) - (KZ->KZ) - (KZ->?)
    """
    ws = report_1_norm["ws"]

    start_text = "Перевозки по типам кортежей"
    end_text = "Аналитика по пунктам пропуска"

    # Поиск границ блока (по любой ячейке строки)
    start_row = None
    end_row = None
    for r in range(1, ws.max_row + 1):
        row_has = False
        for c in range(1, min(ws.max_column, 12) + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and start_text.lower() in v.lower():
                start_row = r
                row_has = True
                break
        if row_has:
            break

    if start_row is None:
        print("[builders.build_table2] FAIL: start marker not found")
        return {"table": "Таблица2", "stage": "data-only", "export_rows": [], "trade_rows": []}

    for r in range(start_row + 1, ws.max_row + 1):
        for c in range(1, min(ws.max_column, 12) + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and end_text.lower() in v.lower():
                end_row = r
                break
        if end_row is not None:
            break

    if end_row is None:
        print("[builders.build_table2] FAIL: end marker not found")
        return {"table": "Таблица2", "stage": "data-only", "export_rows": [], "trade_rows": []}

    # Найдём строку заголовков (где A == "From") и индексы нужных колонок
    header_row = None
    for r in range(start_row + 1, end_row):
        a = ws.cell(row=r, column=1).value
        if isinstance(a, str) and a.strip() == "From":
            header_row = r
            break

    if header_row is None:
        print("[builders.build_table2] FAIL: header row 'From' not found")
        return {"table": "Таблица2", "stage": "data-only", "export_rows": [], "trade_rows": []}

    # Сопоставим имена колонок -> индекс (по заголовкам)
    header_map: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, str) and v.strip():
            header_map[v.strip().lower()] = c

    def find_col_by_contains(substr: str) -> int | None:
        s = substr.lower()
        for k, c in header_map.items():
            if s in k:
                return c
        return None

    col_from = find_col_by_contains("from")
    col_to = find_col_by_contains("to")
    col_trade = find_col_by_contains("взаимной торговли")
    col_export = find_col_by_contains("процедура экспорта")

    if not all([col_from, col_to, col_trade, col_export]):
        print("[builders.build_table2] FAIL: required columns not found")
        print(f"  col_from={col_from}, col_to={col_to}, col_trade={col_trade}, col_export={col_export}")
        return {"table": "Таблица2", "stage": "data-only", "export_rows": [], "trade_rows": []}

    def parse_int(v: Any) -> int:
        if isinstance(v, (int, float)):
            return int(v)
        if isinstance(v, str):
            s = v.strip().replace(" ", "")
            if s.isdigit():
                return int(s)
        return 0

    # Нормализация кода страны -> полное имя
    full_names = cfg.get(
        "country_full_names",
        {
            "RU": "Российская Федерация",
            "BY": "Республика Беларусь",
            "KG": "Кыргызская Республика",
            "RA": "Республика Армения",
        },
    )

    dirs = ["RU", "BY", "KG", "RA"]

    export_sum: Dict[str, int] = {k: 0 for k in dirs}
    trade_sum: Dict[str, int] = {k: 0 for k in dirs}

    kz_all_export = 0
    kz_all_trade = 0
    kz_kz_trade = 0
    kz_q_trade = 0  # KZ->?

    # Парс строк данных: после header_row до end_row
    for r in range(header_row + 1, end_row):
        frm = ws.cell(row=r, column=col_from).value
        to = ws.cell(row=r, column=col_to).value

        frm_s = frm.strip().upper() if isinstance(frm, str) else ""
        to_s = to.strip().upper() if isinstance(to, str) else ""

        # пропуски
        if frm_s == "" and to_s == "":
            continue

        # только строки From=KZ нас интересуют
        if frm_s != "KZ":
            continue

        exp_val = parse_int(ws.cell(row=r, column=col_export).value)
        trd_val = parse_int(ws.cell(row=r, column=col_trade).value)

        # исключения
        if to_s == "KZ":
            # KZ->KZ не обрабатываем, но для проверки взаимной торговли запоминаем
            kz_kz_trade = trd_val
            continue

        if to_s == "ВСЕ":
            # KZ->ВСЕ не суммируем, но используем для проверок
            kz_all_export = exp_val
            kz_all_trade = trd_val
            continue

        if to_s == "?":
            # не суммируем, но используем в проверке взаимной торговли
            kz_q_trade = trd_val
            continue

        # направления
        if to_s in dirs:
            export_sum[to_s] += exp_val
            trade_sum[to_s] += trd_val
            continue

        # остальные направления игнорируем
        continue

    # RA может отсутствовать — по умолчанию 0 (уже так в dict)

    export_total = sum(export_sum.values())
    trade_total = sum(trade_sum.values())

    # проверки
    export_check_ok = (export_total == kz_all_export)

    trade_expected = kz_all_trade - kz_kz_trade - kz_q_trade
    trade_check_ok = (trade_total == trade_expected)

    # формируем строки вывода (данные)
    export_rows: list[dict[str, Any]] = []
    trade_rows: list[dict[str, Any]] = []

    for i, code in enumerate(["RU", "BY", "KG", "RA"], start=1):
        export_rows.append({"№": i, "Страна": full_names[code], "Количество перевозок": int(export_sum[code])})
        trade_rows.append({"№": i, "Страна": full_names[code], "Количество перевозок": int(trade_sum[code])})

    export_rows.append({"№": 5, "Страна": "Итого", "Количество перевозок": int(export_total)})
    trade_rows.append({"№": 5, "Страна": "Итого", "Количество перевозок": int(trade_total)})

    # вывод в консоль
    print("[builders.build_table2] start")
    print(f"[T2] block rows: {start_row}..{end_row} (header_row={header_row})")

    print("\n[T2] Экспорт:")
    for r in export_rows:
        print(" ", r)

    print(f"[T2 CHECK export] sum(RU,BY,KG,RA)={export_total}  vs  KZ->ВСЕ(export)={kz_all_export}  => {'OK' if export_check_ok else 'FAIL'}")

    print("\n[T2] Взаимная торговля:")
    for r in trade_rows:
        print(" ", r)

    print(
        f"[T2 CHECK trade] sum(RU,BY,KG,RA)={trade_total}  vs  (KZ->ВСЕ) - (KZ->KZ) - (KZ->?)"
        f" = {kz_all_trade} - {kz_kz_trade} - {kz_q_trade} = {trade_expected}"
        f"  => {'OK' if trade_check_ok else 'FAIL'}"
    )

    print("[builders.build_table2] done")

    return {
        "table": "Таблица2",
        "stage": "data-only",
        "export_rows": export_rows,
        "trade_rows": trade_rows,
        "checks": {
            "export_total": export_total,
            "export_kz_all": kz_all_export,
            "export_ok": export_check_ok,
            "trade_total": trade_total,
            "trade_expected": trade_expected,
            "trade_ok": trade_check_ok,
        },
    }


def build_table3(
    report_1_norm: Any,
    cfg: Dict[str, Any],
    rail_raw: Any | None = None,
    data_dir: Path | str | None = None,
) -> Any:
    """
    Т3: Статистические данные завершенных перевозок
    Новая форма: 4 отдельных блока, без матрицы.

    Источник: report_1, лист "Отчёт"
    Блок:
      - start: строка содержит "Показатель", а следующая ячейка содержит "Значение"
      - end: строка содержит "Перевозки по типам кортежей"

    Авто:
      value = Деактивированная перевозка + Завершенная перевозка
      учитываем только пары:
        KZ -> RU/BY/KG/ARM
        RU/BY/KG/ARM -> KZ

    Ж/Д:
      по текущему ТЗ источник отдельно не определён,
      поэтому пока формируем нули в новой форме.
    """
    ws = report_1_norm["ws"]

    start_row = None
    end_row = None
    max_scan_cols = min(ws.max_column, 30)

    def row_cells_str(r: int) -> list[str]:
        vals: list[str] = []
        for c in range(1, max_scan_cols + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip():
                vals.append(v.strip())
        return vals

    # ---- 1) найти границы блока
    for r in range(1, ws.max_row + 1):
        for c in range(1, max_scan_cols):
            a = ws.cell(row=r, column=c).value
            b = ws.cell(row=r, column=c + 1).value
            if isinstance(a, str) and isinstance(b, str):
                if a.strip().lower() == "показатель" and b.strip().lower() == "значение":
                    start_row = r
                    break
        if start_row is not None:
            break

    if start_row is None:
        print("[builders.build_table3] FAIL: start marker not found ('Показатель'/'Значение')")
        return {
            "table": "Таблица3",
            "stage": "data-only",
            "auto_out_rows": [],
            "auto_in_rows": [],
            "rail_out_rows": [],
            "rail_in_rows": [],
        }

    for r in range(start_row + 1, ws.max_row + 1):
        vals = row_cells_str(r)
        if any("перевозки по типам кортежей" in v.lower() for v in vals):
            end_row = r
            break

    if end_row is None:
        print("[builders.build_table3] FAIL: end marker not found ('Перевозки по типам кортежей')")
        return {
            "table": "Таблица3",
            "stage": "data-only",
            "auto_out_rows": [],
            "auto_in_rows": [],
            "rail_out_rows": [],
            "rail_in_rows": [],
        }

    # ---- 2) найти header row и нужные колонки
    header_row = None
    for r in range(start_row + 1, end_row):
        a = ws.cell(row=r, column=1).value
        if isinstance(a, str) and a.strip() == "From":
            header_row = r
            break

    if header_row is None:
        print("[builders.build_table3] FAIL: header row 'From' not found")
        return {
            "table": "Таблица3",
            "stage": "data-only",
            "auto_out_rows": [],
            "auto_in_rows": [],
            "rail_out_rows": [],
            "rail_in_rows": [],
        }

    header_map: Dict[str, int] = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, str) and v.strip():
            header_map[v.strip().lower()] = c

    def find_col_contains(substr: str) -> int | None:
        s = substr.lower()
        for k, col in header_map.items():
            if s in k:
                return col
        return None

    col_from = find_col_contains("from")
    col_to = find_col_contains("to")
    col_deact = find_col_contains("деактив")
    col_done = find_col_contains("заверш")

    if not all([col_from, col_to, col_deact, col_done]):
        print("[builders.build_table3] FAIL: required columns not found")
        print(f"  col_from={col_from}, col_to={col_to}, col_deact={col_deact}, col_done={col_done}")
        return {
            "table": "Таблица3",
            "stage": "data-only",
            "auto_out_rows": [],
            "auto_in_rows": [],
            "rail_out_rows": [],
            "rail_in_rows": [],
        }

    def parse_int(v: Any) -> int:
        if isinstance(v, (int, float)):
            return int(v)
        if isinstance(v, str):
            s = v.strip().replace(" ", "")
            if s.isdigit():
                return int(s)
        return 0

    dirs = ["RU", "BY", "KG", "ARM"]

    # Собираем потоки только KZ <-> country
    flows: Dict[tuple[str, str], int] = {}

    for r in range(header_row + 1, end_row):
        frm = ws.cell(row=r, column=col_from).value
        to = ws.cell(row=r, column=col_to).value

        frm_s = frm.strip().upper() if isinstance(frm, str) else ""
        to_s = to.strip().upper() if isinstance(to, str) else ""

        if frm_s == "" and to_s == "":
            continue

        if frm_s == "FROM":
            continue

        if to_s == "?":
            continue

        if frm_s == "KZ" and to_s in {"KZ", "ВСЕ"}:
            continue

        if not ((frm_s == "KZ" and to_s in dirs) or (to_s == "KZ" and frm_s in dirs)):
            continue

        deact = parse_int(ws.cell(row=r, column=col_deact).value)
        done = parse_int(ws.cell(row=r, column=col_done).value)
        val = deact + done

        key = (frm_s, to_s)
        flows[key] = flows.get(key, 0) + val

    def norm_text(value: Any) -> str:
        if value is None:
            return ""

        s = str(value).strip().lower()
        s = s.replace("ё", "е")
        s = s.replace("–", "-").replace("—", "-")
        s = re.sub(r"[\"'«»()]", " ", s)
        s = re.sub(r"\s+", " ", s)
        return s.strip()

    def parse_company_id(value: Any) -> str:
        if value is None:
            return ""

        if isinstance(value, float) and value.is_integer():
            return str(int(value))

        return str(value).strip()

    def load_locations(country_code: str) -> set[str]:
        if data_dir is None:
            base_dir = Path(__file__).resolve().parent / "data"
        else:
            base_dir = Path(data_dir)

        path = base_dir / f"locations_{country_code}.json"
        if not path.exists():
            print(f"[builders.build_table3] WARN: locations file not found: {path}")
            return set()

        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except Exception as exc:
            print(f"[builders.build_table3] WARN: cannot read {path}: {exc}")
            return set()

        return {
            norm_text(item)
            for item in data.get("locations", [])
            if norm_text(item)
        }

    location_sets = {
        "RU": load_locations("RF"),
        "BY": load_locations("RB"),
        "KG": load_locations("KG"),
        "KZ": load_locations("KZ"),
    }

    def location_country(value: Any, allowed_codes: set[str]) -> str | None:
        key = norm_text(value)
        if not key:
            return None

        for code in ["RU", "BY", "KG", "KZ"]:
            if code not in allowed_codes:
                continue

            locations = location_sets.get(code, set())
            if key in locations:
                return code

        # Fallback for obvious names when the JSON dictionary is incomplete.
        if "RU" in allowed_codes and any(x in key for x in ["росси", "себеж", "рф"]):
            return "RU"
        if "BY" in allowed_codes and any(x in key for x in ["беларус", "брест", "минск", "гомел", "рб"]):
            return "BY"
        if "KG" in allowed_codes and not location_sets.get("KG") and any(x in key for x in ["кыргыз", "киргиз", "кг", "аламедин", "бишкек"]):
            return "KG"
        if "KZ" in allowed_codes and any(x in key for x in ["казахстан", "т/п", "алматы", "темиржол", "алтынколь", "болашак"]):
            return "KZ"

        return None

    def find_rail_header(ws_rail) -> int | None:
        for row_idx in range(1, ws_rail.max_row + 1):
            values = [ws_rail.cell(row=row_idx, column=c).value for c in range(1, ws_rail.max_column + 1)]
            if "Владелец записи" in values and "Статус" in values and "company_id" in values:
                return row_idx

        return None

    def calc_rail_flows() -> tuple[Dict[str, int], Dict[str, int]]:
        rail_out = {"RU": 0, "BY": 0, "KG": 0, "ARM": 0}
        rail_in = {"RU": 0, "BY": 0, "KG": 0, "ARM": 0}

        if rail_raw is None:
            print("[builders.build_table3] WARN: ЖД Выгрузка.xlsx not found; rail data = 0")
            return rail_out, rail_in

        ws_rail = rail_raw.get("ws") if isinstance(rail_raw, dict) else None
        if ws_rail is None:
            print("[builders.build_table3] WARN: rail workbook is not readable; rail data = 0")
            return rail_out, rail_in

        rail_header = find_rail_header(ws_rail)
        if rail_header is None:
            print("[builders.build_table3] WARN: rail header row not found; rail data = 0")
            return rail_out, rail_in

        rail_headers: Dict[str, int] = {}
        for c in range(1, ws_rail.max_column + 1):
            v = ws_rail.cell(row=rail_header, column=c).value
            if isinstance(v, str) and v.strip():
                rail_headers[v.strip().lower()] = c

        def rail_col(name: str) -> int | None:
            return rail_headers.get(name.lower())

        col_owner = rail_col("Владелец записи")
        col_status = rail_col("Статус")
        col_company_id = rail_col("company_id")
        col_start = rail_col("Start waypoint")
        col_end = rail_col("End waypoint")

        if not all([col_owner, col_status, col_company_id, col_start, col_end]):
            print("[builders.build_table3] WARN: required rail columns not found; rail data = 0")
            return rail_out, rail_in

        out_owners = {
            norm_text("ТОО «Институт космической техники и технологий»"),
            norm_text("ЭСФ ESF"),
            norm_text("ТТН TTN"),
            norm_text("КГД Интеграция КЕДЕН KEDEN"),
            norm_text("КЕДЕН KEDEN"),
        }
        in_owner_by_country = {
            norm_text("Белтаможсервис (оператор)"): "BY",
            norm_text('ООО "ЦРЦП" (оператор)'): "RU",
        }

        scanned = 0
        used_out = 0
        used_in = 0

        for row_idx in range(rail_header + 1, ws_rail.max_row + 1):
            status = norm_text(ws_rail.cell(row=row_idx, column=col_status).value)
            if status != "finished":
                continue

            company_id = parse_company_id(ws_rail.cell(row=row_idx, column=col_company_id).value)
            if company_id == "1":
                continue

            owner = norm_text(ws_rail.cell(row=row_idx, column=col_owner).value)
            end_value = ws_rail.cell(row=row_idx, column=col_end).value
            scanned += 1

            if owner in out_owners:
                country = location_country(end_value, {"RU", "BY", "KG"})
                if country in {"RU", "BY", "KG"}:
                    rail_out[country] += 1
                    used_out += 1
                continue

            in_country = in_owner_by_country.get(owner)
            if in_country is not None:
                end_country = location_country(end_value, {"KZ"})
                if end_country == "KZ" or end_country is None:
                    rail_in[in_country] += 1
                    used_in += 1

        print(
            "[builders.build_table3] rail rows: "
            f"finished_no_tests={scanned}, out_used={used_out}, in_used={used_in}"
        )
        print(f"[builders.build_table3] rail_out={rail_out}")
        print(f"[builders.build_table3] rail_in={rail_in}")

        return rail_out, rail_in

    rail_out_values, rail_in_values = calc_rail_flows()

    country_names = {
        "RU": "Российская Федерация",
        "BY": "Республика Беларусь",
        "KG": "Кыргызская Республика",
        "ARM": "Республика Армения",
    }

    out_titles = {
        "RU": "В Российскую Федерацию",
        "BY": "В Республику Беларусь",
        "KG": "В Кыргызскую Республику",
        "ARM": "В Республику Армения",
    }

    in_titles = {
        "RU": "Из Российской Федерации",
        "BY": "Из Республики Беларусь",
        "KG": "Из Кыргызской Республики",
        "ARM": "Из Республики Армения",
    }

    # ---- 3) Авто: из РК
    auto_out_rows: list[Dict[str, Any]] = []
    for code in ["RU", "BY", "KG", "ARM"]:
        mixed_value = flows.get(("KZ", code), 0)
        rail_value = rail_out_values.get(code, 0)
        auto_out_rows.append(
            {
                "Направление": out_titles[code],
                "Количество": max(mixed_value - rail_value, 0),
            }
        )

    auto_out_total = sum(r["Количество"] for r in auto_out_rows)
    auto_out_rows.append(
        {
            "Направление": "Итого",
            "Количество": auto_out_total,
        }
    )

    # ---- 4) Авто: в РК
    auto_in_rows: list[Dict[str, Any]] = []
    for code in ["RU", "BY", "KG", "ARM"]:
        mixed_value = flows.get((code, "KZ"), 0)
        rail_value = rail_in_values.get(code, 0)
        auto_in_rows.append(
            {
                "Направление": in_titles[code],
                "Количество": max(mixed_value - rail_value, 0),
            }
        )

    auto_in_total = sum(r["Количество"] for r in auto_in_rows)
    auto_in_rows.append(
        {
            "Направление": "Итого",
            "Количество": auto_in_total,
        }
    )

    auto_total = auto_out_total + auto_in_total

    # ---- 5) Ж/Д: отдельный источник "ЖД Выгрузка.xlsx"
    rail_out_rows: list[Dict[str, Any]] = []
    for code in ["RU", "BY", "KG", "ARM"]:
        rail_out_rows.append(
            {
                "Направление": out_titles[code],
                "Количество": rail_out_values.get(code, 0),
            }
        )
    rail_out_total = sum(r["Количество"] for r in rail_out_rows)
    rail_out_rows.append(
        {
            "Направление": "Итого",
            "Количество": rail_out_total,
        }
    )

    rail_in_rows: list[Dict[str, Any]] = []
    for code in ["RU", "BY", "KG", "ARM"]:
        rail_in_rows.append(
            {
                "Направление": in_titles[code],
                "Количество": rail_in_values.get(code, 0),
            }
        )
    rail_in_total = sum(r["Количество"] for r in rail_in_rows)
    rail_in_rows.append(
        {
            "Направление": "Итого",
            "Количество": rail_in_total,
        }
    )

    rail_total = rail_out_total + rail_in_total

    # ---- 6) вывод в консоль
    print("[builders.build_table3] start")
    print(f"[T3] block rows: {start_row}..{end_row} (header_row={header_row})")

    print("\n[T3] Завершённые автомобильные перевозки из Республики Казахстан:")
    for r in auto_out_rows:
        print(" ", r)

    print("\n[T3] Завершенные автомобильные перевозки в Республику Казахстан:")
    for r in auto_in_rows:
        print(" ", r)

    print(f"\n[T3 CHECK] Auto: out_total={auto_out_total} in_total={auto_in_total} total={auto_total}")

    print("\n[T3] Завершённые железнодорожные перевозки из Республики Казахстан:")
    for r in rail_out_rows:
        print(" ", r)

    print("\n[T3] Завершенные железнодорожные перевозки в Республику Казахстан:")
    for r in rail_in_rows:
        print(" ", r)

    print(f"\n[T3 CHECK] Rail: out_total={rail_out_total} in_total={rail_in_total} total={rail_total}")
    print("[builders.build_table3] done")

    return {
        "table": "Таблица3",
        "stage": "data-only",

        "auto_out_rows": auto_out_rows,
        "auto_in_rows": auto_in_rows,
        "auto_out_total": auto_out_total,
        "auto_in_total": auto_in_total,
        "auto_total": auto_total,

        "rail_out_rows": rail_out_rows,
        "rail_in_rows": rail_in_rows,
        "rail_out_total": rail_out_total,
        "rail_in_total": rail_in_total,
        "rail_total": rail_total,

        "checks": {
            "auto_out_total": auto_out_total,
            "auto_in_total": auto_in_total,
            "auto_total": auto_total,
            "rail_out_total": rail_out_total,
            "rail_in_total": rail_in_total,
            "rail_total": rail_total,
        },
    }

    def blank_row(title: str) -> Dict[str, Any]:
        return {"Откуда\\Куда": title, "РФ": 0, "РБ": 0, "КР": 0, "РА": 0, "РК (Итого)": 0}

    auto_rows: list[Dict[str, Any]] = []

    # Row: Республика Казахстан (KZ -> *)
    rk = blank_row(row_names["KZ"])
    for d in dirs:
        rk[dir_to_col[d]] = flows.get(("KZ", d), 0)
    rk["РК (Итого)"] = rk["РФ"] + rk["РБ"] + rk["КР"] + rk["РА"]
    auto_rows.append(rk)

    # Rows: RU/BY/KG/ARM (only ->KZ goes into RK total column)
    for d in ["RU", "BY", "KG", "ARM"]:
        rrow = blank_row(row_names[d])
        rrow["РК (Итого)"] = flows.get((d, "KZ"), 0)
        auto_rows.append(rrow)

    # Row: ИТОГО (в РК)
    total_row = blank_row("ИТОГО (в РК)")
    total_row["РФ"] = flows.get(("RU", "KZ"), 0)
    total_row["РБ"] = flows.get(("BY", "KZ"), 0)
    total_row["КР"] = flows.get(("KG", "KZ"), 0)
    total_row["РА"] = flows.get(("ARM", "KZ"), 0)

    out_sum = flows.get(("KZ", "RU"), 0) + flows.get(("KZ", "BY"), 0) + flows.get(("KZ", "KG"), 0) + flows.get(("KZ", "ARM"), 0)
    in_sum = total_row["РФ"] + total_row["РБ"] + total_row["КР"] + total_row["РА"]
    total_row["РК (Итого)"] = out_sum + in_sum
    auto_rows.append(total_row)

    # ---- 5) Ж/Д (по ТЗ сейчас нули)
    rail_rows: list[Dict[str, Any]] = []
    for rr in auto_rows:
        rail_rows.append(
            {
                "Откуда\\Куда": rr["Откуда\\Куда"],
                "РФ": 0,
                "РБ": 0,
                "КР": 0,
                "РА": 0,
                "РК (Итого)": 0,
            }
        )

    # ---- 6) вывод в консоль
    print("[builders.build_table3] start")
    print(f"[T3] block rows: {start_row}..{end_row} (header_row={header_row})")

    print("\n[T3] Завершенные автоперевозки (матрица):")
    for r in auto_rows:
        print(" ", r)

    print("\n[T3] Завершенные железнодорожные перевозки (матрица):")
    for r in rail_rows:
        print(" ", r)

    # Простая проверка: итог совпадает с суммой out+in (мы так и считаем)
    print(f"\n[T3 CHECK] Auto totals: out_sum={out_sum} in_sum={in_sum} total={total_row['РК (Итого)']}")
    print("[builders.build_table3] done")

    return {
        "table": "Таблица3",
        "stage": "data-only",
        "auto_rows": auto_rows,
        "rail_rows": rail_rows,
        "checks": {
            "out_sum": out_sum,
            "in_sum": in_sum,
            "total": total_row["РК (Итого)"],
        },
    }


def build_table4(report_1_norm: Any, report_2_raw: Any, cfg: Dict[str, Any]) -> Any:

    print("[builders.build_table4] start")

    ws = report_2_raw["ws"]

    print(f"[builders.build_table4] sheet = {ws.title}")

    def normalize_name(value: Any) -> str:
        if value is None:
            return ""

        s = str(value).lower()
        s = s.replace("\n", " ")
        s = s.replace("\r", " ")
        s = s.replace("ё", "е")
        s = s.replace("ө", "о")
        s = s.replace("ұ", "у")
        s = s.replace("ү", "у")
        s = s.replace("қ", "к")
        s = s.replace("ғ", "г")
        s = s.replace("ә", "а")
        s = s.replace("і", "и")
        s = s.replace("ң", "н")
        s = re.sub(r"[\"'«»]", " ", s)
        s = re.sub(r"\s+", " ", s)
        return s.strip()

    def parse_date(value: Any) -> dt.date | None:
        if isinstance(value, dt.datetime):
            return value.date()

        if isinstance(value, dt.date):
            return value

        if isinstance(value, str):
            s = value.strip()
            for fmt in ("%d.%m.%Y", "%d.%m.%Y г.", "%Y-%m-%d"):
                try:
                    return dt.datetime.strptime(s, fmt).date()
                except ValueError:
                    pass

        return None

    def parse_count(value: Any) -> int:
        if value is None:
            return 0

        if isinstance(value, (int, float)):
            return int(value)

        if isinstance(value, str):
            s = value.strip().replace(" ", "").replace(",", ".")
            if not s:
                return 0

            try:
                return int(float(s))
            except ValueError:
                return 0

        return 0

    date_candidates: list[tuple[int, dt.date]] = []

    for c in range(2, ws.max_column + 1):
        d = parse_date(ws.cell(row=2, column=c).value)
        if d is not None:
            date_candidates.append((c, d))

    today = dt.date.today()
    actual_candidates = [(c, d) for c, d in date_candidates if d <= today]

    if actual_candidates:
        count_col, actual_date = max(actual_candidates, key=lambda x: (x[1], x[0]))
    elif date_candidates:
        count_col, actual_date = max(date_candidates, key=lambda x: (x[1], x[0]))
    else:
        print("[builders.build_table4] FAIL: date columns not found in row 2")
        return {
            "table": "Таблица4",
            "stage": "data-only",
            "rows": [],
            "date": None,
            "date_label": "",
        }

    print(
        f"[builders.build_table4] actual_date={actual_date:%d.%m.%Y} "
        f"count_col={count_col}"
    )

    source_values: Dict[str, int] = {}

    for r in range(3, ws.max_row + 1):
        point = ws.cell(row=r, column=1).value
        key = normalize_name(point)

        if not key:
            continue

        source_values[key] = parse_count(ws.cell(row=r, column=count_col).value)

    def value_for(*source_names: str) -> int:
        for source_name in source_names:
            key = normalize_name(source_name)
            if key in source_values:
                return source_values[key]

        print(f"[builders.build_table4] WARN: source row not found: {source_names[0]}")
        return 0

    output_points = [
        ("тп «Тажен»", ("Тәжен (авто)", "Тажен (авто)")),
        ("тп «Темир-Баба»", ("Темір-баба (авто)", "Темир-баба (авто)")),
        ("тп «Болашак»", ("Болашақ ж/д", "Болашак ж/д")),
        ("тп «Курык»", ("Құрык-порт ж/д", "Курык-порт ж/д", "Құрық-порт ж/д")),
        ("тп «Морпорт»", ("Ақтау-порт ж/д", "Актау-порт ж/д")),
        ("тп «Оазис»", ("Оазис (Бейнеу) ж/д",)),
        ("тп «Бахты»", ("Бақты (авто)", "Бакты (авто)")),
        ("тп «Майкапчагай»", ("Майқапшағай (авто)", "Майкапчагай (авто)")),
        ("тп «Б. Конысбаева»", ("Б. Қонысбаева (авто)", "Б. Конысбаева (авто)")),
        ("тп «Казыгурт»", ("Қазығурт (авто)", "Казыгурт (авто)", "Казығұрт (авто)")),
        ("тп «ст. Сарыагаш»", ("Сарыағаш ж/д", "Сарыагаш ж/д")),
        ("тп «Атамекен»", ("Атамекен (авто)",)),
        ("тп «Калжат»", ("Қалжат (авто)", "Калжат (авто)")),
        ("тп «Алаколь»", ("Алакөл (авто)", "Алаколь (авто)")),
        ("тп «Темиржол»", ("Теміржол ж/д", "Темиржол ж/д")),
        ("тп «Алтынколь-жол»", ("Алтынкөл-жол ж/д", "Алтынколь-жол ж/д")),
        ("тп «Нур Жолы»", ("Нұржолы (авто)", "Нур жолы (авто)", "Нуржолы (авто)")),
        ("г. Астана склад", ("г. Астана",)),
        (
            "г. Караганда склад",
            ("Карагандинская область\n г. Қарағанды\n (склад авто по взаимке)",),
        ),
        ("г. Шымкент, склад", ("г. Шымкент",)),
        (
            "г. Петропавловск, склад",
            ("Северо-Казахстанская область\n г. Петропавл\n (склад авто по взаимке)",),
        ),
        (
            "г. Кокшетау, склад",
            ("Акмолинская область\n г. Көкшетау\n (склад авто по взаимке)",),
        ),
        (
            "г. Уральск, склад",
            ("Западно-Казахстанская область\n г. Орал\n (склад авто по взаимке)",),
        ),
        (
            "г. Актобе, склад",
            ("Актюбинская область\n г. Ақтөбе\n (склад авто по взаимке)",),
        ),
        (
            "г. Павлодар склад",
            ("Павлодарская область\n г. Павлодар\n (склад авто по взаимке)",),
        ),
        (
            "г. Оскемен, склад",
            ("Восточно- Казахстанская область\n г. Өскемен\n (склад авто по взаимке)",),
        ),
        (
            "г. Семей, склад",
            ("Область Абай\n г. Семей\n (склад авто по взаимке)",),
        ),
        ("г. Актау", ("г. Ақтау", "г. Актау")),
        ("Атырауская область", ("Атырауская область",)),
        (
            "ЗКО",
            ("Западно-Казахстанская область\n г. Орал\n (склад авто по взаимке)",),
        ),
        (
            "Костанайская область",
            ("Костанайская область\n г. Костанай\n (склад авто по взаимке)",),
        ),
    ]

    # Эти точки есть в источнике, но во втором шаблоне отчета не выводятся:
    # Жибек жолы, Алматинская область, Область Улытау, Кызылординская область,
    # Жамбылская область, Область Жетису, Капланбек.

    rows = [
        {
            "Точка": output_name,
            "Количество НП": value_for(*source_names),
        }
        for output_name, source_names in output_points
    ]

    def calc_used_np_total() -> int:
        ws_report = report_1_norm.get("ws") if isinstance(report_1_norm, dict) else None
        if ws_report is None:
            print("[builders.build_table4] WARN: report sheet not found for used NP total")
            return 0

        start_row = None
        header_row = None
        point_col = None
        total_col = None
        max_scan_cols = min(ws_report.max_column, 40)

        for r in range(1, ws_report.max_row + 1):
            for c in range(1, max_scan_cols + 1):
                value = ws_report.cell(row=r, column=c).value
                if isinstance(value, str) and "статистика пломб по типам перевозки" in value.lower():
                    start_row = r
                    break
            if start_row is not None:
                break

        if start_row is None:
            print("[builders.build_table4] WARN: used NP source block not found")
            return 0

        for r in range(start_row + 1, min(ws_report.max_row, start_row + 8) + 1):
            for c in range(1, max_scan_cols + 1):
                value = ws_report.cell(row=r, column=c).value
                if isinstance(value, str) and value.strip().lower() == "пункт отправления":
                    header_row = r
                    point_col = c
                    break
            if header_row is not None:
                break

        if header_row is None or point_col is None:
            print("[builders.build_table4] WARN: used NP header row not found")
            return 0

        for c in range(point_col + 1, ws_report.max_column + 1):
            value = ws_report.cell(row=header_row, column=c).value
            if isinstance(value, str) and value.strip().lower() == "total":
                total_col = c
                break

        if total_col is None:
            print("[builders.build_table4] WARN: used NP Total column not found")
            return 0

        total = 0
        for r in range(header_row + 1, ws_report.max_row + 1):
            point_value = ws_report.cell(row=r, column=point_col).value
            point_key = normalize_name(point_value)

            if not point_key:
                continue

            if point_key in {"все", "итого"}:
                break

            if point_key == normalize_name("Карасу"):
                continue

            total += parse_count(ws_report.cell(row=r, column=total_col).value)

        print(f"[builders.build_table4] used_np_total={total}")
        return total

    used_np_total = calc_used_np_total()

    print(f"[builders.build_table4] rows={len(rows)}")

    return {
        "table": "Таблица4",
        "stage": "data-only",
        "date": actual_date,
        "date_label": actual_date.strftime("%d.%m.%Y"),
        "date_label_short": actual_date.strftime("%d.%m.%y"),
        "used_np_total": used_np_total,
        "rows": rows,
    }


def build_table5(report_1_norm: Any, cfg: Dict[str, Any]) -> Any:
    """
    Т5: Статистика по перевозкам, направленным в сторону Республики Казахстан

    Блок:
        start: строка содержит "Показатель" и рядом "Значение"
        end:   строка содержит "Перевозки по типам кортежей"

    Учитываем только:
        From ∈ {RU,BY,KG,ARM}
        To = KZ

    Значение:
        колонка "Активированная перевозка"
    """

    ws = report_1_norm["ws"]

    start_row = None
    end_row = None

    max_scan_cols = min(ws.max_column, 30)

    def row_cells_str(r: int) -> list[str]:
        vals: list[str] = []
        for c in range(1, max_scan_cols + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip():
                vals.append(v.strip())
        return vals

    # ---- поиск начала блока
    for r in range(1, ws.max_row + 1):
        for c in range(1, max_scan_cols):
            a = ws.cell(row=r, column=c).value
            b = ws.cell(row=r, column=c + 1).value
            if isinstance(a, str) and isinstance(b, str):
                if a.strip().lower() == "показатель" and b.strip().lower() == "значение":
                    start_row = r
                    break
        if start_row is not None:
            break

    if start_row is None:
        print("[builders.build_table5] FAIL: start marker not found")
        return {"table": "Таблица5", "stage": "data-only", "rows": []}

    # ---- поиск конца блока
    for r in range(start_row + 1, ws.max_row + 1):
        vals = row_cells_str(r)
        if any("перевозки по типам кортежей" in v.lower() for v in vals):
            end_row = r
            break

    if end_row is None:
        print("[builders.build_table5] FAIL: end marker not found")
        return {"table": "Таблица5", "stage": "data-only", "rows": []}

    # ---- поиск header строки
    header_row = None
    for r in range(start_row + 1, end_row):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip() == "From":
            header_row = r
            break

    if header_row is None:
        print("[builders.build_table5] FAIL: header row not found")
        return {"table": "Таблица5", "stage": "data-only", "rows": []}

    # ---- map колонок
    header_map: Dict[str, int] = {}

    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, str) and v.strip():
            header_map[v.strip().lower()] = c

    def find_col(substr: str) -> int | None:
        s = substr.lower()
        for k, col in header_map.items():
            if s in k:
                return col
        return None

    col_from = find_col("from")
    col_to = find_col("to")
    col_active = find_col("актив")

    if not all([col_from, col_to, col_active]):
        print("[builders.build_table5] FAIL: required columns not found")
        print(f" col_from={col_from}, col_to={col_to}, col_active={col_active}")
        return {"table": "Таблица5", "stage": "data-only", "rows": []}

    def parse_int(v: Any) -> int:
        if isinstance(v, (int, float)):
            return int(v)
        if isinstance(v, str):
            s = v.strip().replace(" ", "")
            if s.isdigit():
                return int(s)
        return 0

    dirs = ["RU", "BY", "KG", "ARM"]

    sums: Dict[str, int] = {k: 0 for k in dirs}

    # ---- парс строк
    for r in range(header_row + 1, end_row):

        frm = ws.cell(row=r, column=col_from).value
        to = ws.cell(row=r, column=col_to).value

        frm_s = frm.strip().upper() if isinstance(frm, str) else ""
        to_s = to.strip().upper() if isinstance(to, str) else ""

        if frm_s == "" and to_s == "":
            continue

        if to_s == "?":
            continue

        if frm_s == "KZ" and to_s == "ВСЕ":
            continue

        if frm_s == "KZ" and to_s == "KZ":
            continue

        if to_s == "KZ" and frm_s in dirs:

            val = parse_int(ws.cell(row=r, column=col_active).value)

            sums[frm_s] += val

    full_names = {
        "RU": "Российская Федерация",
        "BY": "Республика Беларусь",
        "KG": "Кыргызская Республика",
        "ARM": "Республика Армения",
    }

    rows = []

    for code in ["RU", "BY", "KG", "ARM"]:
        rows.append(
            {
                "Страна начала перевозки": full_names[code],
                "Количество перевозок": sums[code],
            }
        )

    total = sum(sums.values())

    rows.append(
        {
            "Страна начала перевозки": "Итого",
            "Количество перевозок": total,
        }
    )

    # ---- вывод
    print("[builders.build_table5] start")
    print(f"[T5] block rows: {start_row}..{end_row} (header_row={header_row})")

    for r in rows:
        print(" ", r)

    print(f"[T5 CHECK] total = {total}")

    print("[builders.build_table5] done")

    return {
        "table": "Таблица5",
        "stage": "data-only",
        "rows": rows,
        "total": total,
    }
