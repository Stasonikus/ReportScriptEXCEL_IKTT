from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, Iterable


STAGE2_BASELINE_FILE = "stage2_baseline_20260831.json"


def load_stage2_baseline(path: Path) -> Dict[str, Any]:
    """Load the fixed cumulative snapshot that closes stage 1."""
    if not path.exists():
        raise FileNotFoundError(f"stage 2 baseline not found: {path}")

    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or not isinstance(data.get("tables"), dict):
        raise ValueError(f"invalid stage 2 baseline: {path}")

    return data


def _number(value: Any) -> int:
    if value is None or value == "-":
        return 0

    if isinstance(value, bool):
        return int(value)

    if isinstance(value, (int, float)):
        return int(value)

    if isinstance(value, str):
        text = value.strip().replace(" ", "").replace(",", ".")
        if not text or text == "-":
            return 0
        return int(float(text))

    return 0


def _subtract_named_rows(
    rows: list[Dict[str, Any]],
    baseline: Dict[str, Any],
    *,
    name_key: str,
    value_key: str,
    total_name: str = "Итого",
) -> int:
    total = 0

    for row in rows:
        name = str(row.get(name_key, "")).strip()
        if not name or name == total_name:
            continue

        current_value = _number(row.get(value_key))
        baseline_value = _number(baseline.get(name))
        delta = current_value - baseline_value
        row[value_key] = delta
        total += delta

    for row in rows:
        if str(row.get(name_key, "")).strip() == total_name:
            row[value_key] = total

    return total


def _apply_table1(table: Dict[str, Any], baseline: Dict[str, Any]) -> None:
    columns = ("РФ", "РБ", "КР", "РА")
    baseline_rows = baseline.get("rows", {}) or {}
    rows = table.get("rows", []) or []

    totals = {column: 0 for column in columns}

    for row in rows:
        name = str(row.get("ТП/ЖД", "")).strip()
        if not name or name == "Итого":
            continue

        base_row = baseline_rows.get(name, {}) or {}
        for column in columns:
            delta = _number(row.get(column)) - _number(base_row.get(column))
            row[column] = delta
            totals[column] += delta

        row["Итого"] = sum(_number(row.get(column)) for column in columns)

    for row in rows:
        if str(row.get("ТП/ЖД", "")).strip() == "Итого":
            for column in columns:
                row[column] = totals[column]
            row["Итого"] = sum(totals.values())

def _apply_table2(table: Dict[str, Any], baseline: Dict[str, Any]) -> None:
    for section in ("export_rows", "trade_rows"):
        total = _subtract_named_rows(
            table.get(section, []) or [],
            baseline.get(section, {}) or {},
            name_key="Страна",
            value_key="Количество перевозок",
        )
        table[f"stage2_{section}_total"] = total


def _apply_table3(table: Dict[str, Any], baseline: Dict[str, Any]) -> None:
    section_totals: Dict[str, int] = {}
    rail_is_stage2_direct = bool(table.get("rail_values_are_stage2_direct"))

    # В старом методе ЖД-строки уже формируются непосредственно из ЖД выгрузки
    # с датой от 01.09.2026. Вычитать из них снимок первого этапа второй раз
    # нельзя. При этом общий накопительный report содержит историческую ЖД
    # часть, поэтому она дополнительно вычитается только из авто-строк.
    direct_rail_pairs = {
        "auto_out_rows": "rail_out_rows",
        "auto_in_rows": "rail_in_rows",
    }

    for section in (
        "auto_out_rows",
        "auto_in_rows",
        "rail_out_rows",
        "rail_in_rows",
    ):
        rows = table.get(section, []) or []
        if rail_is_stage2_direct and section in direct_rail_pairs:
            own_baseline = baseline.get(section, {}) or {}
            rail_baseline = baseline.get(direct_rail_pairs[section], {}) or {}
            combined_baseline = {
                name: _number(own_baseline.get(name)) + _number(rail_baseline.get(name))
                for name in set(own_baseline) | set(rail_baseline)
            }
            section_totals[section] = _subtract_named_rows(
                rows,
                combined_baseline,
                name_key="Направление",
                value_key="Количество",
            )
        elif rail_is_stage2_direct and section.startswith("rail_"):
            section_totals[section] = sum(
                _number(row.get("Количество"))
                for row in rows
                if str(row.get("Направление", "")).strip() != "Итого"
            )
            for row in rows:
                if str(row.get("Направление", "")).strip() == "Итого":
                    row["Количество"] = section_totals[section]
        else:
            section_totals[section] = _subtract_named_rows(
                rows,
                baseline.get(section, {}) or {},
                name_key="Направление",
                value_key="Количество",
            )

    table["auto_out_total"] = section_totals["auto_out_rows"]
    table["auto_in_total"] = section_totals["auto_in_rows"]
    table["auto_total"] = table["auto_out_total"] + table["auto_in_total"]
    table["rail_out_total"] = section_totals["rail_out_rows"]
    table["rail_in_total"] = section_totals["rail_in_rows"]
    table["rail_total"] = table["rail_out_total"] + table["rail_in_total"]


def _apply_table4(table: Dict[str, Any], baseline: Dict[str, Any]) -> None:
    """
    Keep current NP stock (left table) unchanged because it is a snapshot.
    Subtract only cumulative used-NP values in the right table.
    """
    baseline_rows = baseline.get("used_np_rows", {}) or {}
    rows = table.get("used_np_rows", []) or []
    total = 0

    for row in rows:
        name = str(row.get("Пункт", "")).strip()
        if not name or name == "Итого":
            continue

        current_raw = row.get("Количество")
        baseline_has_value = name in baseline_rows

        if current_raw in {None, "-"} and not baseline_has_value:
            row["Количество"] = "-"
            continue

        delta = _number(current_raw) - _number(baseline_rows.get(name))
        row["Количество"] = delta
        total += delta

    for row in rows:
        if str(row.get("Пункт", "")).strip() == "Итого":
            row["Количество"] = total

    table["used_np_total"] = total


def apply_stage2_baseline(
    tables: Iterable[Dict[str, Any]],
    baseline: Dict[str, Any],
) -> list[Dict[str, Any]]:
    """Convert cumulative source tables X to stage-2 values Z = X - Y."""
    result = deepcopy(list(tables))
    baseline_tables = baseline.get("tables", {}) or {}

    # Файл базы хранит исторические ключи Т3/Т4. После вставки новой
    # Таблицы3 прежние расчёты завершённых перевозок и НП стали Т4/Т5.
    # Новая Т3 уже отфильтрована непосредственно с 01.09.2026 и вычитания
    # накопительной базы не требует.
    handlers = {
        "Таблица1": (_apply_table1, "Таблица1"),
        "Таблица2": (_apply_table2, "Таблица2"),
        "Таблица4": (_apply_table3, "Таблица3"),
        "Таблица5": (_apply_table4, "Таблица4"),
    }

    for table in result:
        table_name = str(table.get("table", ""))
        handler_spec = handlers.get(table_name)
        if handler_spec is None:
            continue
        handler, baseline_name = handler_spec
        table_baseline = baseline_tables.get(baseline_name)

        if not isinstance(table_baseline, dict):
            continue

        handler(table, table_baseline)
        table["stage"] = "stage2-delta"
        table["baseline_cutoff"] = baseline.get("cutoff")

    return result
