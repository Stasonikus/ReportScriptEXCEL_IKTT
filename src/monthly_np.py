from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path
from typing import Any, Dict, Iterable


SNAPSHOT_PATTERN = re.compile(r"^np_snapshot_(\d{8})\.json$")
REPORT_DATE_PATTERN = re.compile(r"report_(\d{2})_(\d{2})_(\d{4})", re.IGNORECASE)


def parse_report_date(path: Path, fallback: dt.date | None = None) -> dt.date:
    match = REPORT_DATE_PATTERN.search(path.name)
    if match:
        day, month, year = (int(part) for part in match.groups())
        return dt.date(year, month, day)
    return fallback or dt.date.today()


def _number(value: Any) -> int | None:
    if value in {None, "-"}:
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        text = value.strip().replace(" ", "").replace(",", ".")
        if not text or text == "-":
            return None
        return int(float(text))
    return None


def make_snapshot(
    rows: Iterable[Dict[str, Any]],
    snapshot_date: dt.date,
    source_name: str,
) -> Dict[str, Any]:
    values: Dict[str, int] = {}
    for row in rows:
        name = str(row.get("Пункт", "")).strip()
        if not name:
            continue
        value = _number(row.get("Количество"))
        values[name] = value if value is not None else 0

    return {
        "version": 1,
        "snapshot_date": snapshot_date.isoformat(),
        "source": source_name,
        "used_np_rows": values,
    }


def save_snapshot(snapshot_dir: Path, snapshot: Dict[str, Any]) -> Path:
    snapshot_date = dt.date.fromisoformat(str(snapshot["snapshot_date"]))
    snapshot_dir.mkdir(parents=True, exist_ok=True)
    destination = snapshot_dir / f"np_snapshot_{snapshot_date:%Y%m%d}.json"
    temporary = destination.with_suffix(".json.tmp")
    temporary.write_text(
        json.dumps(snapshot, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    temporary.replace(destination)
    return destination


def _load_snapshot(path: Path, storage: str = "") -> Dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or not isinstance(data.get("used_np_rows"), dict):
        raise ValueError(f"invalid monthly NP snapshot: {path}")
    dt.date.fromisoformat(str(data.get("snapshot_date")))
    if storage:
        data["_baseline_storage"] = storage
    return data


def load_previous_month_snapshot(
    report_date: dt.date,
    snapshot_dir: Path,
    bundled_data_dir: Path,
) -> tuple[Dict[str, Any] | None, bool]:
    current_month_start = report_date.replace(day=1)
    previous_month_end = current_month_start - dt.timedelta(days=1)
    previous_month_start = previous_month_end.replace(day=1)
    exact_name = f"np_snapshot_{previous_month_end:%Y%m%d}.json"

    sources = (
        (snapshot_dir, "user history"),
        (bundled_data_dir, "embedded"),
    )
    for directory, storage in sources:
        exact_path = directory / exact_name
        if exact_path.exists():
            try:
                return _load_snapshot(exact_path, storage), True
            except (OSError, ValueError, json.JSONDecodeError):
                # A damaged user snapshot must not block the verified embedded
                # fallback shipped with the executable.
                continue

    candidates: list[tuple[dt.date, int, Path, str]] = []
    for priority, (directory, storage) in enumerate(sources):
        if not directory.exists():
            continue
        for path in directory.glob("np_snapshot_*.json"):
            match = SNAPSHOT_PATTERN.match(path.name)
            if not match:
                continue
            snapshot_date = dt.datetime.strptime(match.group(1), "%Y%m%d").date()
            if previous_month_start <= snapshot_date <= previous_month_end:
                candidates.append((snapshot_date, -priority, path, storage))

    if not candidates:
        return None, False

    for _, _, path, storage in sorted(candidates, reverse=True):
        try:
            return _load_snapshot(path, storage), False
        except (OSError, ValueError, json.JSONDecodeError):
            continue
    return None, False


def calculate_monthly_rows(
    current_rows: Iterable[Dict[str, Any]],
    baseline: Dict[str, Any] | None,
) -> list[Dict[str, Any]]:
    baseline_values = (baseline or {}).get("used_np_rows", {}) or {}
    result: list[Dict[str, Any]] = []
    calculated_total = 0

    for row in current_rows:
        name = str(row.get("Пункт", "")).strip()
        if not name:
            continue
        if name == "Итого":
            continue

        current_value = _number(row.get("Количество"))
        if baseline is None or current_value is None:
            monthly_value: int | str = "-"
        else:
            baseline_value = _number(baseline_values.get(name)) or 0
            monthly_value = current_value - baseline_value
            calculated_total += monthly_value

        result.append({"Пункт": name, "Количество": monthly_value})

    result.append(
        {
            "Пункт": "Итого",
            "Количество": calculated_total if baseline is not None else "-",
        }
    )
    return result

