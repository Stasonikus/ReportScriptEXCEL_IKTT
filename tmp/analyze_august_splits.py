from __future__ import annotations

import datetime as dt
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from big_export import (  # noqa: E402
    ALLOWED_AUTO_OUT_OWNER_PARTS,
    COUNTED_CURRENT_STATUSES,
    FINISHED_STATUSES,
    REQUIRED_HEADERS,
    _country_code,
    _norm,
    _open_big_export_source,
    _parse_datetime,
    _transport,
    _transportation_type,
    detect_transport_by_vehicle_numbers,
)


SOURCE = Path(
    r"C:\Users\admin\Desktop\Project Python\ReportScriptEXCEL_IKTT 2 STAGE"
    r"\in\processed\file (1)__processed_20260925_1119.xls"
)
VEHICLES = SOURCE.parents[2] / "src" / "data" / "vehicle_plate_formats.json"
START = dt.datetime(2026, 7, 31, 0, 0, 0)
END = dt.datetime(2026, 9, 1, 0, 0, 0)
TARGETS = {"RU", "BY", "KG", "ARM"}
THROUGH = {("BY", "KG"), ("RU", "KG"), ("KG", "RU"), ("KG", "BY")}


def in_period(value: dt.datetime | None) -> bool:
    return value is not None and START <= value < END


def main() -> None:
    workbook, header_row, header_map, source_rows = _open_big_export_source(SOURCE)
    indexes = {name: header_map[name] for name in REQUIRED_HEADERS}
    records: dict[str, dict] = {}
    duplicates = 0
    try:
        for source_row, values in enumerate(source_rows, start=header_row + 1):
            if not any(value not in (None, "") for value in values):
                continue
            row = {
                name: values[index] if index < len(values) else None
                for name, index in indexes.items()
            }
            created = _parse_datetime(row["Дата создания"])
            updated = _parse_datetime(row["Дата обновления"])
            number = str(row["Номер перевозки"] or "").strip()
            key = number or f"__row_{source_row}"
            row["_created"] = created
            row["_updated"] = updated
            row["_source_row"] = source_row
            current = records.get(key)
            order = (updated or created or dt.datetime.min, source_row)
            if current is None:
                records[key] = row
            else:
                duplicates += 1
                current_order = (
                    current["_updated"] or current["_created"] or dt.datetime.min,
                    current["_source_row"],
                )
                if order > current_order:
                    records[key] = row
    finally:
        workbook.close()

    prepared = []
    for row in records.values():
        declared = _transport(row["Тип транспорта"])
        detected = detect_transport_by_vehicle_numbers(row["Номера ТС"], VEHICLES)
        production = detected or declared or "auto"
        prepared.append(
            {
                "created": row["_created"],
                "updated": row["_updated"],
                "status": _norm(row["Статус"]),
                "kind": _transportation_type(row["Тип перевозки"]),
                "declared": declared or "unknown",
                "production": production,
                "country_from": _country_code(row["Страна начала"]),
                "country_to": _country_code(row["Страна завершения"]),
                "owner": _norm(row["Владелец"]),
            }
        )

    def summarize(name: str, rows: list[dict], field: str = "production") -> None:
        values = Counter(item[field] for item in rows)
        print(name, "TOTAL", len(rows), "AUTO", values["auto"], "RAIL", values["rail"], "UNKNOWN", values["unknown"])

    base_created = [r for r in prepared if in_period(r["created"])]
    current_created = [r for r in base_created if r["status"] in COUNTED_CURRENT_STATUSES]
    outgoing_current = [
        r for r in current_created
        if r["country_from"] == "KZ" and r["country_to"] in TARGETS
    ]
    owner_outgoing_current = [
        r for r in outgoing_current
        if any(part in r["owner"] for part in ALLOWED_AUTO_OUT_OWNER_PARTS)
    ]

    print("SOURCE", SOURCE)
    print("UNIQUE", len(records), "DUPLICATES", duplicates)
    print("PERIOD", START, END)
    for label, rows in (
        ("CREATED_ALL", base_created),
        ("CREATED_ACTIVE_FINISHED", current_created),
        ("OUTGOING_CURRENT", outgoing_current),
        ("OWNER_OUTGOING_CURRENT", owner_outgoing_current),
    ):
        summarize(label + " DECLARED", rows, "declared")
        summarize(label + " PRODUCTION", rows, "production")

    for source_name, source_rows_filtered in (
        ("OUTGOING_CURRENT", outgoing_current),
        ("OWNER_OUTGOING_CURRENT", owner_outgoing_current),
        ("CURRENT_ALL_COUNTRIES", current_created),
    ):
        for kind in ("transit", "export", "trade"):
            subset = [r for r in source_rows_filtered if r["kind"] == kind]
            summarize(f"{source_name} {kind} DECLARED", subset, "declared")
            summarize(f"{source_name} {kind} PRODUCTION", subset, "production")

    monitoring = [
        r for r in current_created
        if r["kind"] == "transit"
        and (r["country_from"], r["country_to"]) in THROUGH
    ]
    summarize("MONITORING CREATED DECLARED", monitoring, "declared")
    summarize("MONITORING CREATED PRODUCTION", monitoring, "production")

    for date_field in ("created", "updated"):
        for statuses_name, statuses in (
            ("COMPLETED", {"завершенная перевозка"}),
            ("FINISHED", FINISHED_STATUSES),
            ("ACTIVE_FINISHED", COUNTED_CURRENT_STATUSES),
        ):
            for geography_name, predicate in (
                ("TO_KZ", lambda r: r["country_to"] == "KZ" and r["country_from"] in TARGETS),
                ("ANY_END_KZ", lambda r: r["country_to"] == "KZ"),
                ("ALL", lambda r: True),
            ):
                subset = [
                    r for r in prepared
                    if in_period(r[date_field])
                    and r["status"] in statuses
                    and predicate(r)
                ]
                summarize(
                    f"END {date_field.upper()} {statuses_name} {geography_name} DECLARED",
                    subset,
                    "declared",
                )
                summarize(
                    f"END {date_field.upper()} {statuses_name} {geography_name} PRODUCTION",
                    subset,
                    "production",
                )


if __name__ == "__main__":
    main()
