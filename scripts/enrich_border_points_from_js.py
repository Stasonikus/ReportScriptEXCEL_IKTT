"""Добавляет к справочнику координаты только из переданных JS-источников."""
from __future__ import annotations

import json
import re
import sys
import unicodedata
from pathlib import Path


TRANSLITERATION = str.maketrans({
    "ё": "е", "қ": "к", "ғ": "г", "ң": "н", "ү": "у",
    "ұ": "у", "ә": "а", "і": "и",
})


def normalize(value: str) -> str:
    text = unicodedata.normalize("NFKD", value.lower()).translate(TRANSLITERATION)
    return re.sub(r"[^a-zа-я0-9]+", "", text)


def read_js_points(path: Path) -> list[dict[str, object]]:
    pattern = re.compile(
        r'name:\s*"(?P<name>[^"]+)"\s*,\s*'
        r"latitude:\s*(?P<latitude>-?[\d.]+)\s*,\s*"
        r"longitude:\s*(?P<longitude>-?[\d.]+)",
        re.DOTALL,
    )
    return [
        {
            "name": match.group("name"),
            "latitude": float(match.group("latitude")),
            "longitude": float(match.group("longitude")),
        }
        for match in pattern.finditer(path.read_text(encoding="utf-8"))
    ]


# Отличающиеся, но подтверждённо одинаковые написания из исходных JS-файлов.
POINT_ID_BY_SOURCE_NAME = {
    ("customs.js", "кольжат"): "kalzhat_auto",
    ("train.js", "махтаарал"): "maktaaral_rail",
    ("train.js", "пресногорьковская"): "presnogorkovka_rail",
    ("train.js", "магнаи"): "mangai_rail",
    ("ship.js", "порткурык"): "kuryk_port_rail",
}

EXTRA_POINT_IDS = {
    ("customs.js", "аксаимапп"): ("aksai_mapp_auto", "Аксай МАПП"),
    ("customs.js", "бидаик"): ("bidaik_auto", "Бидаик"),
    ("customs.js", "жибекжолыnчерняевка"): (
        "zhibek_zholy_chernyaevka_auto", "Жибек Жолы (Черняевка)"
    ),
    ("customs.js", "каракога"): ("karakoga_auto", "Каракога"),
    ("customs.js", "кордаи"): ("kordai_auto", "Кордай"),
    ("customs.js", "шаган"): ("shagan_auto", "Шаган"),
    ("customs.js", "шарбакты"): ("sharbakty_auto", "Шарбакты"),
    ("train.js", "жанибек"): ("zhanibek_rail", "Жанибек"),
    ("train.js", "киригильда"): ("kirigilda_rail", "Киригильда"),
    ("train.js", "тобыл"): ("tobyl_rail", "Тобыл"),
    ("train.js", "чалдовар"): ("chaldovar_rail", "Чалдовар"),
}


def add_port_points(points: list[dict[str, object]]) -> None:
    existing = {str(point["id"]) for point in points}
    additions = [
        {
            "id": "aktau_port_rail",
            "name": "Порт Актау",
            "transport": "жд",
            "border": "внешняя",
            "aliases": ["Актау-порт", "Морпорт Актау"],
            "source_files": ["ship.js"],
        },
        {
            "id": "aktau_north_terminal_rail",
            "name": "Актау, Северный терминал",
            "transport": "жд",
            "border": "внешняя",
            "aliases": ["Северный терминал Актау"],
            "source_files": ["ship.js"],
        },
    ]
    points.extend(point for point in additions if point["id"] not in existing)


def enrich(target: Path, source_dir: Path) -> dict[str, int]:
    data = json.loads(target.read_text(encoding="utf-8"))
    points = data["points"]
    add_port_points(points)
    by_id = {str(point["id"]): point for point in points}
    aliases: dict[tuple[str, str], str] = {}
    sources = {
        "customs.js": "авто",
        "train.js": "жд",
        "ship.js": "жд",
    }

    for source_file, transport in sources.items():
        for point in points:
            if point.get("transport") != transport:
                continue
            for value in [str(point.get("name", "")), *point.get("aliases", [])]:
                aliases[(source_file, normalize(str(value)))] = str(point["id"])

    matched = 0
    unmatched = 0
    for source_file in sources:
        for item in read_js_points(source_dir / source_file):
            source_name = normalize(str(item["name"]))
            point_id = POINT_ID_BY_SOURCE_NAME.get((source_file, source_name))
            point_id = point_id or aliases.get((source_file, source_name))
            if point_id is None:
                addition = EXTRA_POINT_IDS.get((source_file, source_name))
                if addition is None:
                    unmatched += 1
                    continue
                point_id, display_name = addition
                point = {
                    "id": point_id,
                    "name": display_name,
                    "transport": sources[source_file],
                    "border": "не определено",
                    "aliases": [str(item["name"])],
                    "classification_status": "требует решения",
                    "note": (
                        "Вид транспорта подтвержден исходным JS-файлом; "
                        "граница требует отдельного источника."
                    ),
                    "source_files": [source_file],
                }
                points.append(point)
                by_id[point_id] = point
                aliases[(source_file, source_name)] = point_id
            point = by_id[point_id]
            point["latitude"] = item["latitude"]
            point["longitude"] = item["longitude"]
            point["coordinate_source"] = source_file
            matched += 1

    data["coordinate_policy"] = (
        "Координаты указываются только при точном или явно сопоставленном "
        "совпадении с customs.js, train.js либо ship.js."
    )
    data["points"] = points
    target.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return {"matched": matched, "unmatched": unmatched, "total": len(points)}


if __name__ == "__main__":
    destination = Path(sys.argv[1])
    source_directory = Path(sys.argv[2])
    print(enrich(destination, source_directory))
