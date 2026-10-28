# src/config.py
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Set


# Фиксированный список строк для столбца "ТП / ЖД" (Т1)
TP_ZHD_LIST_T1 = [
    "Алаколь",
    "Казыгурт",
    "Нур жолы",
    "Бахты",
    "Калжат",
    "Тажен",
    "Атамекен",
    "Б. конысбаева",
    "Майкапчагай",
    "Темир-баба",
    "Капланбек",
    "Морпорт актау",
    "Курык порт",
    "Курык жд",
    "Болашак",
    "Оазис",
    "Мактаарал",
    "Сарыагаш",
    "Алтынколь",
    "Темиржол",
]


# Нормализованный словарь
NORMALIZATION: Dict[str, Any] = {
    "countries": {
        "RU": "РФ",
        "BY": "РБ",
        "KG": "КР",
        "ARM": "РА",
        "РФ": "РФ",
        "РБ": "РБ",
        "КР": "КР",
        "РА": "РА",
        "Российская Федерация": "РФ",
        "Республика Беларусь": "РБ",
        "Кыргызская Республика": "КР",
        "Республика Армения": "РА",
    },
    # маркеры пункта пропуска
    "pp_markers": ["т/п", "тп", "таможенный пост"],
    # очистка строк
    "quotes_to_strip": ['"', "«", "»"],
    # алиасы пунктов пропуска
    "pp_aliases": {
        "Имени бауыржана конысбаева": "Б. конысбаева",
    },
    # служебные значения
    "service": {
        "missing_direction_value": 0,
        "ignore_arrow_zero": True,
        "unresolved_marker": "?",
    },
    # исключения
    "exclude_names": {"ИКТТ Алматы", "Карасу"},
}


@dataclass(frozen=True)
class AppConfig:
    tp_zhd_list_t1: List[str]
    normalization: Dict[str, Any]
    exclude_names: Set[str]


def load_config() -> AppConfig:
    return AppConfig(
        tp_zhd_list_t1=TP_ZHD_LIST_T1,
        normalization=NORMALIZATION,
        exclude_names=set(NORMALIZATION.get("exclude_names", set())),
    )
