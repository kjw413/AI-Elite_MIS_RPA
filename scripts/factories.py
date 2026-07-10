"""Factory domain master data and aggregation helpers."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Mapping

import pandas as pd


@dataclass(frozen=True)
class Factory:
    label: str
    f_code: str
    weather_station: str
    color: str
    production_parent: str | None = None


FACTORY_MASTER: tuple[Factory, ...] = (
    Factory("남양주1", "F10A", "서울", "#a855f7", "F10"),
    Factory("남양주2", "F10B", "서울", "#7c3aed", "F10"),
    Factory("김해", "F20", "김해", "#f97316", "F20"),
    Factory("광주", "F30", "이천", "#22c55e", "F30"),
    Factory("논산", "F40", "부여", "#ef4444", "F40"),
    Factory("경산", "F50", "대구", "#14b8a6", "F50"),  # 2026-07 신규 (냉동)
)

FACTORY_PHYSICAL_DISPLAY_ORDER: tuple[str, ...] = tuple(f.label for f in FACTORY_MASTER)
FACTORY_DISPLAY_ORDER: tuple[str, ...] = (
    "전사", "남양주1", "남양주2", "남양주", "김해", "광주", "논산", "경산",
)
FACTORY_FILTER_OPTIONS: tuple[str, ...] = (
    "남양주1", "남양주2", "남양주", "김해", "광주", "논산", "경산",
)
FACTORY_AGGREGATE_DISPLAY_ORDER: tuple[str, ...] = (
    "전사", "남양주", "김해", "광주", "논산", "경산",
)
DASHBOARD_FACTORY_ORDER: tuple[str, ...] = ("남양주", "김해", "광주", "논산", "경산")

BATCH_ALL_FACTORIES_LABEL = "전체"
PREDICTION_FACTORY_OPTIONS: tuple[str, ...] = (
    BATCH_ALL_FACTORIES_LABEL,
) + FACTORY_DISPLAY_ORDER

AGGREGATE_FACTORY_MEMBERS: dict[str, tuple[str, ...]] = {
    "남양주": ("남양주1", "남양주2"),
    "전사": FACTORY_PHYSICAL_DISPLAY_ORDER,
    "전체": FACTORY_PHYSICAL_DISPLAY_ORDER,
}

FACTORY_CODE_TO_KR: dict[str, str] = {f.f_code: f.label for f in FACTORY_MASTER}
FACTORY_KR_TO_CODE: dict[str, str] = {f.label: f.f_code for f in FACTORY_MASTER}
FACTORY_TO_WEATHER_STATION: dict[str, str] = {
    f.label: f.weather_station for f in FACTORY_MASTER
}
PRODUCTION_PARENT_CODE_BY_FACTORY: dict[str, str] = {
    f.label: f.production_parent or f.f_code for f in FACTORY_MASTER
}
PRODUCTION_DAILY_FACTORY_MAP: dict[str, str] = {
    "남양주1": "남양주",
    "남양주2": "남양주",
    "남양주": "남양주",
    "김해": "김해",
    "광주": "광주",
    "논산": "논산",
    "경산": "경산",
}
PRODUCTION_PARENT_FACTORY_CODES: frozenset[str] = frozenset(
    f.production_parent or f.f_code for f in FACTORY_MASTER
)

NAMYANGJU_PARENT_CODE = "F10"
NAMYANGJU_F10A_CODE = "F10A"
NAMYANGJU_F10B_CODE = "F10B"

FACTORY_COLORS: dict[str, str] = {
    "전사": "#3b82f6",
    "남양주": "#8b5cf6",
    **{f.label: f.color for f in FACTORY_MASTER},
}
DASHBOARD_FACTORY_COLORS: dict[str, str] = {
    key: FACTORY_COLORS[key] for key in ("전사", "남양주", "김해", "광주", "논산", "경산")
}

SHEET_TO_FACTORY_MAP: dict[str, str] = {
    **{f.label: f.label for f in FACTORY_MASTER},
    **FACTORY_CODE_TO_KR,
    "남양주": "남양주1",
    "F10": "남양주1",
    "F11": "남양주2",
}

FACTORY_QUERY_ORDER: tuple[tuple[str, str | None], ...] = (
    ("ALL", None),
    ("남양주1", "남양주1"),
    ("남양주2", "남양주2"),
    ("남양주", "남양주"),
    ("김해", "김해"),
    ("광주", "광주"),
    ("논산", "논산"),
    ("경산", "경산"),
)

YOY_FACTORY_DEFS: tuple[tuple[str, str | None], ...] = (
    ("전사", None),
    ("남양주1", "남양주1"),
    ("남양주2", "남양주2"),
    ("남양주", "남양주"),
    ("김해", "김해"),
    ("광주", "광주"),
    ("논산", "논산"),
    ("경산", "경산"),
)

ENERGY_UNIT_CALC_MAP: dict[str, str] = {
    "power_per_ton_kwh": "total_power_kwh",
    "fuel_per_ton_nm3": "fuel_nm3",
    "water_per_ton_ton": "water_ton",
    # 폐수 원단위(wastewater_per_ton_ton)는 폐기됨 — 대신 화면/메일에서
    # 폐수/용수 비(폐수량/용수량, 소수점 2자리)를 즉석 계산해 표시한다.
}


def expand_factory_members(factory: str) -> tuple[str, ...]:
    """Return physical energy factories represented by a label."""
    return AGGREGATE_FACTORY_MEMBERS.get(factory, (factory,))


def expand_factory_filter(factories: list[str] | tuple[str, ...] | None) -> list[str]:
    """Expand aggregate labels into physical factories for DB filters."""
    if not factories:
        return []
    expanded: list[str] = []
    for factory in factories:
        if factory in ("전체", "전사"):
            return list(FACTORY_PHYSICAL_DISPLAY_ORDER)
        expanded.extend(expand_factory_members(factory))
    return list(dict.fromkeys(expanded))


def filter_factory_frame(df: pd.DataFrame, db_code: str | None) -> pd.DataFrame:
    """Filter a DataFrame by physical or aggregate factory code."""
    if db_code is None:
        return df
    members = expand_factory_members(db_code)
    if len(members) > 1:
        return df[df["factory"].isin(members)]
    return df[df["factory"] == members[0]]


def recalc_unit_rates(
    df: pd.DataFrame,
    *,
    production_col: str = "mix_prod_kg",
    unit_calc_map: Mapping[str, str] = ENERGY_UNIT_CALC_MAP,
) -> pd.DataFrame:
    """Recalculate unit rates as SUM(usage) / SUM(production kg / 1000)."""
    if df.empty or production_col not in df.columns:
        return df

    denom_ton = pd.to_numeric(df[production_col], errors="coerce") / 1000.0
    for unit_col, usage_col in unit_calc_map.items():
        if usage_col not in df.columns:
            continue
        usage = pd.to_numeric(df[usage_col], errors="coerce")
        df[unit_col] = (usage / denom_ton).where(denom_ton > 0)
    return df
