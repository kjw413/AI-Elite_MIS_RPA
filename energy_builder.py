# RawDB_에너지.xlsx — 에너지 수집 데이터의 스키마 정의 · 적재 · 읽기
"""
MIS '원단위 실적입력(일단위)' 화면에서 수집한 에너지 실적을 `RawDB_에너지.xlsx` 에
적재한다. 이 파일이 **BEMS 웹앱이 그대로 읽는 유일한 산출물**이며 재가공 단계는 없다.

    [MIS 원단위 실적입력(일단위)]
      냉동전력/공압기/전력량/전력비/전력단가/연료량/연료비/연료단가/
      용수량/폐수량/원수COD/배출수COD
                     │  utility_daily_rpa (일일 --ym / 과거 --from~--to 겸용)
                     ▼
              RawDB_에너지.xlsx     행 = 일자, 열 = 항목 (tidy)
              시트 = 공장명          → BEMS 웹앱이 startup 에 직접 적재

행=일자 방향은 `DB_생산실적.xlsx` 의 `daily` 시트, `DB_재공품.xlsx` 와 통일된 형태다.

== 이력 ==
- 2026-07 수집 화면이 '유틸리티 일자별 사용량 추이' → '원단위 실적입력(일단위)' 로
  바뀌며 단가·비용·COD 가 추가됐다.
- 한동안 `DB_에너지.xlsx`(행=항목, 열=날짜 전치형)를 중간 산출물로 두고 `build_dataset()`
  으로 재가공했으나, BEMS 파서가 tidy 형태를 그대로 읽을 수 있어(전치 감지 분기를 타지
  않고 머리글 부분매칭만으로 컬럼이 잡힘) 2026-07-30 폐지했다.
- 믹스생산량·원단위 4개 항목(source='legacy')은 구 화면 수집분 보존용으로 열만 남아
  있다. BEMS 가 생산실적(`production_daily.actual_qty`)을 분모로 원단위를 매 조회마다
  재계산하므로(`production_actual_service.overlay_actual_production`) 값이 비어도 무해하다.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import openpyxl

from _common import atomic_save_workbook, sampled_db_path
from factories import FACTORY_PHYSICAL_DISPLAY_ORDER

log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# 경로
# ---------------------------------------------------------------------------
# 수집 산출물 = 웹앱 입력 파일 (단일). env 이름은 BEMS `v5_common.PATH_ENERGY_SOURCE`
# 와 동일한 `ENERGY_SOURCE_XLSX` 를 쓴다 — 한 변수로 양쪽이 같은 파일을 가리킨다.
DEFAULT_RAW_PATH = Path(sampled_db_path("RawDB_에너지.xlsx", "ENERGY_SOURCE_XLSX"))

FACTORY_SHEETS: tuple[str, ...] = FACTORY_PHYSICAL_DISPLAY_ORDER


# ---------------------------------------------------------------------------
# 스키마 — RawDB_에너지 열 정의의 단일 출처
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class EnergyField:
    key: str        # 코드 내부 식별자
    label: str      # RawDB_에너지 열 머리글 (엑셀에 그대로 쓰이는 문자열)
    source: str     # 'unit_input' = MIS 원단위 실적입력 화면에서 수집
                    # 'legacy'     = 2026-07 이전 구 화면 수집분. 더 이상 수집하지
                    #                않으며 과거 값 보존용으로 열/행만 유지한다
                    #                (BEMS validate_columns 가 필수 컬럼으로 검사).
                    #                원단위는 BEMS 가 생산실적 분모로 매 조회 재계산.


# 전 항목 정의. 라벨 문자열은 BEMS 파서가 **부분매칭**으로 컬럼을 찾는 키이므로
# (`excel_parser.kor_to_eng`, `daily_energy_sync_service._KOR_SUBSTR_MAP`)
# 바꿀 때는 양쪽 매핑을 함께 확인해야 한다.
FIELDS: tuple[EnergyField, ...] = (
    EnergyField("freezing_power_kwh",   "냉동전력량[kWh]",            "unit_input"),
    EnergyField("air_compressor_kwh",   "공압기[kWh]",                "unit_input"),
    EnergyField("total_power_kwh",      "전력량[kWh]",                "unit_input"),
    EnergyField("fuel_nm3",             "연료량[N㎥]",                "unit_input"),
    EnergyField("water_ton",            "용수량[ton]",                "unit_input"),
    EnergyField("wastewater_ton",       "폐수량[ton]",                "unit_input"),
    EnergyField("mix_prod_kg",          "믹스생산량[kg]",             "legacy"),
    EnergyField("power_per_ton_kwh",    "전력원단위[kWh/mix-ton]",    "legacy"),
    EnergyField("fuel_per_ton_nm3",     "연료원단위[N㎥/mix-ton]",    "legacy"),
    EnergyField("water_per_ton_ton",    "용수원단위[ton/mix-ton]",    "legacy"),
    EnergyField("power_cost_krw",       "전력비[원]",                 "unit_input"),
    EnergyField("power_price_krw_kwh",  "전력단가[원/kWh]",           "unit_input"),
    EnergyField("fuel_cost_krw",        "연료비[원]",                 "unit_input"),
    EnergyField("fuel_price_krw_nm3",   "연료단가[원/N㎥]",           "unit_input"),
    EnergyField("influent_cod_ppm",     "원수COD[ppm]",               "unit_input"),
    EnergyField("effluent_cod_ppm",     "배출수COD[ppm]",             "unit_input"),
)

FIELD_BY_KEY: dict[str, EnergyField] = {f.key: f for f in FIELDS}

# RawDB_에너지 열 순서 — MIS '원단위 실적입력' 화면 열 순서를 그대로 따르고,
# legacy 항목(믹스/원단위)을 뒤에 붙인다. 사람이 원본과 눈으로 대조하기 쉽게.
RAW_COLUMN_KEYS: tuple[str, ...] = (
    "freezing_power_kwh", "air_compressor_kwh",
    "total_power_kwh", "power_cost_krw", "power_price_krw_kwh",
    "fuel_nm3", "fuel_cost_krw", "fuel_price_krw_nm3",
    "water_ton", "wastewater_ton",
    "influent_cod_ppm", "effluent_cod_ppm",
    "mix_prod_kg",
    "power_per_ton_kwh", "fuel_per_ton_nm3", "water_per_ton_ton",
)
RAW_DATE_HEADER = "날짜"


def _date_key(value) -> str | None:
    """셀 값을 'YY-MM-DD' 비교 키로 정규화한다 (date/datetime/문자열 혼재 대응)."""
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return value.strftime("%y-%m-%d")
    text = str(value).strip().lstrip("'")
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%y-%m-%d", "%Y/%m/%d", "%y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).strftime("%y-%m-%d")
        except ValueError:
            continue
    return text


# ---------------------------------------------------------------------------
# RawDB_에너지 (행=일자, 열=항목) 읽기/쓰기
# ---------------------------------------------------------------------------
def _ensure_raw_sheet(wb, sheet_name: str):
    """RawDB 시트를 확보하고 머리글 행을 보정한다."""
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)

    headers = [RAW_DATE_HEADER] + [FIELD_BY_KEY[k].label for k in RAW_COLUMN_KEYS]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        if cell.value != header:
            cell.value = header
    return ws


def write_raw(records: dict[str, dict[date, dict[str, float]]],
              raw_path: Path | None = None) -> Path:
    """수집 결과를 RawDB_에너지.xlsx 에 upsert 한다 (날짜 1건 = 1행).

    Args:
        records: { 시트명: { date: { field_key: value } } }
                 값이 None 인 항목은 기록하지 않는다 (기존 값 보존).
    """
    raw_path = Path(raw_path or DEFAULT_RAW_PATH)
    raw_path.parent.mkdir(parents=True, exist_ok=True)

    if raw_path.exists():
        wb = openpyxl.load_workbook(raw_path)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    total_new = total_updated = 0
    for sheet_name, by_date in records.items():
        if not by_date:
            continue
        ws = _ensure_raw_sheet(wb, sheet_name)

        # 기존 날짜 행 인덱싱
        row_of_date: dict[str, int] = {}
        last_row = 1
        for row_idx in range(2, ws.max_row + 1):
            key = _date_key(ws.cell(row=row_idx, column=1).value)
            if key:
                row_of_date[key] = row_idx
                last_row = row_idx

        new_cnt = upd_cnt = 0
        for day, values in sorted(by_date.items()):
            key = _date_key(day)
            row_idx = row_of_date.get(key)
            if row_idx is None:
                last_row += 1
                row_idx = last_row
                date_cell = ws.cell(row=row_idx, column=1, value=day)
                date_cell.number_format = "YYYY-MM-DD"
                row_of_date[key] = row_idx
                new_cnt += 1
            else:
                upd_cnt += 1
            for col_idx, field_key in enumerate(RAW_COLUMN_KEYS, start=2):
                value = values.get(field_key)
                if value is not None:
                    ws.cell(row=row_idx, column=col_idx, value=value)

        total_new += new_cnt
        total_updated += upd_cnt
        log.info(f"  [RawDB/{sheet_name}] 신규 {new_cnt}일 / 갱신 {upd_cnt}일")

    # 시트 순서를 공장 순서로 정렬 (수집 순서와 무관하게 항상 동일)
    order = {name: i for i, name in enumerate(FACTORY_SHEETS)}
    wb._sheets.sort(key=lambda s: order.get(s.title, len(order)))

    atomic_save_workbook(wb, str(raw_path))
    wb.close()
    log.info(f"RawDB 저장: {raw_path}  (신규 {total_new} / 갱신 {total_updated})")
    return raw_path


def read_raw(raw_path: Path | None = None) -> dict[str, dict[date, dict[str, float]]]:
    """RawDB_에너지.xlsx 를 { 시트명: { date: { field_key: value } } } 로 읽는다."""
    raw_path = Path(raw_path or DEFAULT_RAW_PATH)
    if not raw_path.exists():
        raise FileNotFoundError(f"RawDB 파일이 없습니다: {raw_path}")

    wb = openpyxl.load_workbook(raw_path, read_only=True, data_only=True)
    try:
        out: dict[str, dict[date, dict[str, float]]] = {}
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            rows = ws.iter_rows(values_only=True)
            try:
                header = next(rows)
            except StopIteration:
                continue

            # 머리글 → field_key 매핑 (열 순서가 바뀌어도 라벨로 찾는다)
            label_to_key = {FIELD_BY_KEY[k].label: k for k in RAW_COLUMN_KEYS}
            col_keys: list[str | None] = []
            for cell in header:
                col_keys.append(label_to_key.get(str(cell).strip()) if cell else None)

            by_date: dict[date, dict[str, float]] = {}
            for row in rows:
                if not row or row[0] is None:
                    continue
                day = row[0]
                if isinstance(day, datetime):
                    day = day.date()
                if not isinstance(day, date):
                    continue
                values = {
                    key: row[i]
                    for i, key in enumerate(col_keys)
                    if key and i < len(row) and row[i] is not None
                }
                if values:
                    by_date[day] = values
            if by_date:
                out[sheet_name] = by_date
        return out
    finally:
        wb.close()
