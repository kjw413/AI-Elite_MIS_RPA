# RawDB_에너지.xlsx (MIS 원본 수집) → DB_에너지.xlsx (웹앱이 읽는 전치형) 재가공
"""
에너지(유틸리티) 수집 데이터의 스키마 정의 · 원본 적재 · 재가공을 담당한다.

== 배경 ==
2026-07 부터 MIS 수집 화면이 '유틸리티 일자별 사용량 추이' → '원단위 실적입력(일단위)'
로 바뀌었다. 신규 화면은 단가·비용·COD 를 함께 제공하지만 믹스생산량·원단위는 없어
기존 화면을 병행 수집한다. 이에 따라 파일이 2단으로 분리됐다.

    [MIS 원단위 실적입력(일단위)]  ─┐
      냉동전력/공압기/전력량/전력비/전력단가/연료량/연료비/연료단가/
      용수량/폐수량/원수COD/배출수COD                                │
                                                                     ├─►  RawDB_에너지.xlsx
    [MIS 유틸리티 일자별 사용량 추이] ─┘                              │    (행=일자, 열=항목)
      믹스생산량/전력원단위/연료원단위/용수원단위                     │     ※ 화면 그대로의 원본
                                                                     ▼
                                                       build_dataset() 재가공
                                                                     ▼
                                                            DB_에너지.xlsx
                                                        (행=항목, 열=날짜)
                                                     → BEMS 웹앱이 startup 에 적재

기존 `RawDB_에너지.xlsx` 는 전치형(행=항목, 열=날짜) 이었고 이제 그 역할이
`DB_에너지.xlsx` 로 넘어간다. `migrate_legacy_rawdb()` 가 1회 자동 이관한다.

Usage:
  python build_energy_dataset.py                 # RawDB_에너지 → DB_에너지 재가공
"""

from __future__ import annotations

import logging
import shutil
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
# RawDB = MIS 화면 원본 수집 (행=일자), DB = 재가공 산출물 (행=항목) — 웹앱 입력
DEFAULT_RAW_PATH = Path(sampled_db_path("RawDB_에너지.xlsx", "ENERGY_RAW_XLSX"))
DEFAULT_OUTPUT_PATH = Path(sampled_db_path("DB_에너지.xlsx", "ENERGY_SOURCE_XLSX"))

FACTORY_SHEETS: tuple[str, ...] = FACTORY_PHYSICAL_DISPLAY_ORDER


# ---------------------------------------------------------------------------
# 스키마 — RawDB 열 / DB 행의 단일 정의처
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class EnergyField:
    key: str        # 코드 내부 식별자
    label: str      # RawDB 열 머리글 = DB_에너지 A열 항목명 (동일 문자열)
    source: str     # 수집 화면: 'unit_input' | 'usage_trend'


# DB_에너지 시트의 행 순서 (row 2 부터). 2~11행은 구 RawDB_에너지와 **동일 순서·동일
# 라벨**이어야 한다 — 과거 데이터와 BEMS 파서(부분매칭)가 그대로 붙는다.
# 12~17행이 신규 화면에서 추가된 항목.
FIELDS: tuple[EnergyField, ...] = (
    EnergyField("freezing_power_kwh",   "냉동전력량[kWh]",            "unit_input"),
    EnergyField("air_compressor_kwh",   "공압기[kWh]",                "unit_input"),
    EnergyField("total_power_kwh",      "전력량[kWh]",                "unit_input"),
    EnergyField("fuel_nm3",             "연료량[N㎥]",                "unit_input"),
    EnergyField("water_ton",            "용수량[ton]",                "unit_input"),
    EnergyField("wastewater_ton",       "폐수량[ton]",                "unit_input"),
    EnergyField("mix_prod_kg",          "믹스생산량[kg]",             "usage_trend"),
    EnergyField("power_per_ton_kwh",    "전력원단위[kWh/mix-ton]",    "usage_trend"),
    EnergyField("fuel_per_ton_nm3",     "연료원단위[N㎥/mix-ton]",    "usage_trend"),
    EnergyField("water_per_ton_ton",    "용수원단위[ton/mix-ton]",    "usage_trend"),
    EnergyField("power_cost_krw",       "전력비[원]",                 "unit_input"),
    EnergyField("power_price_krw_kwh",  "전력단가[원/kWh]",           "unit_input"),
    EnergyField("fuel_cost_krw",        "연료비[원]",                 "unit_input"),
    EnergyField("fuel_price_krw_nm3",   "연료단가[원/N㎥]",           "unit_input"),
    EnergyField("influent_cod_ppm",     "원수COD[ppm]",               "unit_input"),
    EnergyField("effluent_cod_ppm",     "배출수COD[ppm]",             "unit_input"),
)

FIELD_BY_KEY: dict[str, EnergyField] = {f.key: f for f in FIELDS}

# RawDB_에너지 열 순서 — MIS '원단위 실적입력' 화면 열 순서를 그대로 따르고,
# 기존 화면 보완분(믹스/원단위)을 뒤에 붙인다. 사람이 원본과 눈으로 대조하기 쉽게.
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

# DB_에너지 A열: row 1 = '날짜', row 2.. = FIELDS 라벨
DB_DATE_HEADER = "날짜"
DB_FIRST_DATA_ROW = 2

# 구 RawDB_에너지(전치형) 판별용 — A1='날짜', A2 가 첫 항목 라벨
_LEGACY_A1 = DB_DATE_HEADER
_LEGACY_A2_PREFIX = "냉동전력"


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
# 구 파일 이관 (RawDB_에너지[전치형] → DB_에너지)
# ---------------------------------------------------------------------------
def is_legacy_transposed(path: Path) -> bool:
    """파일이 구 형식(행=항목, 열=날짜)인지 판별한다."""
    if not Path(path).exists():
        return False
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
    except Exception:
        return False
    try:
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            a1 = str(ws.cell(row=1, column=1).value or "").strip()
            a2 = str(ws.cell(row=2, column=1).value or "").strip()
            if a1 == _LEGACY_A1 and a2.startswith(_LEGACY_A2_PREFIX):
                return True
        return False
    finally:
        wb.close()


def migrate_legacy_rawdb(raw_path: Path | None = None,
                         db_path: Path | None = None) -> bool:
    """구 `RawDB_에너지.xlsx`(전치형)를 `DB_에너지.xlsx` 로 1회 이관한다.

    - DB_에너지.xlsx 가 이미 있으면 아무것도 하지 않는다.
    - 구 파일은 backup/ 으로 옮겨 원본 경로를 신규 형식(행=일자)에 내준다.
      (삭제하지 않으므로 언제든 복구 가능)

    Returns: 이관을 수행했으면 True.
    """
    raw_path = Path(raw_path or DEFAULT_RAW_PATH)
    db_path = Path(db_path or DEFAULT_OUTPUT_PATH)

    if db_path.exists():
        return False
    if not is_legacy_transposed(raw_path):
        return False

    db_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(raw_path, db_path)

    backup_dir = raw_path.parent / "backup"
    backup_dir.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    archived = backup_dir / f"{raw_path.stem}_legacy_{stamp}{raw_path.suffix}"
    shutil.move(str(raw_path), str(archived))

    log.warning("구 형식 RawDB 이관 완료 — 화면 변경(원단위 실적입력)에 따른 1회 작업")
    log.warning(f"  {raw_path.name}(전치형) → {db_path.name}  (웹앱 입력 파일)")
    log.warning(f"  원본 보관: {archived}")
    log.warning(f"  ※ BEMS .env 의 ENERGY_SOURCE_XLSX 를 {db_path} 로 지정하세요.")
    return True


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


# ---------------------------------------------------------------------------
# DB_에너지 (행=항목, 열=날짜) 재가공
# ---------------------------------------------------------------------------
def _ensure_db_sheet(wb, sheet_name: str):
    """DB 시트를 확보하고 A열 항목 라벨을 보정한다.

    기존 시트의 A2~A11 라벨은 이미 채워져 있으므로 건드리지 않고,
    신규 항목(A12~) 처럼 비어 있는 칸만 채운다.
    """
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)

    if not str(ws.cell(row=1, column=1).value or "").strip():
        ws.cell(row=1, column=1, value=DB_DATE_HEADER)

    for offset, field in enumerate(FIELDS):
        row_idx = DB_FIRST_DATA_ROW + offset
        cell = ws.cell(row=row_idx, column=1)
        current = str(cell.value or "").strip()
        if not current:
            cell.value = field.label
        elif current != field.label:
            log.warning(
                f"  [{sheet_name}] A{row_idx} 항목명 불일치: "
                f"'{current}' (기대: '{field.label}') — 기존 라벨 유지"
            )
    return ws


def build_dataset(raw_path: Path | None = None,
                  output_path: Path | None = None):
    """RawDB_에너지.xlsx → DB_에너지.xlsx 재가공 (날짜 열 upsert).

    기존 DB_에너지.xlsx 의 과거 열은 보존하고, RawDB 에 있는 날짜만 덮어쓰거나
    맨 뒤에 새 열로 추가한다. 값이 없는 항목은 기존 셀을 건드리지 않는다.

    Returns: (통계 dict, 출력 경로)
    """
    raw_path = Path(raw_path or DEFAULT_RAW_PATH)
    output_path = Path(output_path or DEFAULT_OUTPUT_PATH)

    records = read_raw(raw_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        wb = openpyxl.load_workbook(output_path)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    stats: dict[str, dict[str, int]] = {}
    for sheet_name, by_date in records.items():
        if not by_date:
            continue
        ws = _ensure_db_sheet(wb, sheet_name)

        # 기존 날짜 열 인덱싱 (row 1 을 왼쪽부터, 빈 칸에서 중단)
        col_of_date: dict[str, int] = {}
        last_col = 1
        for col_idx in range(2, ws.max_column + 2):
            key = _date_key(ws.cell(row=1, column=col_idx).value)
            if key is None:
                break
            col_of_date[key] = col_idx
            last_col = col_idx

        appended = overwritten = 0
        for day, values in sorted(by_date.items()):
            key = _date_key(day)
            col_idx = col_of_date.get(key)
            if col_idx is None:
                last_col += 1
                col_idx = last_col
                header_cell = ws.cell(row=1, column=col_idx, value=day)
                header_cell.number_format = "YY-MM-DD"
                col_of_date[key] = col_idx
                appended += 1
            else:
                overwritten += 1
            for offset, field in enumerate(FIELDS):
                value = values.get(field.key)
                if value is not None:
                    ws.cell(row=DB_FIRST_DATA_ROW + offset,
                            column=col_idx, value=value)

        stats[sheet_name] = {"appended": appended, "overwritten": overwritten}
        log.info(f"  [DB/{sheet_name}] 신규 {appended}열 / 덮어쓰기 {overwritten}열")

    order = {name: i for i, name in enumerate(FACTORY_SHEETS)}
    wb._sheets.sort(key=lambda s: order.get(s.title, len(order)))

    atomic_save_workbook(wb, str(output_path))
    wb.close()
    log.info(f"DB 저장: {output_path}")
    return stats, output_path
