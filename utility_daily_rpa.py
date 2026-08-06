# MIS 에너지(유틸리티) 일일 실적 자동 샘플링 RPA (pywinauto + openpyxl 기반)
"""
사내 MIS '(신)종합정보' 시스템에서 에너지 일일 실적을 사업장별로 자동 조회 →
클립보드 복사 결과를 먼저 메모리에 수집한 뒤, 별도 가공 단계에서
RawDB_에너지.xlsx 에 적재하는 RPA 프로그램. 이 파일을 BEMS 웹앱이 그대로 읽는다.

== 수집 화면 ==
'원단위 실적입력(일단위)' [unit_input] 한 화면만 사용한다 — 행=일자, 열=항목.

  일자 | 냉동전력 | 공압기 | 전력사용량 | 전력비 | 전력단가 |
         연료사용량 | 연료비 | 연료단가 | 용수사용량 | 폐수발생량 |
         원수COD | 배출수COD

2026-07 이전에는 '유틸리티 일자별 사용량 추이' 화면을 쓰다가 이 화면으로 옮겼다.
신규 화면에 없는 믹스생산량은 방금 수집한 `RawDB_생산실적.xlsx`의 공장별 actual_qty
합계로 해당 월 전체를 다시 동기화한다. 생산량이 없는 공장·일자는 빈 원단위를 저장하지
않고 적재를 중단한다. 과거 기간은 `DB_생산실적.xlsx` daily 값으로 보완한다.
냉동·공압·전력·연료·용수 원단위는 Python에서 계산하지 않고
RawDB_에너지.xlsx 수식으로 관리하며,
RPA 적재 시 빈 수식을 자동 채운 뒤 Excel에서 전체 재계산·저장한다.

업무 절차:
  1. MIS 앱 연결 (pywinauto UIA backend)
  2. 화면 진입 → 기준년월 설정 → 사업장 순회(F1A→F1B→F20→F30→F40→F50)
     조회 → 그리드 복사 → 클립보드 파싱
  3. RawDB_에너지.xlsx 적재 (행=일자, 열=항목)

일일 수집과 과거 데이터 수집은 화면·파싱·적재 경로가 완전히 같고 도는 달 수만 다르다.
모든 월의 MIS 수집을 먼저 끝낸 뒤 가공 단계에서 월 순서대로 적재한다.

Usage:
  python utility_daily_rpa.py                          # D-1 기준월 + 직전 누락 자동 복구
  python utility_daily_rpa.py --ym 2026-07             # 특정 월 1개
  python utility_daily_rpa.py --dry-run                # MIS 조회만, 엑셀 미기록
  python utility_daily_rpa.py --factories 논산,경산       # 특정 사업장만
  python utility_daily_rpa.py --from 2024-01           # 과거 수집: ~ D-1 월까지
  python utility_daily_rpa.py --from 2024-01 --to 2026-06
  python utility_daily_rpa.py --from 2024-01 --resume  # 이미 받은 달은 건너뜀
"""

import sys
import time
import os
import json
import shutil
import hashlib
import logging
import argparse
import re
from collections import OrderedDict
from datetime import datetime, date, timedelta
from pathlib import Path

import pandas as pd
from pywinauto import Application
from pywinauto.keyboard import send_keys
from pywinauto.timings import Timings

# pywinauto 내부 click/keys 대기 시간 단축 (MIS는 즉시 반응한다는 가정)
Timings.after_clickinput_wait = 0.01
Timings.after_setfocus_wait = 0.01
Timings.after_sendkeys_key_wait = 0.001

# 프로젝트 루트(app/tools/...) import 보장 — _common.fast_click 사용을 위함
_PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

import energy_builder  # noqa: E402
from _common import (  # noqa: E402
    fast_click,
    find_mis_window,
    get_clipboard_sequence,
    month_range,
    parse_year_month,
    wait_for_clipboard_change,
)

# ---------------------------------------------------------------------------
# 로깅 설정
# ---------------------------------------------------------------------------
LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)
log = logging.getLogger(__name__)


def _setup_logging(prefix: str = "rpa") -> None:
    if logging.getLogger().handlers:
        return
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(
                LOG_DIR / f"{prefix}_{datetime.now():%Y%m%d_%H%M%S}.log",
                encoding="utf-8",
            ),
        ],
    )

# ---------------------------------------------------------------------------
# 상수 정의
# ---------------------------------------------------------------------------
# 사업장 코드 → Excel 시트명 매핑 (순회 순서 보장)
FACTORY_SHEET_MAP = OrderedDict([
    ("F1A", "남양주1"),
    ("F1B", "남양주2"),
    ("F20", "김해"),
    ("F30", "광주"),
    ("F40", "논산"),
    ("F50", "경산"),
])


def resolve_org_codes(spec: str | None) -> list[str]:
    """`--factories` 값을 사업장 코드 리스트로 변환한다.

    코드(F40)와 한글 공장명(논산)을 모두 받는다. 코드만 받으면 오타가
    조용히 다른 공장을 수집하게 되므로 알 수 없는 값은 즉시 중단한다.

    >>> resolve_org_codes("논산,F50")
    ['F40', 'F50']
    """
    if not spec:
        return list(FACTORY_SHEET_MAP)

    by_name = {name: code for code, name in FACTORY_SHEET_MAP.items()}
    selected: set[str] = set()
    unknown: list[str] = []
    for token in spec.split(","):
        token = token.strip()
        if not token:
            continue
        if token.upper() in FACTORY_SHEET_MAP:
            selected.add(token.upper())
        elif token in by_name:
            selected.add(by_name[token])
        else:
            unknown.append(token)

    if unknown:
        raise SystemExit(
            f"알 수 없는 사업장: {unknown}\n"
            f"  코드 또는 공장명을 쓰세요 — "
            + ", ".join(f"{c}={n}" for c, n in FACTORY_SHEET_MAP.items())
        )
    if not selected:
        raise SystemExit("--factories 에 사업장이 하나도 지정되지 않았습니다.")

    # FACTORY_SHEET_MAP 순서 유지 (입력 순서와 무관하게 항상 동일한 순회 순서)
    return [code for code in FACTORY_SHEET_MAP if code in selected]


# ---------------------------------------------------------------------------
# 기준년월 범위 (과거 데이터 수집)
# ---------------------------------------------------------------------------
# 기간 인자 문법(YYYY-MM)은 가공 CLI(build_energy_dataset.py)와 공유한다.
# 공장 1곳 1개월 조회+복사에 걸리는 대략적 소요 (진행 예상시간 안내용)
_SECONDS_PER_QUERY = 3.0


UTILITY_INPUT_FIELD_KEYS = {
    field.key for field in energy_builder.FIELDS if field.source == "unit_input"
}


def load_utility_collected_dates(
    raw_path: str | Path,
    sheet_names: list[str],
) -> dict[str, set[date]]:
    """사업장별로 MIS 유틸리티 원본 값이 실제 적재된 날짜를 읽는다.

    믹스생산량이나 원단위 수식만 있는 행은 수집 완료로 보지 않는다. 반면
    유틸리티 원본 값 0은 정상적인 실적이므로 완료 날짜에 포함한다.
    """
    path = Path(raw_path)
    if not path.exists():
        return {}

    try:
        existing = energy_builder.read_raw(path)
    except Exception as exc:
        log.warning(f"유틸리티 수집 이력 확인 실패({path}): {exc}")
        return {}

    result: dict[str, set[date]] = {}
    for sheet_name in sheet_names:
        by_date = existing.get(sheet_name) or {}
        result[sheet_name] = {
            day
            for day, values in by_date.items()
            if any(key in values for key in UTILITY_INPUT_FIELD_KEYS)
        }
    return result


def plan_utility_collection_months(
    end_date: date,
    collected_dates_by_sheet: dict[str, set[date]],
) -> list[str]:
    """직전 수집일 누락 여부에 따라 이전 월과 현재 월 수집 순서를 만든다."""
    current_month = end_date.strftime("%Y-%m")
    if not collected_dates_by_sheet or not any(collected_dates_by_sheet.values()):
        return [current_month]

    current_month_start = end_date.replace(day=1)
    previous_day = current_month_start - timedelta(days=1)
    if all(
        previous_day in dates
        for dates in collected_dates_by_sheet.values()
    ):
        return [current_month]

    return [previous_day.strftime("%Y-%m"), current_month]


def _already_collected(year_month: str, sheet_names: list[str],
                       existing: dict) -> bool:
    """--resume 판정: 해당 월에 요청한 모든 사업장의 행이 이미 있는지."""
    year, month = parse_year_month(year_month)
    for sheet_name in sheet_names:
        by_date = existing.get(sheet_name) or {}
        if not any(d.year == year and d.month == month for d in by_date):
            return False
    return True


def drop_collected_months(months: list[str], org_codes: list[str]) -> list[str]:
    """이미 적재된 달을 제외한다 (--resume). 파일이 없으면 전체를 반환."""
    raw_path = Path(energy_builder.DEFAULT_RAW_PATH)
    if not raw_path.exists():
        return months
    existing = energy_builder.read_raw(raw_path)
    sheet_names = [FACTORY_SHEET_MAP[c] for c in org_codes]
    remaining = [m for m in months
                 if not _already_collected(m, sheet_names, existing)]
    skipped = len(months) - len(remaining)
    if skipped:
        log.info(f"--resume: 이미 수집된 {skipped}개월 건너뜀")
    return remaining


def confirm_plan(months: list[str], org_codes: list[str],
                 assume_yes: bool) -> None:
    """여러 달을 도는 실행은 소요가 길어 사용자 확인을 받는다."""
    sheet_names = [FACTORY_SHEET_MAP[c] for c in org_codes]
    queries = len(months) * len(org_codes)
    minutes = queries * _SECONDS_PER_QUERY / 60

    print()
    print("=" * 66)
    print("  MIS 에너지 실적 수집 — '원단위 실적입력(일단위)' 화면")
    print("=" * 66)
    print(f"  기간      : {months[0]} ~ {months[-1]}  ({len(months)}개월)")
    print(f"  사업장    : {', '.join(sheet_names)}")
    print(f"  조회 횟수 : {queries}회  (예상 {minutes:.0f}분)")
    print(f"  적재 대상 : {energy_builder.DEFAULT_RAW_PATH}  (웹앱이 읽는 파일)")
    print()
    print("  ※ 해당 기간의 수집 항목은 화면 값으로 '덮어쓰기' 됩니다.")
    print("     생산실적을 동기화하고, 빈 원단위 수식은 자동으로 채웁니다.")
    print("  ※ 실행 중 마우스/키보드를 사용하지 마세요 (좌표 클릭 기반).")
    print("=" * 66)

    if assume_yes:
        return
    try:
        answer = input("  진행할까요? [y/N] ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        # 파이프/리다이렉트 등 대화형 입력이 불가능한 환경 — 임의 진행하지 않는다
        raise SystemExit(
            "\n확인 입력을 받을 수 없습니다. 확인 없이 실행하려면 --yes 를 주세요."
        )
    if answer not in ("y", "yes"):
        raise SystemExit("사용자가 취소했습니다.")


# 화면 식별자 (utility_coords.json 의 coords 키와 동일)
SCREEN_UNIT_INPUT = "unit_input"

# ── '원단위 실적입력(일단위)' 그리드 열 순서 ──
# 0번은 일자(=일 번호). None 은 무시할 열.
# 주의: 화면 머리글은 '사용량 → 단가 → 비용' 순으로 보이지만 클립보드로 나오는 실제
#       값은 '사용량 → 비용 → 단가' 순이다(2026-07-29 실측: 29700 × 203.8 = 6,052,906).
#       머리글/데이터 어느 쪽이 바뀌어도 안전하도록 아래 _resolve_cost_price_pairs()
#       가 산술(사용량 × 단가 = 비용)로 두 열을 재판별한다.
UNIT_INPUT_COLUMN_KEYS = (
    None,                       # 0  일자
    "freezing_power_kwh",       # 1  냉동전력(kWh)
    "air_compressor_kwh",       # 2  공압기(kWh)
    "total_power_kwh",          # 3  사용량(kWh)
    "power_cost_krw",           # 4  전력비(원)   ← 단가와 자리 뒤바뀔 수 있음
    "power_price_krw_kwh",      # 5  단가(원/kWh)
    "fuel_nm3",                 # 6  연료사용량(N㎥)
    "fuel_cost_krw",            # 7  연료비(원)   ← 단가와 자리 뒤바뀔 수 있음
    "fuel_price_krw_nm3",       # 8  단가(원/N㎥)
    "water_ton",                # 9  용수사용량(ton)
    "wastewater_ton",           # 10 폐수발생량(ton)
    "influent_cod_ppm",         # 11 원수COD(PPM)
    "effluent_cod_ppm",         # 12 배출수COD(PPM)
)

# (사용량, 비용, 단가) 3열 묶음 — 사용량 × 단가 ≒ 비용 관계로 비용/단가를 판별한다.
_COST_PRICE_TRIPLETS = (
    ("total_power_kwh", "power_cost_krw", "power_price_krw_kwh"),
    ("fuel_nm3", "fuel_cost_krw", "fuel_price_krw_nm3"),
)
# 단가는 화면에서 소수 2자리로 반올림되어 표시되므로 오차를 넉넉히 허용한다.
_COST_PRICE_TOLERANCE = 0.02

# 대기 시간 기본값 (utility_coords.json의 "wait" 값으로 덮어씌워짐)
WAIT_SHORT = 0.05        # 클릭/타이핑 후 미세 대기 (MIS는 즉시 반응)
WAIT_MEDIUM = 0.2        # 클립보드 fill / 필드 클릭 후 약간 더 긴 대기
WAIT_DROPDOWN = 0.01     # 드롭다운 펼침 후 항목 클릭 전 대기
WAIT_SCREEN_LOAD = 1.0   # 사이드바 메뉴 클릭 → MIS 화면 전환 로딩
WAIT_QUERY_LOAD = 1.0    # 조회 버튼 → 그리드 데이터 로딩
WAIT_COPY_CONFIRM = 0.4  # 복사 버튼 클릭 후 확인 팝업이 포커스를 받을 때까지 대기


# ---------------------------------------------------------------------------
# 클립보드 헬퍼
# ---------------------------------------------------------------------------
def get_clipboard_text() -> str:
    """Windows 클립보드에서 MIS 데이터를 읽어온다 (win32clipboard)."""
    import win32clipboard

    # MIS가 사용하는 'Csv' 커스텀 클립보드 포맷 ID
    csv_fmt = win32clipboard.RegisterClipboardFormat("Csv")

    for attempt in range(3):
        try:
            win32clipboard.OpenClipboard()
            try:
                # 1차: Csv 포맷 시도
                try:
                    data = win32clipboard.GetClipboardData(csv_fmt)
                    if data:
                        # bytes인 경우 디코딩
                        if isinstance(data, bytes):
                            for enc in ("utf-8", "euc-kr", "cp949"):
                                try:
                                    text = data.decode(enc).rstrip("\x00")
                                    if text.strip():
                                        log.info(f"  클립보드 읽기 성공 (Csv/{enc}, {len(text)}자)")
                                        return text
                                except UnicodeDecodeError:
                                    continue
                        elif isinstance(data, str) and data.strip():
                            log.info(f"  클립보드 읽기 성공 (Csv/str, {len(data)}자)")
                            return data
                except Exception:
                    pass

                # 2차: 표준 텍스트 포맷 시도
                try:
                    text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                    if text and text.strip():
                        log.info(f"  클립보드 읽기 성공 (UNICODETEXT, {len(text)}자)")
                        return text
                except Exception:
                    pass

                try:
                    text = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
                    if text:
                        decoded = text.decode("utf-8", errors="replace") if isinstance(text, bytes) else text
                        if decoded.strip():
                            log.info(f"  클립보드 읽기 성공 (TEXT, {len(decoded)}자)")
                            return decoded
                except Exception:
                    pass

            finally:
                win32clipboard.CloseClipboard()

        except Exception as e:
            log.warning(f"  클립보드 시도 {attempt + 1}/3 실패: {e}")

        time.sleep(0.5)

    log.warning("  클립보드 읽기 실패 (3회 시도)")
    return ""


# ---------------------------------------------------------------------------
# 클립보드 파싱 공통
# ---------------------------------------------------------------------------
class EmptyGridError(ValueError):
    """그리드에 일자 행이 없다 — 대개 해당 월/공장에 실적이 없는 정상 상황.

    과거 월 백필처럼 수백 회 순회할 때 '실적 없음'과 '진짜 오류'를 구분하기 위해
    별도 예외로 둔다 (전자는 traceback 없이 경고만 남긴다).
    """


def _read_clipboard_rows(raw_text: str) -> list[list[str]]:
    """MIS 클립보드 텍스트(CSV/TSV)를 셀 2차원 리스트로 파싱한다."""
    import csv
    import io

    raw_text = raw_text.strip()
    if not raw_text:
        raise ValueError("클립보드 데이터가 비어 있습니다.")

    log.info(f"  클립보드 데이터 미리보기: {repr(raw_text[:300])}")

    # 구분자 감지 (탭 vs 쉼표)
    first_line = raw_text.split("\n")[0]
    sep = "\t" if ("\t" in first_line and first_line.count("\t") >= 2) else ","
    log.info(f"  구분자 감지: {'TAB' if sep == chr(9) else 'COMMA'}")

    # csv.reader 사용 (인용부호 내 쉼표 = 천단위 구분기호 정상 처리)
    rows = list(csv.reader(io.StringIO(raw_text), delimiter=sep))
    if not rows:
        raise ValueError("파싱된 행이 없습니다.")
    return rows


def _to_number(cell) -> float | int | None:
    """MIS 셀 문자열을 숫자로 변환한다. 빈칸(' ')·비수치는 None."""
    if cell is None:
        return None
    text = str(cell).strip().replace(",", "").replace('"', "")
    if not text or text == "-":
        return None
    try:
        return float(text) if ("." in text or "e" in text.lower()) else int(text)
    except ValueError:
        return None


def _resolve_cost_price_pairs(by_date: dict) -> None:
    """(사용량, 비용, 단가) 3열의 비용/단가 위치를 산술로 재판별해 교정한다.

    `사용량 × 단가 = 비용` 관계가 성립하는 조합을 전체 행에서 다수결로 고른다.
    MIS 화면 머리글 순서('사용량→단가→비용')와 실제 클립보드 값 순서
    ('사용량→비용→단가')가 어긋나 있어, 어느 쪽이 오더라도 올바르게 적재하기 위함.
    """
    for usage_key, cost_key, price_key in _COST_PRICE_TRIPLETS:
        as_is = swapped = 0
        for values in by_date.values():
            usage = values.get(usage_key)
            a = values.get(cost_key)
            b = values.get(price_key)
            if not usage or a is None or b is None:
                continue
            # as_is : a=비용, b=단가  →  usage × b ≒ a
            if abs(usage * b - a) <= _COST_PRICE_TOLERANCE * max(abs(a), 1.0):
                as_is += 1
            # swapped: b=비용, a=단가  →  usage × a ≒ b
            if abs(usage * a - b) <= _COST_PRICE_TOLERANCE * max(abs(b), 1.0):
                swapped += 1

        if swapped > as_is:
            for values in by_date.values():
                a = values.pop(cost_key, None)
                b = values.pop(price_key, None)
                if b is not None:
                    values[cost_key] = b
                if a is not None:
                    values[price_key] = a
            log.warning(
                f"  비용/단가 열 위치 교정: {cost_key} ↔ {price_key} "
                f"(일치 {swapped}행 vs 기본배치 {as_is}행)"
            )
        elif as_is == 0 and swapped == 0:
            log.warning(
                f"  비용/단가 검증 불가({usage_key}) — 기본 열 순서를 그대로 사용합니다."
            )


# ---------------------------------------------------------------------------
# 화면 1) '원단위 실적입력(일단위)' 클립보드 파싱
# ---------------------------------------------------------------------------
def parse_unit_input_clipboard(raw_text: str, year_month: str) -> dict:
    """신규 화면 클립보드(행=일자)를 { date: { field_key: value } } 로 파싱한다.

    - 1열이 1~31 의 일 번호인 행만 취한다 (머리글·'TOTAL'·빈 행 자동 제외)
    - 값이 없는 미래 일자(전 항목 공백)는 건너뛴다
    - 비용/단가 열 위치는 _resolve_cost_price_pairs() 로 검증·교정
    """
    rows = _read_clipboard_rows(raw_text)

    year_int, month_int = (int(p) for p in year_month.split("-")[:2])
    max_col = len(UNIT_INPUT_COLUMN_KEYS)

    by_date: dict[date, dict] = {}
    skipped_blank = 0
    for row in rows:
        if not row:
            continue
        day = _to_number(row[0])
        if not isinstance(day, int) or not 1 <= day <= 31:
            continue  # 머리글 / TOTAL / 빈 행

        values = {}
        for idx in range(1, min(len(row), max_col)):
            key = UNIT_INPUT_COLUMN_KEYS[idx]
            if not key:
                continue
            value = _to_number(row[idx])
            if value is not None:
                values[key] = value

        if not values:
            skipped_blank += 1   # 아직 실적이 없는 미래 일자
            continue

        try:
            by_date[date(year_int, month_int, day)] = values
        except ValueError:
            log.warning(f"  존재하지 않는 일자 무시: {year_month}-{day:02d}")

    if not by_date:
        raise EmptyGridError("일자 행을 하나도 찾지 못했습니다 "
                            "(해당 월 실적 없음 또는 그리드 형식 변경)")

    _resolve_cost_price_pairs(by_date)

    log.info(f"파싱 완료: {len(by_date)}일 (빈 일자 {skipped_blank}일 제외)")
    first_day = min(by_date)
    log.info(f"  {first_day} 미리보기: "
             f"{ {k: by_date[first_day][k] for k in list(by_date[first_day])[:5]} }")
    return by_date


# ---------------------------------------------------------------------------
# MIS RPA 클래스
# ---------------------------------------------------------------------------
class MISUtilityRPA:
    """MIS 에너지 일일 실적 자동 샘플링 RPA ('원단위 실적입력' 단일 화면)"""

    def __init__(self, year_month: str = None, dry_run: bool = False,
                 org_codes: list[str] | None = None,
                 year_months: list[str] | None = None):
        self.dry_run = dry_run
        # 순회할 사업장 — None 이면 전체 (백필 시 일부만 재수집하는 용도)
        self.org_codes = list(org_codes) if org_codes else list(FACTORY_SHEET_MAP)
        self.auto_recover_missing_dates = year_month is None and not year_months

        if self.auto_recover_missing_dates:
            end_date = (datetime.now() - timedelta(days=1)).date()
            self.year_month = end_date.strftime("%Y-%m")
            sheet_names = [FACTORY_SHEET_MAP[code] for code in self.org_codes]
            collected = load_utility_collected_dates(
                energy_builder.DEFAULT_RAW_PATH, sheet_names
            )
            self.year_months = plan_utility_collection_months(end_date, collected)

            if len(self.year_months) > 1:
                previous_day = end_date.replace(day=1) - timedelta(days=1)
                missing_sheets = [
                    name
                    for name, dates in collected.items()
                    if previous_day not in dates
                ]
                log.warning(
                    f"이전 유틸리티 수집 누락 감지: {previous_day} "
                    f"(사업장 {len(missing_sheets)}개) → "
                    f"{self.year_months[0]} 먼저 재수집"
                )
            elif collected and any(collected.values()):
                previous_day = end_date.replace(day=1) - timedelta(days=1)
                log.info(f"이전 유틸리티 수집일 확인 완료: {previous_day} 누락 없음")
            else:
                log.info("기존 유틸리티 수집 이력이 없어 현재 월만 수집합니다.")
        else:
            # 명시된 월/백필 범위에는 자동 복구 월을 섞지 않는다.
            self.year_months = (
                list(year_months) if year_months else [str(year_month)]
            )
            self.year_month = year_month or self.year_months[-1]

        # 좌표 오류로 직전 공장 그리드를 재복사한 사례 — (화면, 년월, 공장, 원본공장)
        self.duplicate_grids: list[tuple[str, str, str, str]] = []
        self.pending_month_records: list[tuple[str, dict]] = []
        self.collected_records: dict[str, dict] = {}

        # 좌표 설정 로드 — { 화면키: { 좌표명: 값 } }
        self.coords = self._load_coords()

        self.app = None
        self.main_window = None
        log.info(f"=== MIS 에너지 RPA 초기화 ===")
        if len(self.year_months) == 1:
            log.info(f"  기준년월: {self.year_months[0]}")
        else:
            log.info(f"  기준년월: {self.year_months[0]} ~ {self.year_months[-1]} "
                     f"({len(self.year_months)}개월)")
        log.info(f"  Dry-run: {self.dry_run}")
        log.info(f"  사업장: {', '.join(FACTORY_SHEET_MAP[c] for c in self.org_codes)}")

    # -----------------------------------------------------------------------
    # 설정
    # -----------------------------------------------------------------------
    def _load_coords(self):
        """JSON 파일에서 좌표 및 대기 시간 설정을 로드한다."""
        global WAIT_SHORT, WAIT_MEDIUM, WAIT_DROPDOWN, WAIT_SCREEN_LOAD, WAIT_QUERY_LOAD
        global WAIT_COPY_CONFIRM

        coord_path = os.path.join(os.path.dirname(__file__), "utility_coords.json")
        try:
            with open(coord_path, "r", encoding="utf-8") as f:
                config = json.load(f)
            log.info(f"설정 파일 로드 완료: {coord_path}")
        except Exception as e:
            log.error(f"설정 파일을 읽을 수 없습니다: {coord_path} ({e})")
            raise SystemExit(1)

        wait = config.get("wait", {})
        WAIT_SHORT = wait.get("short", WAIT_SHORT)
        WAIT_MEDIUM = wait.get("medium", WAIT_MEDIUM)
        WAIT_DROPDOWN = wait.get("dropdown", WAIT_DROPDOWN)
        WAIT_SCREEN_LOAD = wait.get("screen_load", WAIT_SCREEN_LOAD)
        WAIT_QUERY_LOAD = wait.get("query_load", WAIT_QUERY_LOAD)
        WAIT_COPY_CONFIRM = wait.get("copy_confirm", WAIT_COPY_CONFIRM)
        log.info(f"  대기 시간: short={WAIT_SHORT}s, medium={WAIT_MEDIUM}s, "
                 f"dropdown={WAIT_DROPDOWN}s, screen={WAIT_SCREEN_LOAD}s, "
                 f"query={WAIT_QUERY_LOAD}s, copy_confirm={WAIT_COPY_CONFIRM}s")

        return config.get("coords", {})

    def _coord(self, screen: str, name: str):
        """화면별 좌표를 가져온다 (미설정 시 None)."""
        return self.coords.get(screen, {}).get(name)

    def _require_coord(self, screen: str, name: str):
        """필수 좌표 — 없으면 안전을 위해 즉시 중단한다.

        '원단위 실적입력'은 조회 전용이 아닌 **실적입력** 화면이라, 좌표를 추정해
        클릭하면 그리드 셀에 값이 입력될 위험이 있다. 추측 대신 중단한다.
        """
        value = self._coord(screen, name)
        if not value:
            log.error(f"[{screen}] 좌표 '{name}' 이(가) utility_coords.json 에 없습니다.")
            log.error("  utils/클릭좌표기록_실행.bat 으로 해당 위치를 측정해 채워주세요.")
            log.error("  (실적입력 화면이라 좌표 추정 클릭은 데이터 오입력 위험이 있어 중단합니다)")
            raise SystemExit(1)
        return value

    def _validate_coords(self) -> None:
        """실행 전 필요한 좌표가 모두 채워져 있는지 확인한다."""
        for name in ("tree_menu", "factory_dropdown", "month_filter",
                     "query_button", "copy_button", "factory_list"):
            self._require_coord(SCREEN_UNIT_INPUT, name)

    # -----------------------------------------------------------------------
    # MIS 연결
    # -----------------------------------------------------------------------
    def attach_existing_window(self, app, main_window) -> None:
        """오케스트레이터가 이미 연결한 MIS 윈도우를 주입 — 재연결 생략."""
        self.app = app
        self.main_window = main_window

    def connect_mis(self):
        """실행 중인 MIS (신)종합정보 창에 연결한다."""
        if self.main_window is not None:
            log.info("MIS 연결 재사용 (이전 단계에서 연결됨)")
            return
        log.info("MIS 앱 연결 중...")
        # Win32 로 HWND 를 먼저 찾고 handle 로 connect 한다. UIA title_re 스캔은
        # 바탕화면 전체 트리를 순회해 수 초~20초+ 까지 들쭉날쭉 걸리지만, 핸들
        # 연결은 트리 스캔이 없어 ~10ms 로 끝난다.
        hwnd, title = find_mis_window("(신)종합정보")
        if not hwnd:
            log.error("MIS (신)종합정보 창을 찾을 수 없습니다.")
            log.error("MIS (신)종합정보를 먼저 실행해주세요.")
            raise SystemExit(1)
        try:
            self.app = Application(backend="uia").connect(handle=hwnd)
            self.main_window = self.app.window(handle=hwnd)
            log.info(f"MIS 연결 성공: {title}")
        except Exception as e:
            log.error(f"MIS 앱 연결 실패: {e}")
            log.error("MIS (신)종합정보를 먼저 실행해주세요.")
            raise SystemExit(1)

    # -----------------------------------------------------------------------
    # 메뉴 진입
    # -----------------------------------------------------------------------
    def navigate_to_screen(self, screen: str):
        """좌측 트리에서 해당 화면 메뉴를 더블클릭한다."""
        log.info(f"[{screen}] 화면으로 이동 중...")
        # 좌표 기반 더블클릭 — UIA child_window(TreeItem) 검색은 트리 전체를
        # 순회해 수 초~20초+ 까지 들쭉날쭉 걸리므로 좌표 클릭으로 대체.
        x, y = self._require_coord(screen, "tree_menu")
        fast_click(self.main_window, x, y, double=True)
        log.info(f"  트리메뉴 더블클릭 ({x}, {y})")
        time.sleep(WAIT_SCREEN_LOAD)
        log.info("화면 이동 완료")

    # -----------------------------------------------------------------------
    # ORG 공장 선택 (순수 좌표 기반 — MIS는 커스텀 렌더링 UI로 UIA 미지원)
    # -----------------------------------------------------------------------
    def select_factory(self, screen: str, org_code: str):
        """ORG 드롭다운에서 공장을 선택한다 (좌표 기반)."""
        org_name = FACTORY_SHEET_MAP.get(org_code, org_code)
        log.info(f"공장 선택: {org_code} ({org_name})")

        # 1. 드롭다운 클릭하여 열기
        x, y = self._require_coord(screen, "factory_dropdown")
        fast_click(self.main_window, x, y)
        log.info(f"  드롭다운 클릭 ({x}, {y})")
        time.sleep(WAIT_DROPDOWN)

        # 2. 드롭다운이 열린 상태에서 해당 공장 항목 좌표 클릭
        factory_list = self._require_coord(screen, "factory_list")
        item_y = factory_list.get(org_code)
        if item_y is None:
            log.error(f"[{screen}] factory_list 에 '{org_code}' y좌표가 없습니다.")
            raise SystemExit(1)

        fast_click(self.main_window, x, item_y)
        log.info(f"  공장 항목 클릭 ({x}, {item_y})")
        time.sleep(WAIT_SHORT)

    # -----------------------------------------------------------------------
    # 기준년월 설정 (순수 좌표 기반)
    # -----------------------------------------------------------------------
    def set_year_month(self, screen: str, year_month: str | None = None):
        """기준년월 필드에 값을 설정한다 (좌표 기반)."""
        year_month = year_month or self.year_month
        log.info(f"기준년월 설정: {year_month}")

        x, y = self._require_coord(screen, "month_filter")
        log.info(f"  기준년월 클릭 ({x}, {y})")
        fast_click(self.main_window, x, y)
        time.sleep(WAIT_SHORT)
        send_keys("^a")
        time.sleep(WAIT_SHORT)
        send_keys(year_month, with_spaces=True)
        time.sleep(WAIT_SHORT)

    # -----------------------------------------------------------------------
    # 조회
    # -----------------------------------------------------------------------
    def click_query(self, screen: str):
        """조회 버튼을 클릭하고 데이터 로딩을 기다린다 (좌표 기반)."""
        log.info("조회 버튼 클릭...")

        x, y = self._require_coord(screen, "query_button")
        fast_click(self.main_window, x, y)
        log.info(f"  조회 클릭 ({x}, {y})")

        # 로딩 대기
        log.info("  데이터 로딩 대기 중...")
        time.sleep(WAIT_QUERY_LOAD)
        log.info("  조회 완료")

    # -----------------------------------------------------------------------
    # 그리드 복사
    # -----------------------------------------------------------------------
    def copy_grid_data(self, screen: str) -> str:
        """
        그리드 상단의 복사 버튼을 클릭하여 클립보드에 데이터를 복사하고,
        확인 팝업을 닫은 후 클립보드 텍스트를 반환한다.

        로딩 도중 클릭이 흡수되면 클립보드가 갱신되지 않을 수 있으므로
        ClipboardSequenceNumber 폴링으로 갱신 여부를 검증하고, 안 바뀌면 재클릭한다.
        """
        log.info("그리드 데이터 클립보드 복사 중...")

        x, y = self._require_coord(screen, "copy_button")
        log.info(f"  복사 버튼 클릭 ({x}, {y})")

        clipboard_text = ""
        for attempt in range(1, 4):  # 최대 3회 시도
            seq_before = get_clipboard_sequence()
            fast_click(self.main_window, x, y)
            changed = wait_for_clipboard_change(seq_before, timeout=2.0)
            if changed:
                if attempt > 1:
                    log.info(f"  클립보드 갱신 확인 (재시도 {attempt}/3)")
                # 클립보드 읽기 (팝업 닫기 전에! — MIS가 닫을 때 클립보드를 비울 수 있음)
                clipboard_text = get_clipboard_text()
                self._handle_copy_confirm_dialog(screen)
                break
            self._handle_copy_confirm_dialog(screen, use_ok_click=False)
            log.warning(f"  복사 시도 {attempt}/3 — 2s 내 클립보드 변화 없음")
            time.sleep(WAIT_MEDIUM)
        else:
            log.warning("  클립보드 변경 감지 실패 → 기존 클립보드 그대로 읽기")
            time.sleep(WAIT_MEDIUM)
            clipboard_text = get_clipboard_text()

        # 결과 확인
        if not clipboard_text.strip():
            # pandas로 재시도 (OLE/DataObject 포맷 대응)
            log.info("  ctypes 실패 → pandas.read_clipboard 시도")
            try:
                df = pd.read_clipboard(sep="\t", header=None)
                clipboard_text = df.to_csv(sep="\t", index=False, header=False)
                log.info(f"  pandas 클립보드 읽기 성공: {df.shape}")
            except Exception as e:
                log.warning(f"  pandas.read_clipboard 실패: {e}")

        if not clipboard_text.strip():
            log.warning("클립보드가 비어 있습니다!")
        else:
            lines = clipboard_text.strip().split("\n")
            log.info(f"  클립보드 데이터: {len(lines)}행")

        return clipboard_text

    def _handle_copy_confirm_dialog(self, screen: str, use_ok_click: bool = True):
        """'데이터가 클립보드로 복사되었습니다' 확인 팝업을 닫는다."""
        time.sleep(WAIT_COPY_CONFIRM)
        ok_xy = self._coord(screen, "confirm_popup_ok") if use_ok_click else None
        if ok_xy:
            fast_click(self.main_window, ok_xy[0], ok_xy[1])
            log.info(f"  확인 팝업 닫기 완료 (OK 클릭, {ok_xy[0]}, {ok_xy[1]})")
        else:
            send_keys("{ENTER}")
            action = "완료" if use_ok_click else "시도"
            log.info(f"  확인 팝업 닫기 {action} (Enter)")
        time.sleep(WAIT_SHORT)

    # -----------------------------------------------------------------------
    # 화면 단위 수집
    # -----------------------------------------------------------------------
    def _collect_factories(self, screen: str, parser, year_month: str) -> dict:
        """열려 있는 화면에서 사업장을 순회하며 한 달치를 수집한다.

        Returns: { 시트명: { date: { field_key: value } } }
        """
        collected: dict[str, dict] = {}
        # 그리드 지문 → 이미 수집한 사업장. 좌표가 어긋나 드롭다운 선택이 바뀌지
        # 않으면 직전 공장의 그리드를 그대로 다시 복사하게 되는데, 값만 봐서는
        # 알아채기 어렵다. 같은 지문이 두 번 나오면 적재하지 않고 오류로 남긴다.
        fingerprints: dict[str, str] = {}
        for org_code in self.org_codes:
            sheet_name = FACTORY_SHEET_MAP[org_code]
            log.info("-" * 40)
            log.info(f"▶ [{screen}/{year_month}] 사업장 처리: {org_code} → {sheet_name}")
            log.info("-" * 40)
            try:
                self.select_factory(screen, org_code)
                time.sleep(WAIT_SHORT)
                self.click_query(screen)

                clipboard_text = self.copy_grid_data(screen)
                if not clipboard_text.strip():
                    log.warning(f"  {org_code}: 데이터 없음 → 스킵")
                    continue

                fingerprint = hashlib.sha1(
                    clipboard_text.strip().encode("utf-8", "replace")
                ).hexdigest()
                twin = fingerprints.get(fingerprint)
                if twin:
                    log.error(f"  {sheet_name}: 그리드가 '{twin}' 과(와) 완전히 동일 "
                              f"→ 공장 선택이 바뀌지 않았습니다. 적재하지 않습니다.")
                    log.error(f"    utility_coords.json 의 coords.{screen}"
                              f".factory_list.{org_code} y좌표를 확인하세요.")
                    self.duplicate_grids.append((screen, year_month, sheet_name, twin))
                    continue
                fingerprints[fingerprint] = sheet_name

                by_date = parser(clipboard_text, year_month)
                if not by_date:
                    log.warning(f"  {org_code}: 파싱된 날짜 없음 → 스킵")
                    continue
                collected[sheet_name] = by_date

                # MIS 창으로 포커스 복귀
                self.main_window.set_focus()
                time.sleep(WAIT_SHORT)

            except EmptyGridError as e:
                # 해당 월/공장에 실적이 없는 정상 상황 — traceback 없이 넘어간다
                log.warning(f"  {org_code} {year_month}: {e} → 스킵")
            except Exception as e:
                log.error(f"  {org_code} 처리 중 오류: {e}", exc_info=True)
                try:
                    self.main_window.set_focus()
                except Exception:
                    pass
                continue

        return collected

    def collect_screen(self, screen: str, parser) -> dict:
        """한 화면을 열고 기준년월 1개월치를 전 사업장 순회 수집한다."""
        self.navigate_to_screen(screen)
        self.set_year_month(screen)
        return self._collect_factories(screen, parser, self.year_month)

    def collect_months(self, screen: str, parser, year_months: list[str],
                       on_month_done=None) -> dict:
        """한 화면을 한 번만 열고 여러 달을 순회 수집한다 (과거 데이터 백필용).

        Args:
            year_months: ['2021-01', '2021-02', ...] 오름차순 권장
            on_month_done: fn(year_month, month_records) - 월별 수집 결과 전달 훅.
                           수집 단계에서는 결과를 메모리에 보관하는 데 사용한다.

        Returns: 전체 월을 병합한 { 시트명: { date: { field_key: value } } }
        """
        self.navigate_to_screen(screen)

        merged: dict[str, dict] = {}
        for idx, year_month in enumerate(year_months, start=1):
            log.info("=" * 60)
            log.info(f"[{idx}/{len(year_months)}] {year_month} 수집")
            log.info("=" * 60)
            self.set_year_month(screen, year_month)

            month_records = self._collect_factories(screen, parser, year_month)
            for sheet_name, by_date in month_records.items():
                merged.setdefault(sheet_name, {}).update(by_date)

            if on_month_done:
                on_month_done(year_month, month_records)

        return merged

    # -----------------------------------------------------------------------
    # 수집 단계
    # -----------------------------------------------------------------------
    def collect(self) -> bool:
        """사업장을 순회해 데이터를 메모리에 수집하고 가공 대기 상태로 둔다."""
        log.info("=" * 60)
        log.info("MIS 에너지 실적 수집 단계 시작")
        log.info("=" * 60)

        self.pending_month_records = []
        self.collected_records = {}

        self._validate_coords()

        # 1. MIS 연결
        self.connect_mis()
        self.main_window.set_focus()
        time.sleep(WAIT_MEDIUM)

        # 2. 수집 — 적재와 수식 가공은 3종 수집 완료 후 별도 단계에서 수행한다.
        t0 = time.time()
        records = self.collect_months(
            SCREEN_UNIT_INPUT, parse_unit_input_clipboard, self.year_months,
            on_month_done=self._stage_month,
        )
        self.collected_records = records
        elapsed = time.time() - t0

        if not records:
            log.error("아무 데이터도 얻지 못했습니다.")
            self.report_duplicate_grids()
            return False

        day_count = sum(len(v) for v in records.values())
        log.info("=" * 60)
        log.info(f"수집 완료 — {elapsed / 60:.1f}분")
        for sheet_name in sorted(records):
            by_date = records[sheet_name]
            log.info(f"  {sheet_name:8s} {len(by_date):>5}일  "
                     f"({min(by_date)} ~ {max(by_date)})")
        log.info(
            "DRY-RUN 수집 완료 (엑셀 미기록)"
            if self.dry_run
            else f"수집 완료: {len(records)}개 사업장 / 총 {day_count}일 "
                 f"(가공 대기 {len(self.pending_month_records)}개월)"
        )
        self.report_duplicate_grids()
        log.info("=" * 60)
        return True

    def _stage_month(self, year_month: str, month_records: dict) -> None:
        """한 달 수집 결과를 가공 단계까지 메모리에 보관한다."""
        if not month_records:
            log.warning(f"  {year_month}: 수집된 사업장 없음")
            return
        days = sum(len(v) for v in month_records.values())
        log.info(
            f"  {year_month} 수집 보관: "
            f"{len(month_records)}개 사업장 / {days}일"
        )
        self.pending_month_records.append((year_month, month_records))

    # -----------------------------------------------------------------------
    # 가공 단계
    # -----------------------------------------------------------------------
    def process_collected_data(self) -> bool:
        """대기 중인 월별 결과를 RawDB_에너지에 적재하고 수식을 가공한다."""
        if self.dry_run:
            log.info("유틸리티 가공 단계 생략 (dry-run 모드)")
            return True
        if not self.pending_month_records:
            log.error("가공할 유틸리티 수집 결과가 없습니다.")
            return False

        log.info("=" * 60)
        log.info(
            f"유틸리티 가공 단계 시작 — "
            f"{len(self.pending_month_records)}개월 순차 적재"
        )
        log.info("=" * 60)
        self._backup(energy_builder.DEFAULT_RAW_PATH)

        for index, (year_month, month_records) in enumerate(
            self.pending_month_records, start=1
        ):
            log.info(
                f"가공 월 [{index}/{len(self.pending_month_records)}]: {year_month}"
            )
            try:
                energy_builder.write_raw(month_records)
            except Exception as exc:
                log.error(
                    f"유틸리티 가공 실패: {year_month} ({exc})",
                    exc_info=True,
                )
                return False
        return True

    # -----------------------------------------------------------------------
    # 단독 실행 호환: 수집 후 가공
    # -----------------------------------------------------------------------
    def run(self) -> bool:
        if not self.collect():
            return False
        return self.process_collected_data()

    def report_duplicate_grids(self) -> bool:
        """좌표 오류로 중복 수집된 사업장을 요약한다. 있으면 True."""
        if not self.duplicate_grids:
            return False
        log.error(f"⚠ 공장 선택이 바뀌지 않은 사례 {len(self.duplicate_grids)}건 "
                  f"— 해당 사업장은 적재하지 않았습니다:")
        for screen, year_month, sheet_name, twin in self.duplicate_grids:
            log.error(f"    [{screen}/{year_month}] {sheet_name} = {twin} 의 그리드")
        log.error("  → utility_coords.json 의 factory_list y좌표를 실측으로 교정하세요.")
        return True

    # -----------------------------------------------------------------------
    # 보조
    # -----------------------------------------------------------------------
    @staticmethod
    def _backup(path) -> None:
        """산출물 엑셀을 backup/ 폴더에 타임스탬프로 복사한다."""
        path = Path(path)
        if not path.exists():
            return
        backup_dir = path.parent / "backup"
        backup_dir.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = backup_dir / f"{path.stem}_backup_{stamp}{path.suffix}"
        shutil.copy2(path, backup_path)
        log.info(f"백업 생성 완료: {backup_path}")


# ---------------------------------------------------------------------------
# 메인 진입점
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="MIS 에너지 실적 RPA (원단위 실적입력 화면)"
    )
    parser.add_argument(
        "--ym", type=str, default=None,
        help="기준년월 (YYYY-MM) 1개월. 미지정 시 D-1 자동 계산"
    )
    parser.add_argument(
        "--from", dest="ym_from", default=None,
        help="시작 기준년월 (YYYY-MM). 과거 데이터 수집 - --ym 대신 사용"
    )
    parser.add_argument(
        "--to", dest="ym_to", default=None,
        help="종료 기준년월 (YYYY-MM). 기본: D-1 기준월"
    )
    parser.add_argument(
        "--resume", action="store_true",
        help="RawDB 에 이미 있는 달은 건너뛴다 (중단 후 이어서 실행)"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="MIS 조회만 실행, 엑셀 기록하지 않음"
    )
    parser.add_argument(
        "--factories", default=None,
        help="수집할 사업장. 코드 또는 공장명 CSV (예: '논산,경산' / 'F40,F50'). "
             "기본: 전체"
    )
    parser.add_argument(
        "--yes", action="store_true",
        help="여러 달 수집 시 확인 프롬프트 생략"
    )
    parser.add_argument(
        "--skip-build", action="store_true",
        help="MIS 수집만 하고 RawDB 적재·가공은 생략 "
             "(가공은 build_energy_dataset.py 로 별도 수행)"
    )
    args = parser.parse_args()

    if args.ym and args.ym_from:
        raise SystemExit("--ym 과 --from 은 함께 쓸 수 없습니다.")
    if args.ym_to and not args.ym_from:
        raise SystemExit("--to 는 --from 과 함께 써야 합니다.")
    if args.skip_build and args.dry_run:
        raise SystemExit("--skip-build 와 --dry-run 은 함께 쓸 필요가 없습니다 "
                         "(둘 다 엑셀을 기록하지 않습니다).")

    org_codes = resolve_org_codes(args.factories)
    default_to = (datetime.now() - timedelta(days=1)).strftime("%Y-%m")

    if args.ym_from:
        # ── 과거 데이터 수집: 여러 달 순회 ──
        _setup_logging("backfill")
        months = month_range(args.ym_from, args.ym_to or default_to)
        if args.resume:
            months = drop_collected_months(months, org_codes)
            if not months:
                print("\n수집할 달이 없습니다 (모두 RawDB 에 존재).")
                return
        confirm_plan(months, org_codes, args.yes)
    else:
        # ── 일일 수집: 1개월 ──
        _setup_logging()
        # 미지정 자동 실행은 RPA 초기화 시 직전 누락 여부를 확인해 이전 월을
        # 앞에 추가한다. --ym 명시 실행에는 자동 복구를 섞지 않는다.
        months = [args.ym] if args.ym else None

    rpa = MISUtilityRPA(dry_run=args.dry_run, org_codes=org_codes,
                        year_months=months)
    if args.skip_build:
        # 수집만 — 가공은 build_energy_dataset.py 가 담당한다.
        # 메모리에만 남으므로 이 모드는 좌표·파싱 검증용이다.
        if not rpa.collect():
            raise SystemExit(1)
        log.info("수집만 완료 (--skip-build). 적재·가공은 실행되지 않았습니다.")
        return
    rpa.run()


if __name__ == "__main__":
    main()
