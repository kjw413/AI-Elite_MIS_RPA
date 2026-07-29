# MIS 에너지(유틸리티) 일일 실적 자동 샘플링 RPA (pywinauto + openpyxl 기반)
"""
사내 MIS '(신)종합정보' 시스템에서 에너지 일일 실적을 사업장별로 자동 조회 →
클립보드 복사 → RawDB_에너지.xlsx 적재 → DB_에너지.xlsx 재가공하는 RPA 프로그램.

== 2026-07 화면 변경 ==
주수집 화면이 '유틸리티 일자별 사용량 추이' → '원단위 실적입력(일단위)' 로 바뀌었다.
신규 화면은 단가·비용·COD 를 함께 제공하지만 믹스생산량·원단위가 없어, 구 화면을
보완용으로 병행 수집한다.

  화면 1) 원단위 실적입력(일단위)      [unit_input]  — 행=일자, 열=항목
          일자 | 냉동전력 | 공압기 | 전력사용량 | 전력비 | 전력단가 |
                 연료사용량 | 연료비 | 연료단가 | 용수사용량 | 폐수발생량 |
                 원수COD | 배출수COD
  화면 2) 유틸리티 일자별 사용량 추이  [usage_trend] — 행=항목, 열=일자
          믹스생산량 / 전력·연료·용수 원단위 만 취한다

업무 절차:
  1. MIS 앱 연결 (pywinauto UIA backend)
  2. 화면 1 진입 → 기준년월 설정 → 사업장 순회(F1A→F1B→F20→F30→F40→F50)
     조회 → 그리드 복사 → 클립보드 파싱
  3. 화면 2 진입 → 동일 순회 → 믹스생산량·원단위 보완 수집
  4. RawDB_에너지.xlsx 적재 (행=일자) → DB_에너지.xlsx 재가공 (행=항목, 열=날짜)

Usage:
  python utility_daily_rpa.py                # 기본: D-1 기준월
  python utility_daily_rpa.py --ym 2026-07   # 특정 월 지정
  python utility_daily_rpa.py --dry-run      # MIS 조회만, 엑셀 미기록
  python utility_daily_rpa.py --skip-trend   # 구 화면(믹스/원단위) 수집 생략
"""

import sys
import time
import os
import json
import shutil
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
    wait_for_clipboard_change,
)

# ---------------------------------------------------------------------------
# 로깅 설정
# ---------------------------------------------------------------------------
LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)
log = logging.getLogger(__name__)


def _setup_logging() -> None:
    if logging.getLogger().handlers:
        return
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(
                LOG_DIR / f"rpa_{datetime.now():%Y%m%d_%H%M%S}.log",
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

# 화면 식별자 (utility_coords.json 의 coords 키와 동일)
SCREEN_UNIT_INPUT = "unit_input"
SCREEN_USAGE_TREND = "usage_trend"

# ── 화면 1) '원단위 실적입력(일단위)' 그리드 열 순서 ──
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

# ── 화면 2) '유틸리티 일자별 사용량 추이' 그리드 행 순서 ──
# 신규 화면에 없는 믹스생산량·원단위만 취하고 나머지는 버린다.
USAGE_TREND_ROW_KEYS = (
    None,                   # 0  냉동전력량   (화면 1에서 수집)
    None,                   # 1  공압기       (화면 1에서 수집)
    None,                   # 2  전력량       (화면 1에서 수집)
    None,                   # 3  연료량       (화면 1에서 수집)
    None,                   # 4  용수량       (화면 1에서 수집)
    None,                   # 5  폐수량       (화면 1에서 수집)
    "mix_prod_kg",          # 6  믹스생산량[kg]
    "power_per_ton_kwh",    # 7  전력원단위[kWh/mix-ton]
    "fuel_per_ton_nm3",     # 8  연료원단위[N㎥/mix-ton]
    "water_per_ton_ton",    # 9  용수원단위[ton/mix-ton]
)

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
        raise ValueError("일자 행을 하나도 찾지 못했습니다 (그리드 형식 변경?)")

    _resolve_cost_price_pairs(by_date)

    log.info(f"파싱 완료: {len(by_date)}일 (빈 일자 {skipped_blank}일 제외)")
    first_day = min(by_date)
    log.info(f"  {first_day} 미리보기: "
             f"{ {k: by_date[first_day][k] for k in list(by_date[first_day])[:5]} }")
    return by_date


# ---------------------------------------------------------------------------
# 화면 2) '유틸리티 일자별 사용량 추이' 클립보드 파싱
# ---------------------------------------------------------------------------
def parse_usage_trend_clipboard(raw_text: str, year_month: str) -> dict:
    """구 화면 클립보드(행=항목, 열=일자)에서 믹스생산량·원단위만 추출한다.

    Returns: { date: { field_key: value } }
    """
    rows = _read_clipboard_rows(raw_text)

    # 첫 행: 헤더 (일자)
    day_numbers = []
    for col in rows[0]:
        m = re.search(r"(\d+)", str(col).strip())
        if m:
            day_numbers.append(int(m.group(1)))

    year_int, month_int = (int(p) for p in year_month.split("-")[:2])
    dates: list[date] = []
    for day in day_numbers:
        try:
            dates.append(date(year_int, month_int, day))
        except ValueError:
            log.warning(f"  존재하지 않는 일자 무시: {year_month}-{day:02d}")

    # 데이터 행 파싱 (첫 열 = 항목명 제거)
    by_date: dict[date, dict] = {}
    for row_idx, row in enumerate(rows[1:]):
        if row_idx >= len(USAGE_TREND_ROW_KEYS):
            break
        key = USAGE_TREND_ROW_KEYS[row_idx]
        if not key or len(row) < 2:
            continue
        for day_idx, cell in enumerate(row[1:]):
            if day_idx >= len(dates):
                break
            value = _to_number(cell)
            if value is not None:
                by_date.setdefault(dates[day_idx], {})[key] = value

    log.info(f"파싱 완료(보완): {len(by_date)}일 × {len([k for k in USAGE_TREND_ROW_KEYS if k])}항목")
    return by_date


# ---------------------------------------------------------------------------
# MIS RPA 클래스
# ---------------------------------------------------------------------------
class MISUtilityRPA:
    """MIS 에너지 일일 실적 자동 샘플링 RPA (신규 화면 + 구 화면 병행 수집)"""

    def __init__(self, year_month: str = None, dry_run: bool = False,
                 skip_trend: bool = False):
        if year_month is None:
            ref_date = datetime.now() - timedelta(days=1)
            self.year_month = ref_date.strftime("%Y-%m")
        else:
            self.year_month = year_month
        self.dry_run = dry_run
        self.skip_trend = skip_trend

        # 좌표 설정 로드 — { 화면키: { 좌표명: 값 } }
        self.coords = self._load_coords()

        self.app = None
        self.main_window = None
        log.info(f"=== MIS 에너지 RPA 초기화 ===")
        log.info(f"  기준년월: {self.year_month}  (D-1 자동 계산)")
        log.info(f"  Dry-run: {self.dry_run} / 구 화면 생략: {self.skip_trend}")

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
        screens = [SCREEN_UNIT_INPUT]
        if not self.skip_trend:
            screens.append(SCREEN_USAGE_TREND)
        for screen in screens:
            for name in ("tree_menu", "factory_dropdown", "month_filter",
                         "query_button", "copy_button", "factory_list"):
                self._require_coord(screen, name)

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
    def set_year_month(self, screen: str):
        """기준년월 필드에 값을 설정한다 (좌표 기반)."""
        log.info(f"기준년월 설정: {self.year_month}")

        x, y = self._require_coord(screen, "month_filter")
        log.info(f"  기준년월 클릭 ({x}, {y})")
        fast_click(self.main_window, x, y)
        time.sleep(WAIT_SHORT)
        send_keys("^a")
        time.sleep(WAIT_SHORT)
        send_keys(self.year_month, with_spaces=True)
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
    def collect_screen(self, screen: str, parser) -> dict:
        """한 화면을 열고 전 사업장을 순회하며 수집한다.

        Returns: { 시트명: { date: { field_key: value } } }
        """
        self.navigate_to_screen(screen)
        self.set_year_month(screen)

        collected: dict[str, dict] = {}
        for org_code, sheet_name in FACTORY_SHEET_MAP.items():
            log.info("-" * 40)
            log.info(f"▶ [{screen}] 사업장 처리: {org_code} → {sheet_name}")
            log.info("-" * 40)
            try:
                self.select_factory(screen, org_code)
                time.sleep(WAIT_SHORT)
                self.click_query(screen)

                clipboard_text = self.copy_grid_data(screen)
                if not clipboard_text.strip():
                    log.warning(f"  {org_code}: 데이터 없음 → 스킵")
                    continue

                by_date = parser(clipboard_text, self.year_month)
                if not by_date:
                    log.warning(f"  {org_code}: 파싱된 날짜 없음 → 스킵")
                    continue
                collected[sheet_name] = by_date

                # MIS 창으로 포커스 복귀
                self.main_window.set_focus()
                time.sleep(WAIT_SHORT)

            except Exception as e:
                log.error(f"  {org_code} 처리 중 오류: {e}", exc_info=True)
                try:
                    self.main_window.set_focus()
                except Exception:
                    pass
                continue

        return collected

    # -----------------------------------------------------------------------
    # 전체 실행
    # -----------------------------------------------------------------------
    def run(self):
        """두 화면을 순회하며 데이터를 추출 → RawDB 적재 → DB 재가공한다."""
        log.info("=" * 60)
        log.info("MIS 에너지 일일 실적 RPA 시작")
        log.info("=" * 60)

        self._validate_coords()

        # 1. 구 형식 파일 1회 이관 (RawDB_에너지[전치형] → DB_에너지)
        if not self.dry_run:
            energy_builder.migrate_legacy_rawdb()

        # 2. MIS 연결
        self.connect_mis()
        self.main_window.set_focus()
        time.sleep(WAIT_MEDIUM)

        # 3. 산출물 백업
        if not self.dry_run:
            for path in (energy_builder.DEFAULT_RAW_PATH,
                         energy_builder.DEFAULT_OUTPUT_PATH):
                self._backup(path)

        # 4. 화면 1 — 주수집 (사용량/단가/비용/COD)
        records = self.collect_screen(SCREEN_UNIT_INPUT, parse_unit_input_clipboard)
        if not records:
            log.error("주수집 화면에서 아무 데이터도 얻지 못했습니다. 중단합니다.")
            raise SystemExit(1)

        # 5. 화면 2 — 보완 수집 (믹스생산량/원단위)
        #    실패해도 주수집 결과는 반드시 적재한다.
        if self.skip_trend:
            log.info("구 화면(믹스/원단위) 수집 생략 — --skip-trend")
        else:
            try:
                trend = self.collect_screen(SCREEN_USAGE_TREND,
                                            parse_usage_trend_clipboard)
                self._merge(records, trend)
            except Exception as e:
                log.error(f"구 화면 보완 수집 실패 (주수집 결과는 적재 계속): {e}",
                          exc_info=True)

        # 6. 적재 + 재가공
        day_count = sum(len(v) for v in records.values())
        if self.dry_run:
            log.info(f"[DRY-RUN] {len(records)}개 사업장 × 총 {day_count}일 (기록 안함)")
            for sheet_name, by_date in records.items():
                log.info(f"  {sheet_name}: {len(by_date)}일 "
                         f"({min(by_date)} ~ {max(by_date)})")
        else:
            energy_builder.write_raw(records)
            energy_builder.build_dataset()

        log.info("=" * 60)
        log.info("DRY-RUN 완료 (엑셀 미기록)" if self.dry_run
                 else f"RPA 완료: {len(records)}개 사업장 × 총 {day_count}일 적재")
        log.info("=" * 60)

    # -----------------------------------------------------------------------
    # 보조
    # -----------------------------------------------------------------------
    @staticmethod
    def _merge(base: dict, extra: dict) -> None:
        """보완 수집 결과를 주수집 결과에 병합한다 (주수집 값 우선 보존)."""
        for sheet_name, by_date in extra.items():
            target = base.setdefault(sheet_name, {})
            for day, values in by_date.items():
                slot = target.setdefault(day, {})
                for key, value in values.items():
                    slot.setdefault(key, value)

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
    _setup_logging()
    parser = argparse.ArgumentParser(
        description="MIS 에너지 일일 실적 RPA (원단위 실적입력 + 사용량 추이)"
    )
    parser.add_argument(
        "--ym", type=str, default=None,
        help="기준년월 (YYYY-MM). 미지정 시 D-1 자동 계산"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="MIS 조회만 실행, 엑셀 기록하지 않음"
    )
    parser.add_argument(
        "--skip-trend", action="store_true",
        help="구 화면(믹스생산량·원단위) 보완 수집 생략"
    )
    args = parser.parse_args()

    rpa = MISUtilityRPA(year_month=args.ym, dry_run=args.dry_run,
                        skip_trend=args.skip_trend)
    rpa.run()


if __name__ == "__main__":
    main()
