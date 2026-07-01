# MIS 재공품 자동 샘플링 RPA (pywinauto + openpyxl 기반)
"""
사내 MIS '(신)종합정보' '생산계획 대비 실적현황(완제품/재공품)' 화면에서
공장별 재공품 (category1=재공품) 데이터를 자동 조회하여
E:\\Sampled DB\\RawDB_재공품.xlsx 의 기존 시트(남양주1/남양주2/김해/광주/논산) 데이터를
지우고 MIS 결과로 통째 교체한다.

  F10 (MIS 남양주공장 통합)  → 남양주1, 남양주2 두 시트에 동일 paste
                              (WIP_refactoring.py 가 ItemCode 기반 F10A/F10B 분리)
  F20 (김해)                  → 김해
  F30 (광주)                  → 광주
  F40 (논산)                  → 논산

업무 절차:
  1. MIS 앱 연결 (pywinauto UIA backend)
  2. 트리 메뉴 '생산계획 대비 실적현황(완제품/재공품)' 진입 (더블클릭)
  3. 기준일자 시작일 클릭 → 입력 → TAB → 종료일 입력 → ENTER
       (시작 = D-2가 속한 월의 1일, 종료 = D-2)
  4. Category1 드롭다운 → '재공품' 선택
  5. 표시구분 '소계' 체크박스 해제 (기본 체크 가정 → 1회 토글)
  6. '실적일자 기준' 탭 선택
  7. 공장(F10→F20→F30→F40) 순회:
     a. ORG 드롭다운 공장 선택
     b. 조회 + 로딩 대기
     c. 그리드 좌상단 헤더 클릭 → 클립보드 복사
     d. 확인 팝업 Enter
     e. RawDB_재공품.xlsx 의 매핑된 기존 시트 데이터 클리어 후 통째 paste
  8. (자동 연결) RawDB_재공품.xlsx → DB_재공품.xlsx 통합
     - tools.scripts.WIP_refactoring.main() 호출
     - WIP_refactoring 가 ItemCode 기반으로 F10A/F10B 분리, 연속 날짜 보정, 출력 백업까지 처리

Usage:
  python wip_daily_rpa.py                  # 기본: D-2 + DB 통합 자동 수행
  python wip_daily_rpa.py --date 2026-05-14
  python wip_daily_rpa.py --dry-run        # MIS 조회만, Excel/DB 미기록
  python wip_daily_rpa.py --skip-db-build  # Raw 샘플링만, DB 통합 생략
"""

import sys
import time
import os
import json
import shutil
import logging
import argparse
from datetime import datetime, timedelta
from pathlib import Path

from pywinauto import Application
from pywinauto.keyboard import send_keys
from pywinauto.timings import Timings

# pywinauto 내부 click/keys 대기 시간 단축 (MIS는 즉시 반응한다는 가정)
Timings.after_clickinput_wait = 0.01
Timings.after_setfocus_wait = 0.01
Timings.after_sendkeys_key_wait = 0.001

# 프로젝트 루트(app/...) import 보장. tools/mis_rpa/file.py → 2단계 위가 루트.
_PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from mis_rpa._common import (
    fast_click,
    find_mis_window,
    get_clipboard_sequence,
    get_clipboard_text,
    parse_clipboard_rows,
    paste_to_sheet,
    sampled_db_path,
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
                LOG_DIR / f"wip_rpa_{datetime.now():%Y%m%d_%H%M%S}.log",
                encoding="utf-8",
            ),
        ],
    )

# ---------------------------------------------------------------------------
# 상수 정의
# ---------------------------------------------------------------------------
RAW_FILE = sampled_db_path("RawDB_재공품.xlsx", "WIP_ITEM_MASTER_XLSX")
CATEGORY1 = "재공품"

# MIS 공장(ORG) 코드 → RawDB_재공품.xlsx 의 기존 시트명 매핑.
# 시트명은 WIP_refactoring.py 의 INPUT_SHEET_BY_PLANT 와 정합.
# 남양주는 MIS에서 F10 통합 추출 → 같은 데이터를 남양주1/남양주2 두 시트에 동일 paste
# (WIP_refactoring.py 가 F10A_ITEMCODES / F10B_ITEMCODES 로 사후 분리).
FACTORY_SHEET_MAP: dict[str, list[str]] = {
    "F10": ["남양주1", "남양주2"],
    "F20": ["김해"],
    "F30": ["광주"],
    "F40": ["논산"],
}

# 그리드 데이터 '공장명' 컬럼에 나타나는 지역 토큰 — 쿼리 로딩 검증용.
# 복사한 클립보드 내용에 현재 선택한 공장의 토큰이 없으면 '아직 로딩 안된
# 이전 화면/공장 데이터(stale)' 로 보고 재복사한다. 각 공장 그리드는 자기
# 공장명만 담으므로(예: F20 → 김해공장 행만) 교차오염을 직접 잡아낸다.
FACTORY_NAME_TOKEN: dict[str, str] = {
    "F10": "남양주",
    "F20": "김해",
    "F30": "광주",
    "F40": "논산",
}

# 대기 시간 기본값 (wip_coords.json의 "wait" 값으로 덮어씌워짐)
WAIT_SHORT = 0.05        # 클릭/타이핑 후 미세 대기 (MIS는 즉시 반응)
WAIT_MEDIUM = 0.2        # 클립보드 fill / 필드 클릭 후 약간 더 긴 대기
WAIT_DROPDOWN = 0.01     # 드롭다운 펼침 후 항목 클릭 전 대기
WAIT_SCREEN_LOAD = 1.0   # 사이드바 메뉴 클릭 → MIS 화면 전환 로딩
WAIT_QUERY_LOAD = 1.5    # 조회 버튼 → 그리드 데이터 로딩(첫 복사 전 기본 대기)
WAIT_QUERY_MAX = 60.0    # 그리드 로딩 검증 재복사 최대 대기(월말 대용량 대비)
WAIT_QUERY_RETRY = 1.0   # 재복사 시도 간 간격
WAIT_COPY_CONFIRM = 0.4  # 복사 버튼 클릭 후 확인 팝업이 포커스를 받을 때까지 대기


# ---------------------------------------------------------------------------
# MIS WIP RPA 클래스
# ---------------------------------------------------------------------------
class MISWIPRPA:
    """MIS 재공품 자동 샘플링 RPA"""

    def __init__(
        self,
        ref_date: str | None = None,
        dry_run: bool = False,
        build_db: bool = True,
    ):
        if ref_date is None:
            d = datetime.now() - timedelta(days=2)
        else:
            d = datetime.strptime(ref_date, "%Y-%m-%d")

        self.end_date_obj = d
        month_start = d.replace(day=1)

        # 최소 3일 조회 보장.
        # WIP_refactoring.get_update_period 가 양 끝 날짜를 부분 데이터 보호 차원에서
        # 제외하므로 (옵션 B), 기간이 3일 미만이면 update window 가 비거나 폴백으로
        # 양 끝 제외가 무력화된다. 기준일이 월 1·2일이면 month_start 만으로는 1~2일
        # 폭이므로 직전 월까지 끌어와 최소 3일을 확보한다.
        MIN_SPAN_DAYS = 3
        min_start = d - timedelta(days=MIN_SPAN_DAYS - 1)
        self.start_date_obj = min(month_start, min_start)

        self.start_date = self.start_date_obj.strftime("%Y-%m-%d")
        self.end_date = self.end_date_obj.strftime("%Y-%m-%d")

        self.dry_run = dry_run
        self.build_db = build_db
        self.coords = self._load_coords()

        self.app = None
        self.main_window = None
        log.info("=== MIS 재공품 RPA 초기화 ===")
        log.info(f"  기준일자: {self.start_date} ~ {self.end_date}")
        log.info(f"  대상 시트: {sum((v for v in FACTORY_SHEET_MAP.values()), [])}")
        log.info(f"  Dry-run : {self.dry_run}")
        log.info(f"  DB 통합 : {'실행' if self.build_db else '생략'}")

    # -----------------------------------------------------------------------
    # 좌표 설정 로드
    # -----------------------------------------------------------------------
    def _load_coords(self):
        global WAIT_SHORT, WAIT_MEDIUM, WAIT_DROPDOWN, WAIT_SCREEN_LOAD, WAIT_QUERY_LOAD
        global WAIT_QUERY_MAX, WAIT_QUERY_RETRY, WAIT_COPY_CONFIRM

        coord_path = os.path.join(os.path.dirname(__file__), "wip_coords.json")
        try:
            with open(coord_path, "r", encoding="utf-8") as f:
                config = json.load(f)
                log.info(f"설정 파일 로드: {coord_path}")
                wait = config.get("wait", {})
                WAIT_SHORT = wait.get("short", WAIT_SHORT)
                WAIT_MEDIUM = wait.get("medium", WAIT_MEDIUM)
                WAIT_DROPDOWN = wait.get("dropdown", WAIT_DROPDOWN)
                WAIT_SCREEN_LOAD = wait.get("screen_load", WAIT_SCREEN_LOAD)
                WAIT_QUERY_LOAD = wait.get("query_load", WAIT_QUERY_LOAD)
                WAIT_QUERY_MAX = wait.get("query_max", WAIT_QUERY_MAX)
                WAIT_QUERY_RETRY = wait.get("query_retry", WAIT_QUERY_RETRY)
                WAIT_COPY_CONFIRM = wait.get("copy_confirm", WAIT_COPY_CONFIRM)
                log.info(f"  대기시간: short={WAIT_SHORT}s, medium={WAIT_MEDIUM}s, "
                         f"dropdown={WAIT_DROPDOWN}s, screen={WAIT_SCREEN_LOAD}s, "
                         f"query={WAIT_QUERY_LOAD}s, query_max={WAIT_QUERY_MAX}s, "
                         f"copy_confirm={WAIT_COPY_CONFIRM}s")
                return config.get("coords", {})
        except Exception as e:
            log.warning(f"설정 파일을 읽을 수 없습니다 ({e}). 기본값 사용.")
            return {}

    # -----------------------------------------------------------------------
    # MIS 연결
    # -----------------------------------------------------------------------
    def attach_existing_window(self, app, main_window) -> None:
        """오케스트레이터가 이미 연결한 MIS 윈도우를 주입 — 재연결 생략."""
        self.app = app
        self.main_window = main_window

    def connect_mis(self):
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
    # 메뉴 진입 (트리메뉴 더블클릭)
    # -----------------------------------------------------------------------
    def navigate_to_wip_screen(self):
        log.info("'생산계획 대비 실적현황(완제품/재공품)' 화면으로 이동 중...")
        # 좌표 기반 더블클릭 (트리메뉴 노드)
        x, y = self.coords.get("tree_menu", [169, 176])
        fast_click(self.main_window, x, y, double=True)
        log.info(f"  트리메뉴 더블클릭 ({x}, {y})")
        time.sleep(WAIT_SCREEN_LOAD)

    # -----------------------------------------------------------------------
    # 공장 선택
    # -----------------------------------------------------------------------
    def select_factory(self, org_code: str):
        log.info(f"공장 선택: {org_code}")
        x, y = self.coords.get("factory_dropdown", [440, 108])
        fast_click(self.main_window, x, y)
        log.info(f"  드롭다운 클릭 ({x}, {y})")
        time.sleep(WAIT_DROPDOWN)

        factory_list = self.coords.get("factory_list", {})
        item_y = factory_list.get(org_code)
        if item_y is None:
            log.error(f"  {org_code} 좌표 미정의 → wip_coords.json 추가 필요")
            raise RuntimeError(f"factory_list[{org_code}] 좌표 없음")

        fast_click(self.main_window, x, item_y)
        log.info(f"  공장 항목 클릭 ({x}, {item_y})")
        time.sleep(WAIT_SHORT)

    # -----------------------------------------------------------------------
    # 기준일자 설정 (시작 클릭/입력/TAB/종료 입력/ENTER)
    # -----------------------------------------------------------------------
    def set_date_range(self):
        log.info(f"기준일자 설정: {self.start_date} ~ {self.end_date}")

        x, y = self.coords.get("start_date_field", [627, 130])
        fast_click(self.main_window, x, y)
        log.info(f"  시작일 클릭 ({x}, {y})")
        time.sleep(WAIT_MEDIUM)
        send_keys("^a")
        time.sleep(WAIT_SHORT)
        send_keys(self.start_date, with_spaces=True)
        log.info(f"  시작일 입력: {self.start_date}")
        time.sleep(WAIT_SHORT)

        # TAB → 종료일로 자동 포커스 이동
        send_keys("{TAB}")
        time.sleep(WAIT_SHORT)

        send_keys("^a")
        time.sleep(WAIT_SHORT)
        send_keys(self.end_date, with_spaces=True)
        log.info(f"  종료일 입력: {self.end_date}")
        time.sleep(WAIT_SHORT)

        send_keys("{ENTER}")
        time.sleep(WAIT_SHORT)

    # -----------------------------------------------------------------------
    # Category1 = 재공품 선택
    # -----------------------------------------------------------------------
    def select_category1_wip(self):
        log.info(f"Category1 선택: {CATEGORY1}")
        x, y = self.coords.get("category1_dropdown", [854, 104])
        fast_click(self.main_window, x, y)
        time.sleep(WAIT_DROPDOWN)

        cat_list = self.coords.get("category1_list", {})
        item_y = cat_list.get(CATEGORY1)
        if item_y is None:
            log.error(f"  {CATEGORY1} 좌표 미정의")
            raise RuntimeError(f"category1_list[{CATEGORY1}] 좌표 없음")

        fast_click(self.main_window, x, item_y)
        log.info(f"  Category1 클릭 ({x}, {item_y})")
        time.sleep(WAIT_SHORT)

    # -----------------------------------------------------------------------
    # 표시구분 '소계' 체크박스 해제 (1회 토글)
    # -----------------------------------------------------------------------
    def toggle_subtotal_off(self):
        x, y = self.coords.get("subtotal_checkbox", [1210, 107])
        fast_click(self.main_window, x, y)
        log.info(f"표시구분 '소계' 체크박스 토글 ({x}, {y})")
        time.sleep(WAIT_SHORT)

    # -----------------------------------------------------------------------
    # '실적일자 기준' 탭 선택
    # -----------------------------------------------------------------------
    def select_actual_basis_tab(self):
        x, y = self.coords.get("actual_basis_tab", [451, 183])
        fast_click(self.main_window, x, y)
        log.info(f"'실적일자 기준' 탭 클릭 ({x}, {y})")
        time.sleep(WAIT_SHORT)

    # -----------------------------------------------------------------------
    # 조회
    # -----------------------------------------------------------------------
    def click_query(self):
        log.info("조회 버튼 클릭...")
        x, y = self.coords.get("query_button", [347, 78])
        fast_click(self.main_window, x, y)
        log.info(f"  조회 클릭 ({x}, {y})")
        log.info("  데이터 로딩 대기 중...")
        time.sleep(WAIT_QUERY_LOAD)

    # -----------------------------------------------------------------------
    # 그리드 복사
    # -----------------------------------------------------------------------
    def copy_grid_data(self, expected_token: str | None = None) -> str:
        """그리드 좌상단 복사 버튼을 눌러 클립보드 데이터를 읽는다.

        두 단계 검증으로 'stale 데이터(이전 화면/공장)' paste 를 방지한다:
          1) ClipboardSequenceNumber — 복사 클릭으로 클립보드가 실제로 갱신됐는지.
          2) 내용 검증 — expected_token(현재 공장의 '공장명' 지역 토큰)이 복사된
             데이터 안에 있는지. 월말 대용량 조회는 로딩이 수십 초 걸릴 수 있어,
             클립보드가 갱신돼도 그 내용이 '아직 안 바뀐 이전 공장 데이터'일 수
             있다. 토큰이 일치할 때까지 WAIT_QUERY_MAX 동안 재복사한다.

        expected_token 이 None 이면(검증 불가) 비어있지 않은 첫 갱신을 채택한다.
        """
        x, y = self.coords.get("copy_button", [329, 207])
        log.info(f"그리드 좌상단 복사 버튼 클릭 ({x}, {y})"
                 + (f"  (기대 공장: {expected_token})" if expected_token else ""))

        deadline = time.monotonic() + WAIT_QUERY_MAX
        clipboard_text = ""
        last_text = ""
        attempt = 0

        while True:
            attempt += 1
            seq_before = get_clipboard_sequence()
            fast_click(self.main_window, x, y)
            changed = wait_for_clipboard_change(seq_before, timeout=2.0)

            if changed:
                # 복사가 실제로 일어났으면 '복사되었습니다' 팝업이 떴으니 닫는다.
                self._handle_copy_confirm_dialog()
                text = get_clipboard_text()
            else:
                self._handle_copy_confirm_dialog(use_ok_click=False)
                text = ""

            if text.strip():
                last_text = text
                if self._content_is_current(text, expected_token):
                    clipboard_text = text
                    if attempt > 1:
                        log.info(f"  유효 데이터 확보 (시도 {attempt}회)")
                    break
                reason = f"기대 공장('{expected_token}') 토큰 없음 → 이전/미로딩 데이터"
            else:
                reason = "클립보드 미갱신(그리드 로딩 중 추정)"

            if time.monotonic() >= deadline:
                log.error(
                    f"  쿼리 로딩 타임아웃({WAIT_QUERY_MAX:.0f}s) — {attempt}회 시도에도 "
                    f"유효 데이터 미확보. stale 가능성 있어 이 공장은 스킵 처리됨."
                )
                # 검증 실패한 stale 데이터는 반환하지 않는다(빈 문자열 → 호출부 스킵).
                return ""

            log.warning(
                f"  복사 시도 {attempt} — {reason}, "
                f"{WAIT_QUERY_RETRY:.0f}s 후 재시도 (남은 {deadline - time.monotonic():.0f}s)"
            )
            time.sleep(WAIT_QUERY_RETRY)

        # win32clipboard 가 끝내 비었으면 pandas 폴백(드문 경우)
        if not clipboard_text.strip() and last_text.strip():
            clipboard_text = last_text

        lines = clipboard_text.strip().split("\n")
        log.info(f"  클립보드 데이터: {len(lines)}행")
        return clipboard_text

    @staticmethod
    def _content_is_current(text: str, expected_token: str | None) -> bool:
        """복사된 데이터가 '현재 공장의 갓 로딩된 데이터'인지 판정.

        expected_token 미지정 시: 비어있지 않으면 통과(검증 불가 → 기존 동작).
        지정 시: 데이터 안에 토큰이 있어야 통과. 각 공장 그리드는 자기 공장명만
        담으므로, 이전 공장/이전 화면(유틸리티 등) 데이터는 토큰이 없어 걸러진다.
        """
        if not expected_token:
            return True
        return expected_token in text

    def _handle_copy_confirm_dialog(self, use_ok_click: bool = True):
        time.sleep(WAIT_COPY_CONFIRM)
        ok_xy = self.coords.get("confirm_popup_ok") if use_ok_click else None
        if ok_xy:
            fast_click(self.main_window, ok_xy[0], ok_xy[1])
            log.info(f"  확인 팝업 닫기 완료 (OK 클릭, {ok_xy[0]}, {ok_xy[1]})")
        else:
            send_keys("{ENTER}")
            action = "완료" if use_ok_click else "시도"
            log.info(f"  확인 팝업 닫기 {action} (Enter)")
        time.sleep(WAIT_SHORT)

    # -----------------------------------------------------------------------
    # 출력 파일 백업 (단일 파일)
    # -----------------------------------------------------------------------
    def backup_output(self):
        if self.dry_run or not os.path.exists(RAW_FILE):
            return
        backup_root = os.path.join(os.path.dirname(RAW_FILE), "backup")
        os.makedirs(backup_root, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(
            backup_root, f"RawDB_재공품_{timestamp}.xlsx"
        )
        try:
            shutil.copy2(RAW_FILE, backup_path)
            log.info(f"백업 생성: {backup_path}")
        except Exception as e:
            log.warning(f"백업 실패: {e}")

    # -----------------------------------------------------------------------
    # DB 통합 (RawDB_재공품.xlsx → DB_재공품.xlsx)
    # -----------------------------------------------------------------------
    def consolidate_to_db(self) -> bool:
        """
        RawDB_재공품.xlsx → DB_재공품.xlsx 통합.

        tools.scripts.WIP_refactoring.main() 을 그대로 호출 — 빌드 로직은 한 곳에서 관리.
        WIP_refactoring가 자체적으로 출력 백업 + 시트별(남양주1/2, 김해, 광주, 논산) 처리.
        """
        log.info("=" * 60)
        log.info("DB 통합 단계 시작 (RawDB_재공품 → DB_재공품.xlsx)")
        log.info("=" * 60)
        try:
            from mis_rpa import wip_refactoring as WIP_refactoring
        except Exception as exc:
            log.error(f"WIP_refactoring import 실패: {exc}", exc_info=True)
            return False

        try:
            t0 = datetime.now()
            WIP_refactoring.main()
            dt = (datetime.now() - t0).total_seconds()
        except Exception as exc:
            log.error(f"DB 통합 실패: {exc}", exc_info=True)
            return False

        log.info(f"DB 통합 완료 — {dt:.1f}s")
        log.info(f"  출력 파일: {WIP_refactoring.OUTPUT_FILE}")
        return True

    # -----------------------------------------------------------------------
    # 전체 실행
    # -----------------------------------------------------------------------
    def run(self):
        log.info("=" * 60)
        log.info("MIS 재공품 RPA 시작")
        log.info("=" * 60)

        self.connect_mis()
        self.main_window.set_focus()
        time.sleep(WAIT_MEDIUM)

        # 화면 진입 + 조회 전 공통 설정 (공장 무관)
        self.navigate_to_wip_screen()
        self.set_date_range()
        self.select_category1_wip()
        self.toggle_subtotal_off()
        self.select_actual_basis_tab()

        self.backup_output()

        total_rows = 0
        success = 0
        failed = 0

        for factory, sheet_names in FACTORY_SHEET_MAP.items():
            log.info("=" * 50)
            log.info(f"▶ {factory} → 시트 {sheet_names}")
            log.info("=" * 50)

            try:
                self.select_factory(factory)
                self.click_query()
                expected_token = FACTORY_NAME_TOKEN.get(factory)
                clipboard_text = self.copy_grid_data(expected_token)

                if not clipboard_text.strip():
                    log.warning(f"  유효 데이터 미확보(로딩 타임아웃/stale) → 스킵")
                    failed += 1
                    continue

                rows = parse_clipboard_rows(clipboard_text)
                if not rows:
                    log.warning(f"  파싱 결과 없음 → 스킵")
                    failed += 1
                    continue

                # 같은 MIS 결과를 매핑된 모든 기존 시트에 동일 paste
                # (남양주는 F10 → 남양주1·남양주2 양쪽으로 복제)
                if self.dry_run:
                    log.info(f"  [DRY-RUN] {sheet_names}: "
                             f"{len(rows)}행 × {len(rows[0])}열 (기록 안함)")
                    success += 1
                else:
                    factory_ok = True
                    for sheet_name in sheet_names:
                        written = paste_to_sheet(RAW_FILE, sheet_name, rows)
                        total_rows += written
                        if written == 0:
                            factory_ok = False
                    if factory_ok:
                        success += 1
                    else:
                        failed += 1

                self.main_window.set_focus()
                time.sleep(WAIT_SHORT)

            except Exception as e:
                log.error(f"  처리 중 오류: {e}", exc_info=True)
                failed += 1
                try:
                    self.main_window.set_focus()
                except Exception:
                    pass
                continue

        log.info("=" * 60)
        if self.dry_run:
            log.info(f"DRY-RUN 완료 (성공 {success} / 실패 {failed})")
        else:
            log.info(f"RPA 완료: 성공 {success} / 실패 {failed} / 총 {total_rows}행 기록")
        log.info("=" * 60)

        # DB 통합 — dry-run / build_db 끔 / 성공 0건이면 생략
        if not self.build_db:
            log.info("DB 통합 단계 생략 (--skip-db-build)")
            return
        if self.dry_run:
            log.info("DB 통합 단계 생략 (dry-run 모드)")
            return
        if success == 0:
            log.warning("DB 통합 단계 생략 (Raw 단계 성공 0건)")
            return

        self.consolidate_to_db()


# ---------------------------------------------------------------------------
# 메인 진입점
# ---------------------------------------------------------------------------
def main():
    _setup_logging()
    parser = argparse.ArgumentParser(
        description="MIS 재공품 자동 샘플링 RPA "
                    "(생산계획 대비 실적현황(완제품/재공품) → RawDB_재공품.xlsx → DB_재공품.xlsx)"
    )
    parser.add_argument(
        "--date", type=str, default=None,
        help="기준 종료일 (YYYY-MM-DD). 미지정 시 D-2 자동. 시작일은 해당 월 1일."
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="MIS 조회만 실행, Excel 기록하지 않음 (DB 통합도 생략)"
    )
    parser.add_argument(
        "--skip-db-build", action="store_true",
        help="Raw 샘플링만 수행하고 DB_재공품.xlsx 통합 단계는 생략"
    )
    args = parser.parse_args()

    rpa = MISWIPRPA(
        ref_date=args.date,
        dry_run=args.dry_run,
        build_db=not args.skip_db_build,
    )
    rpa.run()


if __name__ == "__main__":
    main()
