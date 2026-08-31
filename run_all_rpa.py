# -*- coding: utf-8 -*-
"""
MIS 3종 RPA 자동 실행 오케스트레이터 (in-process 버전).

이전(subprocess) 구조 대비 변경점:
  - 3개 UI RPA 를 같은 Python 프로세스 안에서 순차 실행.
  - MIS 연결(connect_mis)은 생산실적 단계에서만 1회 수행.
    이후 유틸리티/재공품은 그 main_window 를 attach_existing_window() 로
    재사용해 ~9초의 UIA 연결 오버헤드를 건너뛴다.
  - 1단계에서 3개 UI 수집을 모두 끝낸 뒤 2단계 통합·가공을 시작.

Pipeline:
    [수집] production → utility → wip
                         ↓
    [가공] production → wip → utility

Usage:
    python run_all_rpa.py [--date YYYY-MM-DD] [--dry-run]
"""
from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path
from _common import resolve_date_range, resolve_factory_codes

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

STAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
MAIN_LOG = LOG_DIR / f"auto_run_{STAMP}.log"

# ---------------------------------------------------------------------------
# 로깅: 콘솔 + 메인 로그파일 동시 기록.
# 각 RPA 모듈은 logging.getLogger(__name__) (propagate=True) 라서 여기 root
# 핸들러로 자동 흐른다. 또한 각 모듈의 _setup_logging() 은 root 에 핸들러가
# 이미 있으면 no-op 이라 중복 부착도 없다.
# ---------------------------------------------------------------------------
_root = logging.getLogger()
_root.setLevel(logging.INFO)
_fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
_ch = logging.StreamHandler(sys.stdout)
_ch.setFormatter(_fmt)
_fh = logging.FileHandler(MAIN_LOG, encoding="utf-8")
_fh.setFormatter(_fmt)
_root.addHandler(_ch)
_root.addHandler(_fh)

log = logging.getLogger("run_all_rpa")


def header(msg: str) -> None:
    bar = "=" * 60
    log.info("")
    log.info(bar)
    log.info(f"  {msg}")
    log.info(bar)


# ---------------------------------------------------------------------------
# UI 단계 실행 헬퍼 (예외/SystemExit 모두 잡아서 rc 만 반환)
# ---------------------------------------------------------------------------
def _run_rpa_safe(title: str, runner) -> int:
    """수집 또는 가공 단계 함수를 안전하게 호출. 실패해도 다음 단계 진행."""
    try:
        result = runner()
        return 1 if result is False else 0
    except SystemExit as exc:
        rc = int(exc.code) if isinstance(exc.code, int) else 1
        log.error(f"{title}: SystemExit (rc={rc})")
        return rc
    except Exception:  # noqa: BLE001
        log.exception(f"{title}: 예외 발생")
        return 1


# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(
        description="MIS 3종 RPA 자동 실행 (in-process, MIS 연결 1회 공유)"
    )
    ap.add_argument("--date", type=str, default=None,
                    help="기준 종료일 (YYYY-MM-DD). 미지정 시 D-1 자동.")
    ap.add_argument("--from", dest="date_from", default=None,
                    help="공통 조회 시작일 (YYYY-MM-DD). 미지정 시 종료월 1일.")
    ap.add_argument("--to", dest="date_to", default=None,
                    help="공통 조회 종료일 (YYYY-MM-DD). 미지정 시 어제.")
    ap.add_argument(
        "--factories", default=None,
        help="공통 대상 공장 코드/공장명 CSV (예: '광주' / 'F30'). 기본: 전체",
    )
    ap.add_argument("--dry-run", action="store_true",
                    help="MIS 조회만 — Excel/DB 미기록.")
    args, _unknown = ap.parse_known_args()

    if args.date and (args.date_from or args.date_to):
        raise SystemExit("--date 는 --from/--to 와 함께 사용할 수 없습니다.")
    common_from = common_to = None
    if not args.date:
        range_start, range_end = resolve_date_range(args.date_from, args.date_to)
        common_from = range_start.isoformat()
        common_to = range_end.isoformat()

    factory_codes = resolve_factory_codes(args.factories)
    utility_org_codes: list[str] = []
    for factory in factory_codes:
        utility_org_codes.extend(["F1A", "F1B"] if factory == "F10" else [factory])

    header(f"MIS 3종 RPA 자동 실행 — 시작 {datetime.now():%Y-%m-%d %H:%M:%S}")
    log.info(f"메인 로그: {MAIN_LOG}")
    log.info(
        f"공통 기간: {common_from or '해당 월 1일'} ~ {common_to or args.date}"
    )
    log.info(f"공통 대상 공장: {', '.join(factory_codes)}")

    # RPA 모듈 import — 이 시점에 핸들러는 이미 root 에 부착돼있음
    from production_daily_rpa import MISProductionRPA
    from utility_daily_rpa import MISUtilityRPA
    from wip_daily_rpa import MISWIPRPA

    # ==========================================================
    # 1단계: 3종 수집
    # ==========================================================
    header("1단계: 3종 RPA 수집 시작")

    # [수집 1/3] 생산실적 — MIS 연결을 실제로 수행
    header("[수집 1/3] 생산실적 (MIS 연결 1회 수행)")
    prod = MISProductionRPA(
        ref_date=args.date,
        dry_run=args.dry_run,
        build_dw=False,
        date_from=common_from,
        date_to=common_to,
        factory_codes=factory_codes,
    )
    rc_prod_collect = _run_rpa_safe("생산실적 수집", prod.collect)

    shared_app = getattr(prod, "app", None)
    shared_window = getattr(prod, "main_window", None)
    can_share = shared_window is not None
    if not can_share:
        log.warning("MIS 윈도우 객체를 얻지 못함 — 후속 단계는 각자 재연결을 시도합니다.")
    else:
        log.info("✓ MIS 연결 객체 확보 — 유틸리티/재공품 단계는 connect 건너뜀.")

    # [수집 2/3] 유틸리티 — MIS 연결 재사용
    header("[수집 2/3] 유틸리티 (연결 재사용)")
    # 구 --date 는 월 단위 호환, 새 --from/--to 는 정확한 가공 범위로 전달한다.
    utility_year_month = args.date[:7] if args.date else None
    util = MISUtilityRPA(
        year_month=utility_year_month,
        date_from=common_from,
        date_to=common_to,
        dry_run=args.dry_run,
        org_codes=utility_org_codes,
    )
    if can_share:
        util.attach_existing_window(shared_app, shared_window)
    rc_util_collect = _run_rpa_safe("유틸리티 수집", util.collect)

    # [수집 3/3] 재공품 — MIS 연결 재사용
    header("[수집 3/3] 재공품 (연결 재사용)")
    wip = MISWIPRPA(
        ref_date=args.date,
        dry_run=args.dry_run,
        build_db=False,
        date_from=common_from,
        date_to=common_to,
        factory_codes=factory_codes,
    )
    if can_share:
        wip.attach_existing_window(shared_app, shared_window)
    rc_wip_collect = _run_rpa_safe("재공품 수집", wip.collect)

    header("1단계 완료: 3종 RPA 수집 시도 종료")

    # ==========================================================
    # 2단계: 통합 및 가공
    # ==========================================================
    rc_prod_process = None
    rc_util_process = None
    rc_wip_process = None

    if args.dry_run:
        header("2단계 생략: dry-run 모드")
    else:
        header("2단계: 통합 및 가공 시작")

        if rc_prod_collect == 0:
            header("[가공 1/3] 생산실적 DW 통합")
            rc_prod_process = _run_rpa_safe(
                "생산실적 가공", prod.process_collected_data
            )
        else:
            log.error("[가공 생략] 생산실적 수집 실패")

        if rc_wip_collect == 0:
            header("[가공 2/3] 재공품 DB 통합")
            rc_wip_process = _run_rpa_safe(
                "재공품 가공", wip.process_collected_data
            )
        else:
            log.error("[가공 생략] 재공품 수집 실패")

        # 유틸리티 원단위 가공은 최신 생산실적과 광주 재공품 DB를 모두 사용한다.
        if (rc_util_collect == 0 and rc_prod_process == 0
                and rc_wip_process == 0):
            header("[가공 3/3] 유틸리티 적재 및 원단위 가공")
            rc_util_process = _run_rpa_safe(
                "유틸리티 가공", util.process_collected_data
            )
        elif rc_util_collect != 0:
            log.error("[가공 생략] 유틸리티 수집 실패")
        elif rc_prod_process != 0:
            log.error("[가공 생략] 생산실적 가공 실패로 유틸리티 의존성 미충족")
        else:
            log.error("[가공 생략] 재공품 가공 실패로 유틸리티 의존성 미충족")

    # ──────────────────────────────────────────────────────────
    # 요약
    def _fmt(rc) -> str:
        return "생략" if rc is None else str(rc)

    summary = (
        "\n============================================================\n"
        "  실행 결과 요약  [0 = 성공]\n"
        f"    [수집 1] 생산실적     : {rc_prod_collect}\n"
        f"    [수집 2] 유틸리티     : {rc_util_collect}\n"
        f"    [수집 3] 재공품       : {rc_wip_collect}\n"
        f"    [가공 1] 생산실적     : {_fmt(rc_prod_process)}\n"
        f"    [가공 2] 재공품       : {_fmt(rc_wip_process)}\n"
        f"    [가공 3] 유틸리티     : {_fmt(rc_util_process)}\n"
        f"  종료: {datetime.now():%Y-%m-%d %H:%M:%S}\n"
        f"  로그: {MAIN_LOG}\n"
        "============================================================\n"
    )
    log.info(summary)

    all_rcs = [
        rc_prod_collect,
        rc_util_collect,
        rc_wip_collect,
        rc_prod_process,
        rc_util_process,
        rc_wip_process,
    ]
    return 1 if any(isinstance(rc, int) and rc != 0 for rc in all_rcs) else 0


if __name__ == "__main__":
    sys.exit(main())
