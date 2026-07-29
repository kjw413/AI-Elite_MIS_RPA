# -*- coding: utf-8 -*-
"""
MIS 3종 RPA 자동 실행 오케스트레이터 (in-process 버전).

이전(subprocess) 구조 대비 변경점:
  - 3개 UI RPA 를 같은 Python 프로세스 안에서 순차 실행.
  - MIS 연결(connect_mis)은 생산실적 단계에서만 1회 수행.
    이후 유틸리티/재공품은 그 main_window 를 attach_existing_window() 로
    재사용해 ~9초의 UIA 연결 오버헤드를 건너뛴다.
  - DW/DB 통합은 여전히 subprocess 백그라운드(다음 UI 단계와 병렬).

Pipeline:
    prod_UI ─→ util_UI ─→ wip_UI ─→ wait
            └ prod_DW(BG)        └ wip_DB(BG)

Usage:
    python run_all_rpa.py [--date YYYY-MM-DD] [--dry-run]
"""
from __future__ import annotations

import argparse
import io
import logging
import subprocess
import sys
from datetime import datetime
from pathlib import Path

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
# 백그라운드 잡 (DW/DB 통합 — UI 와 무관한 파일·DB I/O)
# ---------------------------------------------------------------------------
def start_bg(title: str, args: list[str], bg_log_path: Path):
    log.info(f"[BG start] {title} → {bg_log_path.name}")
    cmd = [sys.executable, "-u", *args]
    fh = bg_log_path.open("w", encoding="utf-8")
    fh.write(f"$ {' '.join(cmd)}\n")
    fh.flush()
    proc = subprocess.Popen(
        cmd,
        cwd=str(SCRIPT_DIR),
        stdout=fh,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    return title, proc, fh


def wait_bg(jobs: list) -> dict[str, int]:
    results: dict[str, int] = {}
    for title, proc, fh in jobs:
        rc = proc.wait()
        fh.close()
        results[title] = rc
        log.info(f"[BG done ] {title} 종료 (exit={rc}) — {Path(fh.name).name}")
        bg_path = Path(fh.name)
        if bg_path.exists():
            try:
                content = bg_path.read_text(encoding="utf-8", errors="replace")
                with MAIN_LOG.open("a", encoding="utf-8") as mf:
                    mf.write(f"\n----- BG output: {title} ({bg_path.name}) -----\n")
                    mf.write(content)
                    mf.write(f"----- end BG: {title} -----\n")
            except Exception as exc:  # noqa: BLE001
                log.warning(f"BG 로그 머지 실패: {exc}")
    return results


# ---------------------------------------------------------------------------
# UI 단계 실행 헬퍼 (예외/SystemExit 모두 잡아서 rc 만 반환)
# ---------------------------------------------------------------------------
def _run_rpa_safe(title: str, runner) -> int:
    """RPA 인스턴스의 run() 을 안전하게 호출. 실패해도 다음 단계 진행."""
    try:
        runner()
        return 0
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
    ap.add_argument("--dry-run", action="store_true",
                    help="MIS 조회만 — Excel/DB 미기록.")
    args, _unknown = ap.parse_known_args()

    header(f"MIS 3종 RPA 자동 실행 — 시작 {datetime.now():%Y-%m-%d %H:%M:%S}")
    log.info(f"메인 로그: {MAIN_LOG}")
    log.info(f"공통 인자: --date={args.date}  --dry-run={args.dry_run}")

    # RPA 모듈 import — 이 시점에 핸들러는 이미 root 에 부착돼있음
    from production_daily_rpa import MISProductionRPA
    from utility_daily_rpa import MISUtilityRPA
    from wip_daily_rpa import MISWIPRPA

    bg_jobs: list = []

    # ──────────────────────────────────────────────────────────
    # [1/3] 생산실적 UI — MIS 연결을 실제로 수행
    header("[1/3] 생산실적 RPA — MIS 화면 작업 (MIS 연결 1회 수행)")
    prod = MISProductionRPA(
        ref_date=args.date,
        dry_run=args.dry_run,
        build_dw=False,  # DW 통합은 아래서 BG 로 분리
    )
    rc_prod_ui = _run_rpa_safe("생산실적 UI", prod.run)

    shared_app = getattr(prod, "app", None)
    shared_window = getattr(prod, "main_window", None)
    can_share = shared_window is not None
    if not can_share:
        log.warning("MIS 윈도우 객체를 얻지 못함 — 후속 단계는 각자 재연결을 시도합니다.")
    else:
        log.info("✓ MIS 연결 객체 확보 — 유틸리티/재공품 단계는 connect 건너뜀.")

    # 생산실적 DW 통합을 BG 로 — 다음 UI 단계와 병렬
    if rc_prod_ui == 0 and not args.dry_run:
        bg_jobs.append(start_bg(
            "생산실적 DW 통합",
            ["build_production_dataset.py"],
            LOG_DIR / f"auto_run_{STAMP}_prod_dw_bg.log",
        ))
        log.info("→ 생산실적 DW 통합 BG 시작.")
    else:
        log.info(f"[건너뜀] DW 통합 (rc={rc_prod_ui}, dry-run={args.dry_run})")

    # ──────────────────────────────────────────────────────────
    # [2/3] 유틸리티 UI — MIS 연결 재사용
    header("[2/3] 유틸리티 RPA — MIS 화면 작업 (연결 재사용)")
    # 유틸리티는 기준년월만 받으므로, 공통 기준 종료일의 YYYY-MM을 전달한다.
    # --date 미지정 때는 기존처럼 유틸리티가 D-1 기준으로 계산한다.
    utility_year_month = args.date[:7] if args.date else None
    util = MISUtilityRPA(year_month=utility_year_month, dry_run=args.dry_run)
    if can_share:
        util.attach_existing_window(shared_app, shared_window)
    rc_util = _run_rpa_safe("유틸리티", util.run)

    # ──────────────────────────────────────────────────────────
    # [3/3] 재공품 UI — MIS 연결 재사용
    header("[3/3] 재공품 RPA — MIS 화면 작업 (연결 재사용)")
    wip = MISWIPRPA(
        ref_date=args.date,
        dry_run=args.dry_run,
        build_db=False,  # DB 통합은 아래서 BG 로 분리
    )
    if can_share:
        wip.attach_existing_window(shared_app, shared_window)
    rc_wip_ui = _run_rpa_safe("재공품 UI", wip.run)

    if rc_wip_ui == 0 and not args.dry_run:
        wip_ref_script = SCRIPT_DIR / "wip_refactoring.py"
        bg_jobs.append(start_bg(
            "재공품 DB 통합",
            [str(wip_ref_script)],
            LOG_DIR / f"auto_run_{STAMP}_wip_db_bg.log",
        ))
        log.info("→ 재공품 DB 통합 BG 시작.")
    else:
        log.info(f"[건너뜀] DB 통합 (rc={rc_wip_ui}, dry-run={args.dry_run})")

    # ──────────────────────────────────────────────────────────
    # BG 완료 대기
    rc_prod_dw = None
    rc_wip_db = None
    if bg_jobs:
        header("백그라운드 통합 작업 완료 대기...")
        results = wait_bg(bg_jobs)
        rc_prod_dw = results.get("생산실적 DW 통합")
        rc_wip_db = results.get("재공품 DB 통합")

    # ──────────────────────────────────────────────────────────
    # 요약
    def _fmt(rc) -> str:
        return "생략" if rc is None else str(rc)

    summary = (
        "\n============================================================\n"
        "  실행 결과 요약  [0 = 성공]\n"
        f"    [1] 생산실적 UI       : {rc_prod_ui}\n"
        f"    [2] 유틸리티          : {rc_util}\n"
        f"    [3] 재공품 UI         : {rc_wip_ui}\n"
        f"    [BG] 생산실적 DW 통합 : {_fmt(rc_prod_dw)}\n"
        f"    [BG] 재공품 DB 통합   : {_fmt(rc_wip_db)}\n"
        f"  종료: {datetime.now():%Y-%m-%d %H:%M:%S}\n"
        f"  로그: {MAIN_LOG}\n"
        "============================================================\n"
    )
    log.info(summary)

    all_rcs = [rc_prod_ui, rc_util, rc_wip_ui, rc_prod_dw, rc_wip_db]
    return 1 if any(isinstance(rc, int) and rc != 0 for rc in all_rcs) else 0


if __name__ == "__main__":
    sys.exit(main())
