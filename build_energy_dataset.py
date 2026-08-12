# 에너지 가공 전용 CLI — 영속 수집 원본 반영 + 생산량/원단위 갱신
"""
MIS 화면을 열지 않고 `RawDB_에너지.xlsx`의 좌표 수집 결과를
`DB_에너지.xlsx`에 반영하고, 믹스생산량과 원단위 수식을 갱신한다.

수집(`utility_daily_rpa.py`, MIS 클릭)과 가공(이 스크립트, 엑셀 재집계)을 분리한 것은
두 단계의 신선도가 어긋날 수 있기 때문이다:

  - 생산실적 입력이 늦으면 에너지 수집 시점에 믹스생산량이 0 으로 굳는다
    (2026-08-06 경산 사례: 09:56 에너지 수집 → 10:51 생산실적 도착 → 0 잔류)
  - 구 '유틸리티 일자별 사용량 추이' 화면에서 받아 둔 값이 남아 있는 과거 행도 있다

이 스크립트를 생산실적 갱신 뒤에 돌리면 MIS 접속 없이 수 초 만에 정합이 맞는다.

권위 값 = `DB_생산실적.xlsx`(가공 완료본) + `RawDB_생산실적.xlsx`(최신 수집분 우선).

Usage:
  python build_energy_dataset.py                          # 전체 기간 재집계
  python build_energy_dataset.py --from 2026-08           # 2026-08 ~ 끝
  python build_energy_dataset.py --from 2024-04 --to 2024-08
  python build_energy_dataset.py --factories 경산          # 특정 공장만
  python build_energy_dataset.py --raw D:/data/RawDB_에너지.xlsx --out D:/data/DB_에너지.xlsx
  python build_energy_dataset.py --dry-run                # 변경 예정만 출력
  python build_energy_dataset.py --no-recalc              # Excel COM 재계산 생략
"""
from __future__ import annotations

import argparse
import logging
import sys
import time
from pathlib import Path

# Windows cp949 콘솔에서 한글/특수문자 출력 가능하게
try:
    sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    sys.stderr.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
except Exception:
    pass

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import energy_builder  # noqa: E402
from _common import month_bounds  # noqa: E402


def _setup_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def main() -> int:
    p = argparse.ArgumentParser(
        description="에너지 수집 원본 반영 + 믹스생산량/원단위 가공 (MIS 접속 없음)"
    )
    p.add_argument("--from", dest="ym_from", default=None,
                   help="시작 기준년월 (YYYY-MM). 미지정 시 전체 기간")
    p.add_argument("--to", dest="ym_to", default=None,
                   help="종료 기준년월 (YYYY-MM). 미지정 시 마지막 행까지")
    p.add_argument("--factories", default=None,
                   help="대상 사업장. 공장명 또는 공장코드 CSV (예: '경산' / 'F50'). "
                        "기본: 전체")
    p.add_argument("--raw", "--in", dest="raw", default=None,
                   help="입력 Raw 파일 "
                        f"(기본: {energy_builder.DEFAULT_RAW_PATH})")
    p.add_argument("--out", dest="output", default=None,
                   help=f"출력 DB 파일 (기본: {energy_builder.DEFAULT_OUTPUT_PATH})")
    p.add_argument("--dry-run", action="store_true",
                   help="파일을 저장하지 않고 변경 예정만 출력")
    p.add_argument("--no-recalc", action="store_true",
                   help="저장 후 Excel COM 수식 재계산을 생략 (엑셀 미설치 환경)")
    p.add_argument("-v", "--verbose", action="store_true")
    args = p.parse_args()

    _setup_logging(args.verbose)

    date_from, date_to = month_bounds(args.ym_from, args.ym_to)
    sheet_names = energy_builder.resolve_sheet_names(args.factories)

    period = (f"{args.ym_from or '처음'} ~ {args.ym_to or '끝'}")
    print()
    print("=" * 66)
    print("  에너지 수집 원본 반영 + 믹스생산량/원단위 가공")
    print("=" * 66)
    print(f"  수집 원본 : {args.raw or energy_builder.DEFAULT_RAW_PATH}")
    print(f"  가공 출력 : {args.output or energy_builder.DEFAULT_OUTPUT_PATH}")
    print(f"  기간      : {period}")
    print(f"  사업장    : {', '.join(sheet_names)}")
    print(f"  권위 값   : {energy_builder.DEFAULT_PRODUCTION_PATH}")
    print(f"              + {energy_builder.DEFAULT_PRODUCTION_RAW_PATH} (최신 우선)")
    if args.dry_run:
        print("  모드      : DRY-RUN (파일 미저장)")
    print("=" * 66)

    t0 = time.time()
    collected_days, stats, changes = energy_builder.process_collected_raw(
        Path(args.raw) if args.raw else None,
        Path(args.output) if args.output else None,
        date_from=date_from,
        date_to=date_to,
        sheet_names=sheet_names,
        recalculate=not args.no_recalc,
        dry_run=args.dry_run,
    )
    elapsed = time.time() - t0

    print()
    mode_text = "반영 예정" if args.dry_run else "반영"
    print(f"수집 원본 {collected_days}일 {mode_text}")
    print(f"{'공장':10s}{'갱신':>8}{'동일':>8}{'실적없음':>10}")
    for sheet_name in sheet_names:
        s = stats.get(sheet_name)
        if s is None:
            continue
        print(f"{sheet_name:10s}{s['updated']:>8}{s['unchanged']:>8}{s['missing']:>10}")

    if changes:
        print(f"\n변경 {len(changes)}건" + (" (DRY-RUN — 미적용)" if args.dry_run else ""))

        # 차이 크기로 나눠 본다. 구 '유틸리티 일자별' 화면 값이 남은 행은 생산실적과
        # ±1~3kg 만 다른 경우가 대부분이라, 그대로 나열하면 정작 중요한 큰 차이가
        # 묻힌다. 큰 것부터 보여주고 잔차는 건수로만 요약한다.
        def _diff(change) -> float:
            _sheet, _day, before, after = change
            return abs(after - (before or 0.0))

        buckets = [
            ("10,000kg 초과", lambda d: d > 10_000),
            ("1,000 ~ 10,000kg", lambda d: 1_000 < d <= 10_000),
            ("10 ~ 1,000kg", lambda d: 10 < d <= 1_000),
            ("10kg 이하 (반올림 수준)", lambda d: d <= 10),
        ]
        print()
        for label, cond in buckets:
            count = sum(1 for c in changes if cond(_diff(c)))
            if count:
                print(f"  {label:26s} {count:>6}건")

        major = sorted(changes, key=_diff, reverse=True)
        major = [c for c in major if _diff(c) > 10]
        if major:
            print(f"\n  차이 큰 순 상위 {min(len(major), 25)}건:")
            for sheet_name, day, before, after in major[:25]:
                before_text = "-" if before is None else f"{before:,.0f}"
                print(f"    {sheet_name:8s} {day}  {before_text:>13} → {after:>13,.0f}"
                      f"   (차이 {_diff((sheet_name, day, before, after)):>11,.0f})")
            if len(major) > 25:
                print(f"    … 외 {len(major) - 25}건")
    else:
        print("\n생산량 변경 없음 — 이미 생산실적과 일치합니다.")

    print(f"\n완료 — {elapsed:.1f}s")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
