# 이 스크립트는 단일 Raw 파일(RawDB_생산실적.xlsx)의 카테고리 시트들을 읽어
# DB_생산실적.xlsx (공장별 wide + 제품마스터 + 계획 + daily)를 만듭니다.
#
# 사용 예:
#   python tools/mis_rpa/build_production_dataset.py
#   python tools/mis_rpa/build_production_dataset.py \
#       --raw "E:/Sampled DB/RawDB_생산실적.xlsx" \
#       --out "E:/Sampled DB/DB_생산실적.xlsx"
from __future__ import annotations

import argparse
import logging
import sys
import time
from pathlib import Path

# Windows cp949 콘솔에서 한글/특수문자(— · ✓) 출력 가능하게
try:
    sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    sys.stderr.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
except Exception:
    pass

# tools/mis_rpa/file.py → 2단계 위가 프로젝트 루트
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from mis_rpa.production_builder import (  # noqa: E402
    DEFAULT_OUTPUT_PATH,
    DEFAULT_RAW_PATH,
    build_dataset,
    validate_subcategory_coverage,
)


def _setup_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


def main() -> int:
    p = argparse.ArgumentParser(description="RawDB_생산실적.xlsx → DB_생산실적.xlsx 빌드")
    p.add_argument("--raw", type=str, default=str(DEFAULT_RAW_PATH),
                   help=f"입력 Raw 파일 (기본: {DEFAULT_RAW_PATH})")
    p.add_argument("--out", type=str, default=str(DEFAULT_OUTPUT_PATH),
                   help=f"출력 파일 (기본: {DEFAULT_OUTPUT_PATH})")
    p.add_argument("-v", "--verbose", action="store_true")
    args = p.parse_args()

    _setup_logging(args.verbose)

    t0 = time.time()
    df, out_path = build_dataset(raw_path=args.raw, output_path=args.out)
    dt = time.time() - t0

    print(f"\n완료 — {dt:.1f}s")
    print(f"  rows         : {len(df):,}")
    if not df.empty:
        print(f"  date range   : {df['date'].min()} ~ {df['date'].max()}")
        print(f"  factories    : {sorted(df['factory'].unique())}")
        print(f"  category1    : {sorted(df['category1'].unique())}  (보관유형)")
        c2_vals = sorted([s for s in df['category2'].dropna().unique()])
        print(f"  category2    : {c2_vals if c2_vals else '(없음)'}  (제품유형: IC=아이스크림, MY=유음료, FM=발효유, SN=스낵)")
        print(f"  unique items : {df['item_code'].nunique():,}")

        # category2 커버리지 리포트
        cov = validate_subcategory_coverage(df)
        if cov["total_actual"] > 0:
            print(f"\n  [category2 분류 검증]")
            print(f"    총 실적          : {cov['total_actual']:>15,.0f}")
            print(f"    분류된 합        : {cov['classified_actual']:>15,.0f}  ({cov['coverage_pct']:.1f}%)")
            print(f"    미분류 합        : {cov['unclassified_actual']:>15,.0f}  ({cov['unclassified_pct']:.1f}%)")
            if not cov["is_complete"]:
                print(f"    ⚠ 미분류 품목 {len(cov['unclassified_items'])}개 — TOP 5:")
                for _, row in cov["unclassified_items"].head(5).iterrows():
                    print(f"      · {row['item_code']:>10} {row['item_name']:30s} {row['actual_qty']:>12,.0f}")
                print(f"    → production_dw_service._CATEGORY2_KEYWORDS 보강 또는 'category2_unclassified' 시트 참고")
            else:
                print(f"    ✓ 모든 품목이 제품유형(category2)으로 분류됨")
    print(f"\n  output       : {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
