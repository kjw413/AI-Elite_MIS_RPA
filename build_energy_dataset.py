# 이 스크립트는 RawDB_에너지.xlsx (행=일자, MIS 화면 원본)를 읽어
# DB_에너지.xlsx (행=항목, 열=날짜 — BEMS 웹앱 입력 파일)를 만듭니다.
#
# 사용 예:
#   python build_energy_dataset.py
#   python build_energy_dataset.py --raw "E:/DB_MIS/RawDB_에너지.xlsx" \
#                                  --out "E:/DB_MIS/DB_에너지.xlsx"
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

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from energy_builder import (  # noqa: E402
    DEFAULT_OUTPUT_PATH,
    DEFAULT_RAW_PATH,
    build_dataset,
    migrate_legacy_rawdb,
)


def _setup_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


def main() -> int:
    p = argparse.ArgumentParser(description="RawDB_에너지.xlsx → DB_에너지.xlsx 빌드")
    p.add_argument("--raw", type=str, default=str(DEFAULT_RAW_PATH),
                   help=f"입력 Raw 파일 (기본: {DEFAULT_RAW_PATH})")
    p.add_argument("--out", type=str, default=str(DEFAULT_OUTPUT_PATH),
                   help=f"출력 파일 (기본: {DEFAULT_OUTPUT_PATH})")
    p.add_argument("-v", "--verbose", action="store_true")
    args = p.parse_args()

    _setup_logging(args.verbose)

    # 구 형식(RawDB_에너지가 전치형) 이면 1회 이관 후 진행
    migrate_legacy_rawdb(Path(args.raw), Path(args.out))

    t0 = time.time()
    stats, out_path = build_dataset(raw_path=args.raw, output_path=args.out)
    dt = time.time() - t0

    print(f"\n완료 — {dt:.1f}s")
    for sheet_name, s in stats.items():
        print(f"  {sheet_name:8s} 신규 {s['appended']:>4}열 / 덮어쓰기 {s['overwritten']:>4}열")
    print(f"\n  output       : {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
