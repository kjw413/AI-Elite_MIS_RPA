# 에너지 수집·재가공 파이프라인 검증 (MIS 화면 변경 2026-07 대응).
#
# MIS/E: 드라이브 없이 임시 fixture 로
#   utility_daily_rpa.parse_unit_input_clipboard  (신규 화면 클립보드 파싱)
#   utility_daily_rpa.parse_usage_trend_clipboard (구 화면 보완 파싱)
#   energy_builder.write_raw / read_raw / build_dataset / migrate_legacy_rawdb
# 를 검증한다.
#
# 실행: python tests/test_energy_builder.py
from __future__ import annotations

import sys
import tempfile
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook

try:
    sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    sys.stderr.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
except Exception:
    pass

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import energy_builder as eb  # noqa: E402
import utility_daily_rpa as rpa  # noqa: E402

# 2026-07 남양주1 실측 발췌 (E:\DB_MIS\RawDB_에너지_new.xlsx Sheet1).
# 컬럼: 일자, 냉동전력, 공압기, 전력사용량, 전력비, 전력단가,
#       연료사용량, 연료비, 연료단가, 용수, 폐수, 원수COD, 배출수COD
_SAMPLE_DAYS = [
    (1, 9393, 2435, 35136, 6823950.9, 194.22, 2614, 2496056.32, 954.88, 652, 505, 520, 10),
    (2, 8984, 2458, 34128, 6607428.1, 193.61, 2435, 2325132.8, 954.88, 610, 484, 480, 11),
    (3, 9252, 2815, 34740, 6786505.7, 195.35, 2340, 2234419.2, 954.88, 674, 534, 540, 12),
]
_TOTAL = ("TOTAL", 254784, 62321, 907669, 170117726.6, 185.65,
          58516, 55875758.08, 954.88, 15467, 12116, 17960, 287)
_TRAILING_BLANKS = 5   # MIS 그리드가 뒤에 붙여 내보내는 빈 열 5개


def _cells(row) -> list[str]:
    return [str(v) for v in row] + [" "] * _TRAILING_BLANKS


def _unit_input_clipboard(swap_cost_price: bool = False) -> str:
    """신규 화면 클립보드 텍스트를 재현한다 (머리글 빈 행 + 일자 행 + TOTAL)."""
    rows: list[list[str]] = [[""] * (13 + _TRAILING_BLANKS)]   # MIS 는 첫 행을 비워 내보낸다
    for day in _SAMPLE_DAYS:
        day = list(day)
        if swap_cost_price:
            day[4], day[5] = day[5], day[4]     # 전력비 ↔ 전력단가
            day[7], day[8] = day[8], day[7]     # 연료비 ↔ 연료단가
        rows.append(_cells(day))
    # 아직 실적이 없는 미래 일자 (일 번호만 있고 값은 공백)
    for day_no in (29, 30, 31):
        rows.append([str(day_no)] + [" "] * (12 + _TRAILING_BLANKS))
    rows.append(_cells(_TOTAL))
    return "\n".join(",".join(c for c in row) for row in rows)


def _usage_trend_clipboard() -> str:
    """구 화면 클립보드(행=항목, 열=일자) 텍스트를 재현한다."""
    header = ["", "01일", "02일", "03일"]
    rows = [
        ["냉동전력량", "9393", "8984", "9252"],
        ["공압기", "2435", "2458", "2815"],
        ["전력량", "35136", "34128", "34740"],
        ["연료량", "2614", "2435", "2340"],
        ["용수량", "652", "610", "674"],
        ["폐수량", "505", "484", "534"],
        ["믹스생산량", "362122", "61586", "31795"],
        ["전력원단위", "97.03", "554.14", "1092.62"],
        ["연료원단위", "7.219", "39.538", "73.596"],
        ["용수원단위", "1.8", "9.905", "21.198"],
    ]
    return "\n".join("\t".join(r) for r in [header] + rows)


def test_parse_unit_input() -> None:
    parsed = rpa.parse_unit_input_clipboard(_unit_input_clipboard(), "2026-07")

    assert set(parsed) == {date(2026, 7, 1), date(2026, 7, 2), date(2026, 7, 3)}, \
        f"일자 파싱 실패: {sorted(parsed)}"

    day1 = parsed[date(2026, 7, 1)]
    assert day1["freezing_power_kwh"] == 9393
    assert day1["air_compressor_kwh"] == 2435
    assert day1["total_power_kwh"] == 35136
    assert day1["power_cost_krw"] == 6823950.9, day1
    assert day1["power_price_krw_kwh"] == 194.22, day1
    assert day1["fuel_nm3"] == 2614
    assert day1["fuel_cost_krw"] == 2496056.32, day1
    assert day1["fuel_price_krw_nm3"] == 954.88, day1
    assert day1["water_ton"] == 652
    assert day1["wastewater_ton"] == 505
    assert day1["influent_cod_ppm"] == 520
    assert day1["effluent_cod_ppm"] == 10
    print("  ✓ 신규 화면 파싱 — TOTAL/빈 일자 제외, 13개 항목 매핑")


def test_cost_price_autodetect() -> None:
    """화면 머리글 순서('사용량→단가→비용')로 나와도 올바르게 배치돼야 한다."""
    parsed = rpa.parse_unit_input_clipboard(
        _unit_input_clipboard(swap_cost_price=True), "2026-07"
    )
    day1 = parsed[date(2026, 7, 1)]
    assert day1["power_cost_krw"] == 6823950.9, day1
    assert day1["power_price_krw_kwh"] == 194.22, day1
    assert day1["fuel_cost_krw"] == 2496056.32, day1
    assert day1["fuel_price_krw_nm3"] == 954.88, day1
    print("  ✓ 비용/단가 열 자동 판별 — 순서가 뒤바뀌어도 교정")


def test_parse_usage_trend() -> None:
    parsed = rpa.parse_usage_trend_clipboard(_usage_trend_clipboard(), "2026-07")
    day1 = parsed[date(2026, 7, 1)]
    assert set(day1) == {"mix_prod_kg", "power_per_ton_kwh",
                         "fuel_per_ton_nm3", "water_per_ton_ton"}, day1
    assert day1["mix_prod_kg"] == 362122
    assert day1["power_per_ton_kwh"] == 97.03
    print("  ✓ 구 화면 파싱 — 믹스생산량·원단위만 추출")


def _legacy_workbook(path: Path) -> None:
    """구 형식 RawDB_에너지(행=항목, 열=날짜) fixture 를 만든다."""
    wb = Workbook()
    wb.remove(wb.active)
    for sheet_name in ("남양주1", "김해"):
        ws = wb.create_sheet(sheet_name)
        ws.cell(row=1, column=1, value="날짜")
        for offset, field in enumerate(eb.FIELDS[:10]):     # 구 파일은 10개 항목뿐
            ws.cell(row=2 + offset, column=1, value=field.label)
        ws.cell(row=1, column=2, value=date(2026, 6, 30))
        for offset in range(10):
            ws.cell(row=2 + offset, column=2, value=offset + 1)
    wb.save(path)


def test_migrate_and_build() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        raw_path = tmp_dir / "RawDB_에너지.xlsx"
        db_path = tmp_dir / "DB_에너지.xlsx"

        # ── 1. 구 형식 이관 ──
        _legacy_workbook(raw_path)
        assert eb.is_legacy_transposed(raw_path)
        assert eb.migrate_legacy_rawdb(raw_path, db_path) is True
        assert db_path.exists(), "DB_에너지.xlsx 가 생성되지 않음"
        assert not raw_path.exists(), "구 RawDB 가 backup 으로 이동되지 않음"
        assert list((tmp_dir / "backup").glob("RawDB_에너지_legacy_*.xlsx")), "백업 없음"
        assert eb.migrate_legacy_rawdb(raw_path, db_path) is False, "이관은 1회만"
        print("  ✓ 구 형식 이관 — RawDB(전치형) → DB_에너지, 원본은 backup/ 보관")

        # ── 2. RawDB 적재 ──
        records = rpa.parse_unit_input_clipboard(_unit_input_clipboard(), "2026-07")
        trend = rpa.parse_usage_trend_clipboard(_usage_trend_clipboard(), "2026-07")
        collected = {"남양주1": records}
        rpa.MISUtilityRPA._merge(collected, {"남양주1": trend})
        eb.write_raw(collected, raw_path)

        round_tripped = eb.read_raw(raw_path)
        assert set(round_tripped) == {"남양주1"}
        day1 = round_tripped["남양주1"][date(2026, 7, 1)]
        assert day1["power_cost_krw"] == 6823950.9
        assert day1["mix_prod_kg"] == 362122
        assert len(day1) == len(eb.RAW_COLUMN_KEYS), f"항목 누락: {sorted(day1)}"
        print("  ✓ RawDB 적재/재읽기 — 16개 항목 왕복 일치")

        # 같은 데이터 재적재 시 행이 늘지 않아야 한다 (upsert)
        eb.write_raw(collected, raw_path)
        wb = load_workbook(raw_path)
        assert wb["남양주1"].max_row == 1 + 3, wb["남양주1"].max_row
        wb.close()
        print("  ✓ RawDB upsert — 재실행해도 행 중복 없음")

        # ── 3. DB 재가공 ──
        stats, out = eb.build_dataset(raw_path, db_path)
        assert stats["남양주1"]["appended"] == 3, stats
        wb = load_workbook(db_path)
        ws = wb["남양주1"]

        labels = [ws.cell(row=2 + i, column=1).value for i in range(len(eb.FIELDS))]
        assert labels == [f.label for f in eb.FIELDS], labels

        # 이관된 과거 열(6/30)은 그대로 남아 있어야 한다
        assert ws.cell(row=1, column=2).value.date() == date(2026, 6, 30)
        assert ws.cell(row=2, column=2).value == 1

        # 신규 열(7/1) — 구 항목 + 신규 항목이 모두 채워졌는지
        assert ws.cell(row=1, column=3).value.date() == date(2026, 7, 1)
        by_label = {ws.cell(row=2 + i, column=1).value: ws.cell(row=2 + i, column=3).value
                    for i in range(len(eb.FIELDS))}
        assert by_label["전력량[kWh]"] == 35136
        assert by_label["믹스생산량[kg]"] == 362122
        assert by_label["전력비[원]"] == 6823950.9
        assert by_label["전력단가[원/kWh]"] == 194.22
        assert by_label["배출수COD[ppm]"] == 10
        wb.close()
        print("  ✓ DB 재가공 — 과거 열 보존 + 신규 항목행(12~17) 추가")

        # ── 4. 재빌드 idempotent ──
        stats2, _ = eb.build_dataset(raw_path, db_path)
        assert stats2["남양주1"] == {"appended": 0, "overwritten": 3}, stats2
        print("  ✓ DB 재빌드 — 같은 날짜는 덮어쓰기(열 증가 없음)")


def main() -> int:
    print("에너지 파이프라인 검증")
    for fn in (test_parse_unit_input, test_cost_price_autodetect,
               test_parse_usage_trend, test_migrate_and_build):
        fn()
    print("\n전체 통과")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
