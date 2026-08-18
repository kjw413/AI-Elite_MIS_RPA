from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import production_daily_rpa as production_rpa  # noqa: E402
from production_daily_rpa import (  # noqa: E402
    MISProductionRPA,
    load_collected_dates_by_sheet,
    plan_collection_periods,
)


def _days(start: date, end: date) -> set[date]:
    return {
        start + timedelta(days=offset)
        for offset in range((end - start).days + 1)
    }


class ProductionCollectionPeriodTests(unittest.TestCase):
    def test_recovers_gap_before_current_month_even_if_newer_dates_exist(self) -> None:
        collected = _days(date(2026, 7, 1), date(2026, 7, 30))
        collected |= _days(date(2026, 8, 1), date(2026, 8, 2))

        periods = plan_collection_periods(
            date(2026, 8, 2),
            {"F10_냉동": collected, "F20_냉동": collected},
        )

        self.assertEqual(
            periods,
            [
                (date(2026, 7, 31), date(2026, 7, 31)),
                (date(2026, 8, 1), date(2026, 8, 2)),
            ],
        )

    def test_uses_earliest_missing_start_across_target_sheets(self) -> None:
        periods = plan_collection_periods(
            date(2026, 8, 2),
            {
                "F10_냉동": _days(date(2026, 7, 1), date(2026, 7, 29)),
                "F20_냉동": _days(date(2026, 7, 1), date(2026, 7, 30)),
            },
        )

        self.assertEqual(
            periods[0],
            (date(2026, 7, 30), date(2026, 7, 31)),
        )

    def test_does_not_add_recovery_when_previous_day_is_complete(self) -> None:
        collected = _days(date(2026, 7, 1), date(2026, 7, 31))

        self.assertEqual(
            plan_collection_periods(
                date(2026, 8, 2),
                {"F10_냉동": collected, "F20_냉동": collected},
            ),
            [(date(2026, 8, 1), date(2026, 8, 2))],
        )

    def test_empty_history_keeps_existing_current_month_behavior(self) -> None:
        self.assertEqual(
            plan_collection_periods(date(2026, 8, 2), {}),
            [(date(2026, 8, 1), date(2026, 8, 2))],
        )

    def test_explicit_range_is_split_into_month_safe_periods(self) -> None:
        rpa = object.__new__(MISProductionRPA)
        rpa.auto_recover_missing_dates = False
        rpa.requested_start_date = date(2026, 1, 15)
        rpa.requested_end_date = date(2026, 3, 2)

        self.assertEqual(
            rpa._plan_collection_periods([]),
            [
                (date(2026, 1, 15), date(2026, 1, 31)),
                (date(2026, 2, 1), date(2026, 2, 28)),
                (date(2026, 3, 1), date(2026, 3, 2)),
            ],
        )

    def test_loads_dates_for_each_output_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "production.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "F10_냉동"
            ws.append(["날짜", "111100"])
            ws.append([date(2026, 7, 30), 10])
            ws.append([date(2026, 7, 31), 20])
            ws2 = wb.create_sheet("F20_냉동")
            ws2.append(["날짜", "222200"])
            ws2.append(["2026-07-30", 0])
            ws2.append(["2026-07-31", None])
            wb.save(path)

            loaded = load_collected_dates_by_sheet(
                path,
                ["F10_냉동", "F20_냉동"],
            )

        self.assertEqual(
            loaded,
            {
                "F10_냉동": {date(2026, 7, 30), date(2026, 7, 31)},
                "F20_냉동": {date(2026, 7, 30)},
            },
        )


class _DummyWindow:
    def set_focus(self) -> None:
        pass


class _GridGuardRPA(MISProductionRPA):
    """MIS 없이 생산 그리드 중복 감지·재조회 흐름을 검증하는 스텁."""

    def __init__(self, grids: dict[tuple[str, str], list[str]]) -> None:
        self.grids = {key: list(values) for key, values in grids.items()}
        self.current_factory = ""
        self.current_category = ""
        self.copy_counts: dict[tuple[str, str], int] = {}
        self.query_counts: dict[tuple[str, str], int] = {}
        self.start_date = "2026-08-01"
        self.end_date = "2026-08-11"
        self.dry_run = True
        self.main_window = _DummyWindow()
        self.duplicate_grids: list[tuple[str, str, str, str]] = []

    def select_factory(self, factory: str) -> None:
        self.current_factory = factory

    def select_category(self, category: str) -> None:
        self.current_category = category

    def click_query(self) -> None:
        key = (self.current_factory, self.current_category)
        self.query_counts[key] = self.query_counts.get(key, 0) + 1

    def copy_grid_data(self) -> str:
        key = (self.current_factory, self.current_category)
        self.copy_counts[key] = self.copy_counts.get(key, 0) + 1
        values = self.grids[key]
        if len(values) > 1:
            return values.pop(0)
        return values[0]


def _production_grid(item_code: str, value: int) -> str:
    return (
        "Item Code,Item 명,물품대,누계계획,누계실적,누계진척율,01일\n"
        f'"{item_code}","품목-{item_code}","1",'
        f'"{value}","{value}","100.00","{value}"'
    )



class _StubProductionRPA(MISProductionRPA):
    def __init__(self, *, consolidation_succeeds: bool = True) -> None:
        self.requested_end_date = date(2026, 8, 2)
        self.auto_recover_missing_dates = True
        self.dry_run = False
        self.build_dw = False
        self.dw_output = None
        self.factory_codes = ["F10"]
        self.main_window = _DummyWindow()
        self.collected_periods: list[tuple[str, str]] = []
        self.consolidation_calls = 0
        self.consolidation_succeeds = consolidation_succeeds
        self._pending_raw_snapshots = []
        self.collection_success_count = 0
        self._set_collection_period(date(2026, 8, 1), date(2026, 8, 2))

    def _plan_collection_periods(self, targets):
        return [
            (date(2026, 7, 31), date(2026, 7, 31)),
            (date(2026, 8, 1), date(2026, 8, 2)),
        ]

    def connect_mis(self) -> None:
        pass

    def navigate_to_production_screen(self) -> None:
        pass

    def backup_raw_file(self) -> None:
        pass

    def set_date_range(self) -> None:
        pass

    def select_item_type(self) -> None:
        pass

    def _collect_period(self, targets):
        self.collected_periods.append((self.start_date, self.end_date))
        return 1, 0, 1

    def _capture_raw_snapshot(self) -> None:
        self._pending_raw_snapshots.append(
            (self.start_date, self.end_date, self.start_date.encode())
        )

    def consolidate_to_dw(self, raw_path=None) -> bool:
        self.consolidation_calls += 1
        return self.consolidation_succeeds


class ProductionCollectionFlowTests(unittest.TestCase):
    def test_requeries_when_previous_factory_grid_is_copied(self) -> None:
        f30_grid = _production_grid("130294", 100)
        f40_grid = _production_grid("110388", 200)
        rpa = _GridGuardRPA({
            ("F30", "상온"): [f30_grid],
            # 첫 복사는 F30 stale 데이터, 재조회 후 F40 정상 데이터
            ("F40", "냉동"): [f30_grid, f40_grid],
        })
        targets = [
            {
                "sheet_name": "F30_상온",
                "factory": "F30",
                "category": "상온",
                "suffix": "",
            },
            {
                "sheet_name": "F40_냉동",
                "factory": "F40",
                "category": "냉동",
                "suffix": "",
            },
        ]

        with patch.object(production_rpa.time, "sleep", return_value=None):
            success, failed, written = rpa._collect_period(targets)

        self.assertEqual((success, failed, written), (2, 0, 0))
        self.assertEqual(rpa.copy_counts[("F40", "냉동")], 2)
        self.assertEqual(rpa.query_counts[("F40", "냉동")], 2)
        self.assertEqual(rpa.duplicate_grids, [])

    def test_rejects_persistent_previous_factory_grid(self) -> None:
        f30_grid = _production_grid("130294", 100)
        rpa = _GridGuardRPA({
            ("F30", "상온"): [f30_grid],
            ("F40", "냉동"): [f30_grid, f30_grid, f30_grid],
        })
        targets = [
            {
                "sheet_name": "F30_상온",
                "factory": "F30",
                "category": "상온",
                "suffix": "",
            },
            {
                "sheet_name": "F40_냉동",
                "factory": "F40",
                "category": "냉동",
                "suffix": "",
            },
        ]

        with patch.object(production_rpa.time, "sleep", return_value=None):
            success, failed, written = rpa._collect_period(targets)

        self.assertEqual((success, failed, written), (1, 1, 0))
        self.assertEqual(rpa.copy_counts[("F40", "냉동")], 3)
        self.assertEqual(
            rpa.duplicate_grids,
            [("2026-08-01", "2026-08-11", "F40_냉동", "F30_상온")],
        )
        self.assertTrue(rpa.report_duplicate_grids())

    def test_collects_all_periods_before_processing_snapshots(self) -> None:
        rpa = _StubProductionRPA()
        targets = [
            {
                "sheet_name": "F10_냉동",
                "factory": "F10",
                "category": "냉동",
                "suffix": "",
            }
        ]

        with (
            patch.object(production_rpa, "discover_targets", return_value=targets),
            patch.object(production_rpa.time, "sleep", return_value=None),
        ):
            collected = rpa.collect()

        self.assertTrue(collected)
        self.assertEqual(
            rpa.collected_periods,
            [
                ("2026-07-31", "2026-07-31"),
                ("2026-08-01", "2026-08-02"),
            ],
        )
        self.assertEqual(rpa.consolidation_calls, 0)
        self.assertEqual(len(rpa._pending_raw_snapshots), 2)

        self.assertTrue(rpa.process_collected_data())
        self.assertEqual(rpa.consolidation_calls, 2)

    def test_processing_failure_happens_after_all_collection_is_preserved(self) -> None:
        rpa = _StubProductionRPA(consolidation_succeeds=False)
        targets = [
            {
                "sheet_name": "F10_냉동",
                "factory": "F10",
                "category": "냉동",
                "suffix": "",
            }
        ]

        with (
            patch.object(production_rpa, "discover_targets", return_value=targets),
            patch.object(production_rpa.time, "sleep", return_value=None),
        ):
            self.assertTrue(rpa.collect())

        self.assertEqual(
            rpa.collected_periods,
            [
                ("2026-07-31", "2026-07-31"),
                ("2026-08-01", "2026-08-02"),
            ],
        )
        self.assertFalse(rpa.process_collected_data())
        self.assertEqual(len(rpa._pending_raw_snapshots), 2)


if __name__ == "__main__":
    unittest.main()
