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


class _StubProductionRPA(MISProductionRPA):
    def __init__(self, *, consolidation_succeeds: bool = True) -> None:
        self.requested_end_date = date(2026, 8, 2)
        self.auto_recover_missing_dates = True
        self.dry_run = False
        self.build_dw = False
        self.dw_output = None
        self.main_window = _DummyWindow()
        self.collected_periods: list[tuple[str, str]] = []
        self.consolidation_calls = 0
        self.consolidation_succeeds = consolidation_succeeds
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

    def consolidate_to_dw(self) -> bool:
        self.consolidation_calls += 1
        return self.consolidation_succeeds


class ProductionCollectionFlowTests(unittest.TestCase):
    def test_integrates_recovery_before_current_period_overwrites_raw(self) -> None:
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
            rpa.run()

        self.assertEqual(
            rpa.collected_periods,
            [
                ("2026-07-31", "2026-07-31"),
                ("2026-08-01", "2026-08-02"),
            ],
        )
        self.assertEqual(rpa.consolidation_calls, 1)

    def test_stops_before_overwrite_when_recovery_integration_fails(self) -> None:
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
            self.assertRaises(RuntimeError),
        ):
            rpa.run()

        self.assertEqual(
            rpa.collected_periods,
            [("2026-07-31", "2026-07-31")],
        )


if __name__ == "__main__":
    unittest.main()
