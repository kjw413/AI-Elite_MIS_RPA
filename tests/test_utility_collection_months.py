from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import date, datetime, timedelta
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import energy_builder  # noqa: E402
import utility_daily_rpa as utility_rpa  # noqa: E402
from utility_daily_rpa import (  # noqa: E402
    MISUtilityRPA,
    load_utility_collected_dates,
    plan_utility_collection_months,
)


def _days(start: date, end: date) -> set[date]:
    return {
        start + timedelta(days=offset)
        for offset in range((end - start).days + 1)
    }


class UtilityCollectionMonthTests(unittest.TestCase):
    def test_adds_previous_month_when_previous_day_is_missing(self) -> None:
        collected = _days(date(2026, 7, 1), date(2026, 7, 30))
        collected |= _days(date(2026, 8, 1), date(2026, 8, 2))

        self.assertEqual(
            plan_utility_collection_months(
                date(2026, 8, 2),
                {"남양주1": collected, "김해": collected},
            ),
            ["2026-07", "2026-08"],
        )

    def test_adds_previous_month_when_any_factory_is_missing(self) -> None:
        complete = _days(date(2026, 7, 1), date(2026, 7, 31))
        incomplete = _days(date(2026, 7, 1), date(2026, 7, 30))

        self.assertEqual(
            plan_utility_collection_months(
                date(2026, 8, 2),
                {"남양주1": complete, "김해": incomplete},
            ),
            ["2026-07", "2026-08"],
        )

    def test_keeps_current_month_when_previous_day_is_complete(self) -> None:
        complete = _days(date(2026, 7, 1), date(2026, 7, 31))

        self.assertEqual(
            plan_utility_collection_months(
                date(2026, 8, 2),
                {"남양주1": complete, "김해": complete},
            ),
            ["2026-08"],
        )

    def test_empty_history_keeps_current_month_behavior(self) -> None:
        self.assertEqual(
            plan_utility_collection_months(date(2026, 8, 2), {}),
            ["2026-08"],
        )

    def test_loader_requires_real_utility_value_but_accepts_zero(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "energy.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "남양주1"
            ws.append([
                energy_builder.RAW_DATE_HEADER,
                energy_builder.FIELD_BY_KEY["total_power_kwh"].label,
                energy_builder.FIELD_BY_KEY["mix_prod_kg"].label,
            ])
            ws.append([date(2026, 7, 30), 0, 100])
            ws.append([date(2026, 7, 31), None, 100])
            wb.save(path)

            loaded = load_utility_collected_dates(path, ["남양주1", "김해"])

        self.assertEqual(
            loaded,
            {
                "남양주1": {date(2026, 7, 30)},
                "김해": set(),
            },
        )

    def test_default_rpa_initialization_enables_missing_date_recovery(self) -> None:
        class FixedDatetime(datetime):
            @classmethod
            def now(cls, tz=None):
                return cls(2026, 8, 3)

        collected = {
            "김해": _days(date(2026, 7, 1), date(2026, 7, 30))
        }
        with (
            patch.object(utility_rpa, "datetime", FixedDatetime),
            patch.object(
                utility_rpa,
                "load_utility_collected_dates",
                return_value=collected,
            ),
        ):
            rpa = MISUtilityRPA(dry_run=True, org_codes=["F20"])

        self.assertTrue(rpa.auto_recover_missing_dates)
        self.assertEqual(rpa.year_month, "2026-08")
        self.assertEqual(rpa.year_months, ["2026-07", "2026-08"])

    def test_explicit_month_does_not_add_automatic_recovery(self) -> None:
        with patch.object(
            utility_rpa,
            "load_utility_collected_dates",
        ) as load_history:
            rpa = MISUtilityRPA(
                year_month="2026-08",
                dry_run=True,
                org_codes=["F20"],
            )

        load_history.assert_not_called()
        self.assertFalse(rpa.auto_recover_missing_dates)
        self.assertEqual(rpa.year_months, ["2026-08"])

    def test_explicit_date_range_maps_to_months_and_keeps_exact_filter(self) -> None:
        with patch.object(MISUtilityRPA, "_load_coords", return_value={}):
            rpa = MISUtilityRPA(
                date_from="2026-01-15",
                date_to="2026-03-02",
                dry_run=True,
                org_codes=["F20"],
            )

        self.assertFalse(rpa.auto_recover_missing_dates)
        self.assertEqual(rpa.year_months, ["2026-01", "2026-02", "2026-03"])
        self.assertEqual(rpa.requested_date_from, date(2026, 1, 15))
        self.assertEqual(rpa.requested_date_to, date(2026, 3, 2))

    def test_collection_persists_month_before_independent_processing(self) -> None:
        class Window:
            def set_focus(self):
                pass

        rpa = MISUtilityRPA(
            year_month="2026-08",
            dry_run=False,
            org_codes=["F20"],
        )
        rpa.main_window = Window()
        month_records = {
            "김해": {
                date(2026, 8, 1): {"total_power_kwh": 0}
            }
        }

        def fake_collect_months(screen, parser, months, on_month_done=None):
            on_month_done("2026-08", month_records)
            return month_records

        with (
            patch.object(rpa, "_validate_coords"),
            patch.object(rpa, "connect_mis"),
            patch.object(rpa, "collect_months", side_effect=fake_collect_months),
            patch.object(rpa, "_backup"),
            patch.object(utility_rpa.time, "sleep", return_value=None),
            patch.object(energy_builder, "write_collected_raw") as write_collected,
            patch.object(
                energy_builder,
                "process_collected_raw",
                return_value=(1, {"김해": {
                    "updated": 0, "unchanged": 1, "missing": 0,
                }}, []),
            ) as process_collected,
        ):
            self.assertTrue(rpa.collect())
            write_collected.assert_called_once_with(month_records)
            self.assertEqual(rpa.collected_months, ["2026-08"])

            self.assertTrue(rpa.process_collected_data())
            process_collected.assert_called_once()


if __name__ == "__main__":
    unittest.main()
