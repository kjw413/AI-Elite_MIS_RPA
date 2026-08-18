from __future__ import annotations

import unittest
from unittest.mock import patch

import wip_daily_rpa as wip_rpa
from wip_daily_rpa import MISWIPRPA


class _Window:
    def set_focus(self) -> None:
        pass


class WIPPhaseSplitTests(unittest.TestCase):
    def test_explicit_date_range_is_used_without_month_start_override(self) -> None:
        with patch.object(MISWIPRPA, "_load_coords", return_value={}):
            rpa = MISWIPRPA(
                date_from="2026-01-15",
                date_to="2026-03-02",
                dry_run=True,
            )

        self.assertEqual(rpa.start_date, "2026-01-15")
        self.assertEqual(rpa.end_date, "2026-03-02")

    def test_collection_writes_raw_but_defers_database_processing(self) -> None:
        rpa = object.__new__(MISWIPRPA)
        rpa.dry_run = False
        rpa.build_db = True
        rpa.factory_codes = ["F20"]
        rpa.main_window = _Window()
        rpa.collection_success_count = 0

        with (
            patch.object(rpa, "connect_mis"),
            patch.object(rpa, "navigate_to_wip_screen"),
            patch.object(rpa, "set_date_range"),
            patch.object(rpa, "select_actual_basis_tab"),
            patch.object(rpa, "select_category1_wip"),
            patch.object(rpa, "toggle_subtotal_off"),
            patch.object(rpa, "backup_output"),
            patch.object(rpa, "select_factory"),
            patch.object(rpa, "click_query"),
            patch.object(rpa, "copy_grid_data", return_value="grid"),
            patch.object(rpa, "consolidate_to_db", return_value=True) as consolidate,
            patch.object(wip_rpa, "parse_clipboard_rows", return_value=[["row"]]),
            patch.object(wip_rpa, "paste_to_sheet", return_value=1),
            patch.object(wip_rpa.time, "sleep", return_value=None),
        ):
            self.assertTrue(rpa.collect())
            consolidate.assert_not_called()

            self.assertTrue(rpa.process_collected_data())
            consolidate.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
