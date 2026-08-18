from __future__ import annotations

import sys
import unittest
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from _common import resolve_date_range, split_date_range_by_month  # noqa: E402


class DateRangeTests(unittest.TestCase):
    def test_blank_range_defaults_to_month_start_through_d_minus_one(self) -> None:
        self.assertEqual(
            resolve_date_range(None, None, default_end=date(2026, 8, 17)),
            (date(2026, 8, 1), date(2026, 8, 17)),
        )

    def test_start_only_defaults_end_to_d_minus_one(self) -> None:
        self.assertEqual(
            resolve_date_range(
                "2026-01-01",
                None,
                default_end=date(2026, 8, 17),
            ),
            (date(2026, 1, 1), date(2026, 8, 17)),
        )

    def test_legacy_year_month_input_remains_supported(self) -> None:
        self.assertEqual(
            resolve_date_range("2026-01", "2026-03"),
            (date(2026, 1, 1), date(2026, 3, 31)),
        )

    def test_production_range_splits_at_month_boundaries(self) -> None:
        self.assertEqual(
            split_date_range_by_month(date(2026, 1, 15), date(2026, 3, 2)),
            [
                (date(2026, 1, 15), date(2026, 1, 31)),
                (date(2026, 2, 1), date(2026, 2, 28)),
                (date(2026, 3, 1), date(2026, 3, 2)),
            ],
        )

    def test_rejects_reversed_range(self) -> None:
        with self.assertRaises(SystemExit):
            resolve_date_range("2026-03-01", "2026-02-28")


class BatchPromptTests(unittest.TestCase):
    def test_all_launchers_prompt_for_the_same_date_range(self) -> None:
        launchers = [
            "생산실적_RPA_실행.bat",
            "유틸리티_RPA_실행.bat",
            "재공품_RPA_실행.bat",
            "전체_RPA_자동실행.bat",
        ]
        for filename in launchers:
            with self.subTest(filename=filename):
                text = (ROOT / filename).read_text(encoding="utf-8")
                self.assertIn("set /p DATE_FROM=", text)
                self.assertIn("set /p DATE_TO=", text)
                self.assertIn("--from !DATE_FROM!", text)
                self.assertIn("--to !DATE_TO!", text)
                self.assertIn('if not "%~1"==""', text)


if __name__ == "__main__":
    unittest.main()
