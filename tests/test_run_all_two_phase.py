from __future__ import annotations

import sys
import types
import unittest
from argparse import Namespace
from unittest.mock import patch

import run_all_rpa


class _Window:
    def set_focus(self) -> None:
        pass


class RunAllTwoPhaseTests(unittest.TestCase):
    def test_all_collection_finishes_before_any_processing_starts(self) -> None:
        events: list[str] = []
        created: dict[str, dict] = {}
        window = _Window()

        class Production:
            def __init__(self, **kwargs):
                created["production"] = kwargs
                self.app = object()
                self.main_window = window

            def collect(self):
                events.append("production.collect")
                return True

            def process_collected_data(self):
                events.append("production.process")
                return True

        class Utility:
            def __init__(self, **kwargs):
                created["utility"] = kwargs

            def attach_existing_window(self, app, main_window):
                self.main_window = main_window

            def collect(self):
                events.append("utility.collect")
                return True

            def process_collected_data(self):
                events.append("utility.process")
                return True

        class WIP:
            def __init__(self, **kwargs):
                created["wip"] = kwargs

            def attach_existing_window(self, app, main_window):
                self.main_window = main_window

            def collect(self):
                events.append("wip.collect")
                return True

            def process_collected_data(self):
                events.append("wip.process")
                return True

        modules = {
            "production_daily_rpa": types.SimpleNamespace(MISProductionRPA=Production),
            "utility_daily_rpa": types.SimpleNamespace(MISUtilityRPA=Utility),
            "wip_daily_rpa": types.SimpleNamespace(MISWIPRPA=WIP),
        }
        args = Namespace(
            date=None,
            date_from="2026-01-01",
            date_to="2026-03-15",
            factories="광주",
            dry_run=False,
        )

        with (
            patch.dict(sys.modules, modules),
            patch.object(
                run_all_rpa.argparse.ArgumentParser,
                "parse_known_args",
                return_value=(args, []),
            ),
        ):
            rc = run_all_rpa.main()

        self.assertEqual(rc, 0)
        self.assertEqual(
            events,
            [
                "production.collect",
                "utility.collect",
                "wip.collect",
                "production.process",
                "wip.process",
                "utility.process",
            ],
        )
        for rpa_name in ("production", "utility", "wip"):
            self.assertEqual(created[rpa_name]["date_from"], "2026-01-01")
            self.assertEqual(created[rpa_name]["date_to"], "2026-03-15")
        self.assertEqual(created["production"]["factory_codes"], ["F30"])
        self.assertEqual(created["utility"]["org_codes"], ["F30"])
        self.assertEqual(created["wip"]["factory_codes"], ["F30"])


    def test_utility_processing_waits_for_successful_production_processing(self) -> None:
        events: list[str] = []
        window = _Window()

        class Production:
            def __init__(self, **kwargs):
                self.app = object()
                self.main_window = window

            def collect(self):
                events.append("production.collect")
                return True

            def process_collected_data(self):
                events.append("production.process")
                return False

        class Utility:
            def __init__(self, **kwargs):
                pass

            def attach_existing_window(self, app, main_window):
                pass

            def collect(self):
                events.append("utility.collect")
                return True

            def process_collected_data(self):
                events.append("utility.process")
                return True

        class WIP:
            def __init__(self, **kwargs):
                pass

            def attach_existing_window(self, app, main_window):
                pass

            def collect(self):
                events.append("wip.collect")
                return True

            def process_collected_data(self):
                events.append("wip.process")
                return True

        modules = {
            "production_daily_rpa": types.SimpleNamespace(MISProductionRPA=Production),
            "utility_daily_rpa": types.SimpleNamespace(MISUtilityRPA=Utility),
            "wip_daily_rpa": types.SimpleNamespace(MISWIPRPA=WIP),
        }
        args = Namespace(
            date=None,
            date_from=None,
            date_to=None,
            factories=None,
            dry_run=False,
        )

        with (
            patch.dict(sys.modules, modules),
            patch.object(
                run_all_rpa.argparse.ArgumentParser,
                "parse_known_args",
                return_value=(args, []),
            ),
        ):
            rc = run_all_rpa.main()

        self.assertEqual(rc, 1)
        self.assertNotIn("utility.process", events)
        self.assertIn("wip.process", events)

    def test_utility_processing_waits_for_successful_wip_processing(self) -> None:
        events: list[str] = []
        window = _Window()

        class Production:
            def __init__(self, **kwargs):
                self.app = object()
                self.main_window = window

            def collect(self):
                events.append("production.collect")
                return True

            def process_collected_data(self):
                events.append("production.process")
                return True

        class Utility:
            def __init__(self, **kwargs):
                pass

            def attach_existing_window(self, app, main_window):
                pass

            def collect(self):
                events.append("utility.collect")
                return True

            def process_collected_data(self):
                events.append("utility.process")
                return True

        class WIP:
            def __init__(self, **kwargs):
                pass

            def attach_existing_window(self, app, main_window):
                pass

            def collect(self):
                events.append("wip.collect")
                return True

            def process_collected_data(self):
                events.append("wip.process")
                return False

        modules = {
            "production_daily_rpa": types.SimpleNamespace(MISProductionRPA=Production),
            "utility_daily_rpa": types.SimpleNamespace(MISUtilityRPA=Utility),
            "wip_daily_rpa": types.SimpleNamespace(MISWIPRPA=WIP),
        }
        args = Namespace(
            date=None,
            date_from=None,
            date_to=None,
            factories=None,
            dry_run=False,
        )

        with (
            patch.dict(sys.modules, modules),
            patch.object(
                run_all_rpa.argparse.ArgumentParser,
                "parse_known_args",
                return_value=(args, []),
            ),
        ):
            rc = run_all_rpa.main()

        self.assertEqual(rc, 1)
        self.assertIn("wip.process", events)
        self.assertNotIn("utility.process", events)


if __name__ == "__main__":
    unittest.main()
