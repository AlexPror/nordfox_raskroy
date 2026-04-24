import sys
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.album_plan_service import build_album_plan_rows  # noqa: E402
from nordfox_raskroy.bar_advisor_service import BarAdvisorResult, run_bar_advisor  # noqa: E402
from nordfox_raskroy.layout_plan_service import build_layout_plan_rows  # noqa: E402
from nordfox_raskroy.models import CutEvent, PartDemand, SpecRow  # noqa: E402
from nordfox_raskroy.spec_profile_filters import (  # noqa: E402
    filter_rows_by_selected_profiles,
    profile_filter_key,
)


def _sample_demands() -> list[PartDemand]:
    return [
        PartDemand(
            spec_row_index=1,
            module_name="Модуль М1",
            profile_code="СК-0-100",
            length_mm=1000,
            cut_angle=90,
            cut_angle_2=45,
        )
    ]


def _sample_cuts() -> list[CutEvent]:
    d1 = PartDemand(
        spec_row_index=1,
        module_name="Модуль М2",
        profile_code="СК-0-100",
        length_mm=1000,
        cut_angle=90,
        cut_angle_2=45,
    )
    d2 = PartDemand(
        spec_row_index=2,
        module_name="Модуль М2",
        profile_code="L15",
        length_mm=800,
        cut_angle=90,
        cut_angle_2=90,
    )
    return [
        CutEvent(demand=d1, stock_length_mm=6000, stock_source="new_bar", remainder_mm=0, waste_mm=0, stock_opening_id=1),
        CutEvent(demand=d2, stock_length_mm=6000, stock_source="new_bar", remainder_mm=0, waste_mm=0, stock_opening_id=1),
    ]


class RefactorServicesTests(unittest.TestCase):
    def test_profile_filter_key_series_and_custom(self):
        self.assertEqual(profile_filter_key("СК-0-100"), "Н20")
        self.assertEqual(profile_filter_key("L15"), "L15")
        self.assertEqual(profile_filter_key(""), "")

    def test_filter_rows_by_selected_profiles(self):
        rows = [
            SpecRow(1, None, "М1", "СК-0-100", 1000, 90, 1),
            SpecRow(2, None, "М1", "L15", 900, 90, 1),
        ]
        kept, warns = filter_rows_by_selected_profiles(rows, {"Н20"})
        self.assertEqual(len(kept), 1)
        self.assertEqual(kept[0].row_index, 1)
        self.assertEqual(len(warns), 1)

    @patch("nordfox_raskroy.bar_advisor_service.pick_recommended")
    @patch("nordfox_raskroy.bar_advisor_service.format_scenario_report")
    @patch("nordfox_raskroy.bar_advisor_service.compare_bar_scenarios")
    def test_run_bar_advisor_wiring(
        self,
        mock_compare,
        mock_report,
        mock_pick,
    ):
        mock_compare.return_value = ["outcome"]
        mock_report.return_value = "report"
        rec = type("Rec", (), {"bars_mm": (6850,), "name": "Только 6850 мм"})()
        mock_pick.return_value = rec

        result = run_bar_advisor(
            _sample_demands(),
            kerf_mm=5,
            offset_90_mm=30,
            offset_other_mm=50,
            base_len_mm=7000,
            standard_mode=True,
            initial_scraps_mm=[],
            min_scrap_mm=0,
            mode="waste_first",
        )

        self.assertIsInstance(result, BarAdvisorResult)
        self.assertEqual(result.report_text, "report")
        self.assertEqual(result.recommended_length_mm, 6850)
        self.assertEqual(result.recommended_name, "Только 6850 мм")
        self.assertEqual(result.recommended_bars_mm, (6850,))
        self.assertTrue(mock_compare.called)

    def test_run_bar_advisor_validates_base_length(self):
        with self.assertRaises(ValueError):
            run_bar_advisor(
                _sample_demands(),
                kerf_mm=5,
                offset_90_mm=30,
                offset_other_mm=50,
                base_len_mm=0,
                standard_mode=True,
                initial_scraps_mm=None,
            )

    def test_layout_plan_rows(self):
        result = build_layout_plan_rows(
            _sample_cuts(),
            kerf_mm=5,
            offset_90_mm=30,
            offset_other_mm=50,
            opening_color_rgb=lambda _o: (10, 20, 30),
        )
        self.assertEqual(len(result.rows), 1)
        self.assertIn("Н20", result.profile_names)
        self.assertIn("L15", result.profile_names)
        segments = result.rows[0]["segments"]
        self.assertTrue(any(s.get("kind") == "profile" for s in segments))

    def test_album_plan_rows_details_and_joints(self):
        cuts = _sample_cuts()
        details = build_album_plan_rows(
            cuts,
            mode="details",
            offset_90_mm=30,
            offset_other_mm=50,
            kerf_mm=5,
        )
        joints = build_album_plan_rows(
            cuts,
            mode="joints",
            offset_90_mm=30,
            offset_other_mm=50,
            kerf_mm=5,
        )
        self.assertEqual(len(details.rows), 2)
        self.assertEqual(len(joints.rows), 1)
        self.assertIn("Н20", details.profile_names)
        self.assertIn("L15", details.profile_names)


if __name__ == "__main__":
    unittest.main()
