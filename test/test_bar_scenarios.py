import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.bar_scenarios import (  # noqa: E402
    ScenarioOutcome,
    compare_bar_scenarios,
    pick_recommended,
)
from nordfox_raskroy.models import PartDemand  # noqa: E402


class BarScenarioTests(unittest.TestCase):
    def test_recommend_only_shorter_bars_when_possible(self):
        demands = [
            PartDemand(1, "M1", "СК-0-500", 800, 90),
            PartDemand(2, "M1", "СК-0-500", 800, 90),
        ]
        out = compare_bar_scenarios(demands, kerf_mm=0, min_scrap_mm=50)
        rec = pick_recommended(out)
        assert rec is not None
        self.assertIn(6000, rec.bars_mm)
        self.assertLessEqual(max(rec.bars_mm), 6000)

    def test_initial_scrap_reduces_new_bars(self):
        demands = [PartDemand(1, "M1", "СК-0-100", 500, 90)]
        without = compare_bar_scenarios(
            demands, kerf_mm=0, min_scrap_mm=50, initial_scraps_mm=None
        )
        with_scrap = compare_bar_scenarios(
            demands,
            kerf_mm=0,
            min_scrap_mm=50,
            initial_scraps_mm=[6000],
        )
        w0 = pick_recommended(without)
        w1 = pick_recommended(with_scrap)
        assert w0 and w1
        self.assertLessEqual(w1.total_bars, w0.total_bars)

    def test_pick_recommended_mode_waste_first(self):
        outs = [
            ScenarioOutcome("A", (6000,), True, "", 10, 60000, 57000, 3000, 5.0, object()),
            ScenarioOutcome("B", (12000,), True, "", 8, 96000, 83000, 13000, 13.5, object()),
        ]
        rec = pick_recommended(outs, mode="waste_first")
        assert rec is not None
        self.assertEqual(rec.name, "A")

    def test_pick_recommended_mode_bars_first(self):
        outs = [
            ScenarioOutcome("A", (6000,), True, "", 10, 60000, 57000, 3000, 5.0, object()),
            ScenarioOutcome("B", (12000,), True, "", 8, 96000, 83000, 13000, 13.5, object()),
        ]
        rec = pick_recommended(outs, mode="bars_first")
        assert rec is not None
        self.assertEqual(rec.name, "B")


if __name__ == "__main__":
    unittest.main()
