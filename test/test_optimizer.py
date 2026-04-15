import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.excel_io import parse_specification  # noqa: E402
from nordfox_raskroy.models import PartDemand, SpecRow  # noqa: E402
from nordfox_raskroy.optimizer import (  # noqa: E402
    optimize_cutting,
    sort_cuts_for_display,
    spec_rows_to_demands,
    summarize,
)
from nordfox_raskroy.profile_codes import filter_spec_by_profiles  # noqa: E402


class OptimizerTests(unittest.TestCase):
    def test_spec_rows_to_demands_expands_quantity(self):
        rows = [
            SpecRow(2, 1, "Модуль M1", "P", 1000, 90, 3, None),
        ]
        d = spec_rows_to_demands(rows)
        self.assertEqual(len(d), 3)
        self.assertTrue(all(x.length_mm == 1000 for x in d))

    def test_optimize_prefers_scrap_before_new_bar(self):
        demands = [
            PartDemand(1, "M", "a", 3000, 90),
            PartDemand(2, "M", "b", 2900, 90),
        ]
        r = optimize_cutting(
            demands,
            bar_lengths_mm=[6000],
            kerf_mm=0,
            min_scrap_mm=50,
        )
        self.assertEqual(sum(r.bars_used.values()), 1)
        sources = [c.stock_source for c in r.cuts]
        self.assertIn("new_bar", sources)
        self.assertIn("scrap", sources)


class ExcelPipelineTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.spec_path = ROOT / "test" / "spec_10x5_modules.xlsx"

    def test_parse_generated_spec_10x5(self):
        if not self.spec_path.is_file():
            self.skipTest("Запустите scripts/generate_spec_10x5.py")
        rows = parse_specification(self.spec_path)
        self.assertEqual(len(rows), 50)
        modules = {r.module_name for r in rows}
        self.assertEqual(len(modules), 10)
        self.assertEqual(sum(r.quantity for r in rows), 50)

    def test_full_pipeline_generated_xlsx(self):
        if not self.spec_path.is_file():
            self.skipTest("Запустите scripts/generate_spec_10x5.py")
        rows = parse_specification(self.spec_path)
        rows, _warns = filter_spec_by_profiles(rows, {0, 1, 2, 3})
        demands = spec_rows_to_demands(rows)
        result = optimize_cutting(
            demands,
            bar_lengths_mm=[6000, 7500, 12000],
            kerf_mm=0,
            min_scrap_mm=50,
        )
        self.assertTrue(result.cuts)
        s = summarize(result)
        self.assertIn("Всего заготовок", s)
        ordered = sort_cuts_for_display(result.cuts, by_module=True)
        self.assertEqual(len(ordered), len(result.cuts))
        self.assertEqual(ordered[0].demand.module_name, "Модуль M1")
        self.assertEqual(ordered[-1].demand.module_name, "Модуль M10")


if __name__ == "__main__":
    unittest.main()
