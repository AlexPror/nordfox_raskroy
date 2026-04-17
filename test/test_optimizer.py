import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.excel_io import parse_specification  # noqa: E402
from nordfox_raskroy.models import PartDemand, SpecRow  # noqa: E402
from nordfox_raskroy.optimizer import (  # noqa: E402
    angle_offset_mm,
    demand_cut_length_mm,
    optimize_cutting,
    sort_cuts_for_display,
    spec_rows_to_demands,
    summarize,
)
from nordfox_raskroy.profile_dimensions import extra_trailing_end_clearance_mm  # noqa: E402
from nordfox_raskroy.profile_codes import filter_spec_by_profiles  # noqa: E402


class OptimizerTests(unittest.TestCase):
    def test_angle_offset_rule(self):
        self.assertEqual(angle_offset_mm(90), 30)
        self.assertEqual(angle_offset_mm(45), 50)

    def test_demand_cut_length_includes_offset_and_kerf(self):
        d90 = PartDemand(1, "M", "a", 1000, 90)
        d45 = PartDemand(2, "M", "b", 1000, 45)
        # length + техотступ + 2×kerf (см. optimizer.demand_cut_length_mm)
        self.assertEqual(demand_cut_length_mm(d90, kerf_mm=2), 1000 + 30 + 4)
        self.assertEqual(demand_cut_length_mm(d45, kerf_mm=2), 1000 + 50 + 4)

    def test_two_angles_only_first_counts_for_length(self):
        d = PartDemand(1, "M", "a", 1000, 90, cut_angle_2=45)
        self.assertEqual(demand_cut_length_mm(d, kerf_mm=0), 1000 + 30)

    def test_optimize_applies_offset_for_each_cut(self):
        # Один профиль — хвост первого реза можно резать второй деталью.
        demands = [
            PartDemand(1, "M", "СК-0-1000", 1000, 90),
            PartDemand(2, "M", "СК-0-2000", 1000, 45),
        ]
        r = optimize_cutting(
            demands,
            bar_lengths_mm=[6000],
            kerf_mm=0,
            min_scrap_mm=50,
        )
        self.assertEqual(sum(r.bars_used.values()), 1)
        consumed = sum(c.stock_length_mm - c.remainder_mm for c in r.cuts)
        self.assertEqual(consumed, 2080)

    def test_spec_rows_to_demands_expands_quantity(self):
        rows = [
            SpecRow(2, 1, "Модуль M1", "P", 1000, 90, 3, None),
        ]
        d = spec_rows_to_demands(rows)
        self.assertEqual(len(d), 3)
        self.assertTrue(all(x.length_mm == 1000 for x in d))

    def test_optimize_prefers_scrap_before_new_bar(self):
        demands = [
            PartDemand(1, "M", "СК-0-1", 3000, 90),
            PartDemand(2, "M", "СК-0-2", 2900, 90),
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

    def test_stock_opening_order_new_bar_then_scrap_same_lineage(self):
        """Сначала строки по 1-му прутку (новый + хвост), затем по следующему."""
        demands = [
            PartDemand(1, "M", "СК-0-1", 3000, 90),
            PartDemand(2, "M", "СК-0-2", 2900, 90),
            PartDemand(3, "M", "СК-0-3", 500, 90),
        ]
        r = optimize_cutting(
            demands,
            bar_lengths_mm=[6000],
            kerf_mm=0,
            min_scrap_mm=50,
        )
        ids = [c.stock_opening_id for c in r.cuts]
        self.assertEqual(ids, sorted(ids))
        self.assertEqual(ids[0], 1)
        self.assertEqual(ids[1], 1)
        new_ids = {c.stock_opening_id for c in r.cuts if c.stock_source == "new_bar"}
        self.assertEqual(new_ids, {1, 2})

    def test_scrap_rejected_if_short_for_trailing_miter_geometry(self):
        """Обрезок достаточен по 1D cut_len, но короче cut_len + гео. запаса — берём новый пруток."""
        # Н20: h_max=60, trailing 45° -> +60 мм к длине куска при выборе обрезка.
        self.assertEqual(extra_trailing_end_clearance_mm("СК-0-1", 45), 60)
        d = PartDemand(1, "M", "СК-0-1", 1000, 45)
        self.assertEqual(demand_cut_length_mm(d, kerf_mm=0), 1050)
        r = optimize_cutting(
            [d],
            bar_lengths_mm=[6000],
            kerf_mm=0,
            min_scrap_mm=50,
            initial_scraps_mm=[1100],
        )
        self.assertEqual(r.cuts[0].stock_source, "new_bar")
        self.assertEqual(r.cuts[0].stock_length_mm, 6000)

    def test_scrap_accepted_when_long_enough_for_miter(self):
        d = PartDemand(1, "M", "СК-0-1", 1000, 45)
        r = optimize_cutting(
            [d],
            bar_lengths_mm=[6000],
            kerf_mm=0,
            min_scrap_mm=50,
            initial_scraps_mm=[1110],
        )
        self.assertEqual(r.cuts[0].stock_source, "scrap")
        self.assertEqual(r.cuts[0].stock_length_mm, 1110)


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
