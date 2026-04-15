import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.export_results import export_cuts_excel  # noqa: E402
from nordfox_raskroy.models import CutEvent, PartDemand  # noqa: E402
from nordfox_raskroy.module_colors import (  # noqa: E402
    module_base_rgb,
    module_palette_index,
    module_row_rgb,
    rgb_to_openpyxl_argb,
)


class ModuleColorsTests(unittest.TestCase):
    def test_stable_module_index(self):
        self.assertEqual(module_palette_index("Модуль M1"), 0)
        self.assertEqual(module_palette_index("Модуль M2"), 1)

    def test_scrap_darkens(self):
        a = module_row_rgb("Модуль M1", is_scrap=False)
        b = module_row_rgb("Модуль M1", is_scrap=True)
        self.assertLess(sum(b), sum(a))

    def test_argb_format(self):
        s = rgb_to_openpyxl_argb(module_base_rgb("Модуль M1"))
        self.assertEqual(len(s), 8)
        self.assertTrue(s.startswith("FF"))


class ExportExcelTests(unittest.TestCase):
    def test_export_excel_writes(self):
        cuts = [
            CutEvent(
                PartDemand(1, "Модуль M1", "СК-0-100", 100, 90),
                6000,
                "new_bar",
                5900,
                0,
            ),
        ]
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "out.xlsx"
            export_cuts_excel(cuts, p, summary="test")
            self.assertTrue(p.is_file())
            self.assertGreater(p.stat().st_size, 200)


if __name__ == "__main__":
    unittest.main()
