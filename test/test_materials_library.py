import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.materials_library import (  # noqa: E402
    kg_per_meter_from_description,
    kg_per_meter_from_profile_code,
    kg_per_meter_nordfox_series,
    row_mass_kg_display,
    total_mass_kg,
)


class MaterialsLibraryTests(unittest.TestCase):
    def test_series_digits_match_nordfox(self):
        self.assertEqual(kg_per_meter_nordfox_series(0), 0.937)
        self.assertEqual(kg_per_meter_nordfox_series(3), 1.48)

    def test_codes_without_suffix_letters(self):
        for code, exp in [
            ("СК-1-456", 1.132),
            ("СС-2-554", 1.358),
            ("Р-0-1230", 0.937),
        ]:
            with self.subTest(code=code):
                m = kg_per_meter_from_profile_code(code)
                self.assertEqual(m.kg_per_m, exp)
                self.assertTrue(m.source.startswith("nordfox_series"))

    def test_rolled_channel(self):
        m = kg_per_meter_from_description("Швеллер 16П ГОСТ 8240, длина 6000")
        self.assertAlmostEqual(m.kg_per_m, 15.4, places=2)
        self.assertIn("rolled", m.source)

    def test_total_mass(self):
        self.assertAlmostEqual(total_mass_kg(6000, 1.206, 2), 1.206 * 6 * 2, places=4)

    def test_row_mass_display(self):
        kg, txt = row_mass_kg_display("СК-1-100", 1000.0, 1.0)
        self.assertIsNotNone(kg)
        self.assertGreater(kg, 0)
        self.assertNotEqual(txt, "—")


if __name__ == "__main__":
    unittest.main()
