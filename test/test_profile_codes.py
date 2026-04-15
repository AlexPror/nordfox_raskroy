import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from nordfox_raskroy.models import SpecRow  # noqa: E402
from nordfox_raskroy.profile_codes import (  # noqa: E402
    filter_spec_by_profiles,
    parse_profile_series_digit,
    profile_label_for_code,
)


class ProfileCodesTests(unittest.TestCase):
    def test_parse_sk_ss_r(self):
        self.assertEqual(parse_profile_series_digit("СК-1-1045Л"), 1)
        self.assertEqual(parse_profile_series_digit("сс-0-456"), 0)
        self.assertEqual(parse_profile_series_digit("Р-3-3030В"), 3)
        self.assertEqual(parse_profile_series_digit("СС-2-1200"), 2)

    def test_invalid_digit_or_format(self):
        self.assertIsNone(parse_profile_series_digit("СК-4-1000"))
        self.assertIsNone(parse_profile_series_digit("XX-0-1000"))
        self.assertIsNone(parse_profile_series_digit(""))

    def test_label(self):
        self.assertEqual(profile_label_for_code("СК-0-1"), "Н20")
        self.assertEqual(profile_label_for_code("bad"), "—")

    def test_filter(self):
        rows = [
            SpecRow(2, 1, "M", "СК-0-100", 100, 90, 1, None),
            SpecRow(3, 1, "M", "СК-1-200", 200, 90, 1, None),
        ]
        kept, w = filter_spec_by_profiles(rows, {0})
        self.assertEqual(len(kept), 1)
        self.assertEqual(kept[0].profile_code, "СК-0-100")
        self.assertTrue(any("не включён" in x for x in w))


if __name__ == "__main__":
    unittest.main()
