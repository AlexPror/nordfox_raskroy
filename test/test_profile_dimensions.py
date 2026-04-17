"""Тесты справочника габаритов сечения профилей."""

from __future__ import annotations

import pytest

from nordfox_raskroy.profile_dimensions import (
    extra_trailing_end_clearance_mm,
    part_trailing_angle_deg,
    profile_section_max_side_mm,
    profile_section_mm,
)


@pytest.mark.parametrize(
    "code, h, w",
    [
        ("СК-0-3000", 60.0, 60.0),
        ("сс-1-1000", 82.9, 60.0),
        ("Р-2-500", 90.0, 60.0),
        ("Р-3-500", 100.0, 60.0),
    ],
)
def test_series_digit_maps_to_n20_n23(code: str, h: float, w: float) -> None:
    s = profile_section_mm(code)
    assert s is not None
    assert s.height_mm == h
    assert s.width_mm == w


def test_l16_preferred_over_l1_substring() -> None:
    """Длинный ключ раньше: «l16» не должен матчиться как «l1»."""
    s = profile_section_mm("Профиль L16-1000")
    assert s is not None
    assert s.height_mm == 62.1
    assert s.width_mm == 40.0


def test_free_text_h22() -> None:
    s = profile_section_mm("Несущий Н 22 модуль")
    assert s is not None
    assert s.height_mm == 90.0


def test_dt20() -> None:
    s = profile_section_mm("DT20 тёплый")
    assert s is not None
    assert s.height_mm == 55.3
    assert s.width_mm == 60.0


def test_max_side() -> None:
    assert profile_section_max_side_mm("СК-1-1000") == pytest.approx(82.9)
    assert profile_section_max_side_mm("L20") == pytest.approx(60.0)


def test_unknown_returns_none() -> None:
    assert profile_section_mm("") is None
    assert profile_section_mm("швеллер 20п") is None
    assert profile_section_mm("СК-9-1000") is None


def test_part_trailing_angle() -> None:
    assert part_trailing_angle_deg(90, None) == 90
    assert part_trailing_angle_deg(90, 45) == 45


def test_extra_trailing_clearance_square() -> None:
    assert extra_trailing_end_clearance_mm("СК-0-1", 90) == 0


def test_extra_trailing_clearance_miter_n20() -> None:
    # h_max=60, |90-45|=45° -> 60*tan(45°)=60
    assert extra_trailing_end_clearance_mm("СК-0-1", 45) == 60


def test_extra_trailing_unknown_profile() -> None:
    assert extra_trailing_end_clearance_mm("unknown-xyz", 45) == 0
