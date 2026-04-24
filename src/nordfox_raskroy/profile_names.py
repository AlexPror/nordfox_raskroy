from __future__ import annotations

from nordfox_raskroy.profile_codes import profile_label_for_code


def display_profile_name(profile_code: str) -> str:
    label = profile_label_for_code(profile_code)
    if label == "—":
        return profile_code
    return label
