from __future__ import annotations

import re


def module_order_key(name: str) -> tuple[int, str]:
    m = re.search(r"[MМ]\s*(\d+)", name, flags=re.IGNORECASE)
    if m:
        return (int(m.group(1)), name)
    return (10**9, name)


def module_short_name(name: str) -> str:
    m = re.search(r"[MМ]\s*(\d+)", name, flags=re.IGNORECASE)
    if m:
        return f"М{int(m.group(1))}"
    return name
