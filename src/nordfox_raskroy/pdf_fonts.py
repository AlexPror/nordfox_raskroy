from __future__ import annotations

import logging
from pathlib import Path


def reportlab_cyrillic_fonts(logger: logging.Logger | None = None) -> tuple[str, str]:
    """Return ReportLab font names with Cyrillic support if available."""
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except ImportError:
        return "Helvetica", "Helvetica-Bold"
    if "NF-Regular" in pdfmetrics.getRegisteredFontNames():
        return "NF-Regular", "NF-Bold"
    candidates = [
        (Path(r"C:\Windows\Fonts\arial.ttf"), Path(r"C:\Windows\Fonts\arialbd.ttf")),
        (Path(r"C:\Windows\Fonts\DejaVuSans.ttf"), Path(r"C:\Windows\Fonts\DejaVuSans-Bold.ttf")),
        (
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
            Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ),
    ]
    for reg_path, bold_path in candidates:
        if reg_path.is_file() and bold_path.is_file():
            try:
                pdfmetrics.registerFont(TTFont("NF-Regular", str(reg_path)))
                pdfmetrics.registerFont(TTFont("NF-Bold", str(bold_path)))
                return "NF-Regular", "NF-Bold"
            except Exception:  # noqa: BLE001
                if logger is not None:
                    logger.exception("reportlab TTF register failed: %s", reg_path)
    return "Helvetica", "Helvetica-Bold"
