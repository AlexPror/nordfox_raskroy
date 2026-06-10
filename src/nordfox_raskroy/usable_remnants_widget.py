"""Виджет колонны деловых остатков (схема в цветах профиля)."""

from __future__ import annotations

try:
    from PySide6.QtCore import Qt, QRectF
    from PySide6.QtGui import QColor, QFont, QPainter, QPen
    from PySide6.QtWidgets import QWidget
except ImportError as e:  # pragma: no cover
    raise ImportError(
        "Нужен пакет PySide6. Установите: pip install PySide6"
    ) from e


class UsableRemnantsWidget(QWidget):
    """Вертикальная колонна остатков: цветная полоса слева, подпись профиля и длины справа."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._rows: list[dict[str, object]] = []
        self._color_map: dict[str, QColor] = {}
        self._zoom = 1.0
        self.setMinimumHeight(200)
        self.setMinimumWidth(520)

    def set_rows(self, rows: list[dict[str, object]], color_map: dict[str, QColor]) -> None:
        self._rows = rows
        self._color_map = color_map
        row_h = int(44 * self._zoom)
        self.setMinimumHeight(max(200, 56 + len(rows) * row_h))
        self.setMinimumWidth(max(520, int(520 * self._zoom)))
        self.updateGeometry()
        self.update()

    def clear_rows(self) -> None:
        self._rows = []
        self.setMinimumHeight(200)
        self.update()

    def set_zoom(self, zoom: float) -> None:
        self._zoom = max(1.0, min(2.0, float(zoom)))
        row_h = int(44 * self._zoom)
        self.setMinimumHeight(max(200, 56 + len(self._rows) * row_h))
        self.setMinimumWidth(max(520, int(520 * self._zoom)))
        self.update()

    def paintEvent(self, _event) -> None:  # noqa: N802
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.fillRect(self.rect(), QColor(248, 250, 252))
        z = self._zoom

        if not self._rows:
            p.setPen(QColor(51, 65, 85))
            p.setFont(QFont("Segoe UI", 9))
            p.drawText(16, 28, "Нет деловых остатков ≥100 мм")
            return

        if self._color_map:
            p.setPen(QColor(30, 41, 59))
            p.setFont(QFont("Segoe UI", 9, QFont.Bold))
            p.drawText(14, 20, "Легенда профилей:")
            lx = 140.0
            p.setFont(QFont("Segoe UI", 8))
            for name in sorted(self._color_map):
                c = self._color_map[name]
                p.setPen(QColor(148, 163, 184))
                p.setBrush(c)
                p.drawRect(QRectF(lx, 10, 10, 10))
                p.setPen(QColor(30, 41, 59))
                p.drawText(int(lx + 14), 19, name)
                lx += max(72.0, 14.0 + len(name) * 6.0)

        bar_area_w = max(200.0, float(self.width()) * 0.52)
        label_x = 16.0 + bar_area_w + 20.0
        label_w = max(160.0, float(self.width()) - label_x - 12.0)
        row_h = 38.0 * z
        gap = 8.0 * z
        y = 36.0 * z
        bar_h = 22.0 * z
        max_len = max(int(r.get("max_length_mm", 1) or 1) for r in self._rows)

        for r in self._rows:
            length_mm = int(r.get("length_mm", 0))
            profile_code = str(r.get("profile_code", ""))
            profile_name = str(r.get("profile_name", ""))
            remnant_no = int(r.get("remnant_no", 0))
            color = self._color_map.get(profile_name, QColor(147, 197, 253))

            block = QRectF(8, y, self.width() - 16, row_h)
            p.setPen(QPen(QColor(203, 213, 225), 1.0))
            p.setBrush(QColor(255, 255, 255))
            p.drawRoundedRect(block, 5, 5)

            frac = min(1.0, float(length_mm) / float(max_len)) if max_len > 0 else 0.0
            bar_w = max(12.0, (bar_area_w - 24.0) * frac)
            bar_x = 20.0
            bar_y = y + (row_h - bar_h) / 2.0
            bar_rect = QRectF(bar_x, bar_y, bar_w, bar_h)

            p.setPen(QPen(QColor(148, 163, 184), 1.0))
            p.setBrush(QColor(241, 245, 249))
            p.drawRoundedRect(QRectF(bar_x, bar_y, bar_area_w - 24.0, bar_h), 3, 3)

            p.setPen(QPen(color.darker(115), 1.0))
            p.setBrush(color)
            p.drawRoundedRect(bar_rect, 3, 3)

            p.setPen(QColor(15, 23, 42))
            p.setFont(QFont("Segoe UI", max(7, int(round(8 * z))), QFont.Bold))
            if bar_w >= 36:
                p.drawText(
                    bar_rect.adjusted(4, 0, -4, 0),
                    int(Qt.AlignCenter),
                    f"{length_mm} мм",
                )

            p.setPen(QColor(30, 41, 59))
            p.setFont(QFont("Segoe UI", max(8, int(round(9 * z))), QFont.Bold))
            code_txt = profile_code if profile_code and profile_code != "—" else profile_name
            p.drawText(
                QRectF(label_x, y + 4, label_w, row_h / 2.0),
                int(Qt.AlignLeft | Qt.AlignVCenter),
                code_txt,
            )
            p.setFont(QFont("Segoe UI", max(7, int(round(8 * z)))))
            p.setPen(QColor(71, 85, 105))
            sub = f"№{remnant_no}  ·  {length_mm} мм"
            opening = int(r.get("opening", 0) or 0)
            if opening > 0:
                sub += f"  ·  пруток {opening}"
            p.drawText(
                QRectF(label_x, y + row_h / 2.0 - 2, label_w, row_h / 2.0),
                int(Qt.AlignLeft | Qt.AlignVCenter),
                sub,
            )
            y += row_h + gap
