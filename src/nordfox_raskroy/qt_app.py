"""Десктопное приложение на Qt (PySide6)."""

from __future__ import annotations

from collections import defaultdict
import logging
from pathlib import Path
import re

try:
    from PySide6.QtCore import QObject, QEvent, QPointF, QRectF, Qt, Signal
    from PySide6.QtGui import (
        QAction,
        QBrush,
        QColor,
        QFont,
        QFontMetrics,
        QKeySequence,
        QPainter,
        QPen,
        QPolygonF,
        QRegion,
        QShortcut,
    )
    from PySide6.QtWidgets import (
        QAbstractItemView,
        QApplication,
        QCheckBox,
        QComboBox,
        QFileDialog,
        QGridLayout,
        QGroupBox,
        QHBoxLayout,
        QHeaderView,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMenu,
        QMessageBox,
        QInputDialog,
        QPushButton,
        QPlainTextEdit,
        QScrollArea,
        QSizePolicy,
        QSplitter,
        QTableWidget,
        QTableWidgetItem,
        QTabWidget,
        QTextEdit,
        QVBoxLayout,
        QWidget,
    )
except ImportError as e:  # pragma: no cover
    raise ImportError(
        "Нужен пакет PySide6. Установите: pip install PySide6"
    ) from e

from nordfox_raskroy.bar_scenarios import (
    compare_bar_scenarios,
    format_scenario_report,
    pick_recommended,
)
from nordfox_raskroy.excel_io import parse_specification, parse_specification_with_stats
from nordfox_raskroy.export_results import export_cuts_excel, export_cuts_pdf
from nordfox_raskroy.models import CutEvent, OptimizationResult, PartDemand
from nordfox_raskroy.module_colors import module_row_rgb
from nordfox_raskroy.optimizer import (
    demand_cut_length_mm,
    format_cut_angles,
    optimize_cutting,
    spec_rows_to_demands,
    summarize,
)
from nordfox_raskroy.result_sort import SORT_MODES, sort_cuts
from nordfox_raskroy.materials_library import (
    kg_per_meter_from_profile_code,
    row_mass_kg_display,
    total_mass_kg,
)
from nordfox_raskroy.profile_codes import (
    PROFILE_DIGIT_TO_NAME,
    filter_spec_by_profiles,
    profile_label_for_code,
)
from nordfox_raskroy.scrap_stock_io import parse_scrap_inventory
from nordfox_raskroy import __version__
from nordfox_raskroy.table_demand_import import demands_from_cut_table_rows
logger = logging.getLogger("nordfox_raskroy.qt_app")

# Техзона и пропил на схеме — отдельные цвета; углы только подписью, без линий на профиле.
_TECH_VIS_FILL = QColor(255, 237, 213)
_TECH_VIS_LINE = QColor(234, 88, 12)
_KERF_VIS_FILL = QColor(207, 250, 254)
_KERF_VIS_LINE = QColor(14, 116, 144)


def _reportlab_cyrillic_fonts() -> tuple[str, str]:
    """Имена шрифтов ReportLab с поддержкой кириллицы; иначе Helvetica (без кириллицы)."""
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
                logger.exception("reportlab TTF register failed: %s", reg_path)
    return "Helvetica", "Helvetica-Bold"


def _paint_diagonal_hatch(
    painter: QPainter,
    polygon: QPolygonF,
    *,
    fill: QColor,
    line_color: QColor,
    spacing: float,
    line_width: float,
    cross: bool = False,
) -> None:
    """Заливка + явные диагонали в клипе полигона (читаемо, как в CAD)."""
    painter.setPen(Qt.NoPen)
    painter.setBrush(fill)
    painter.drawPolygon(polygon)
    br = polygon.boundingRect()
    if br.width() < 0.5 or br.height() < 0.5:
        return
    painter.save()
    painter.setClipRegion(QRegion(polygon.toPolygon()))
    painter.setBrush(Qt.NoBrush)
    painter.setPen(QPen(line_color, line_width))
    h = max(br.height(), 1.0)
    x = br.left() - h
    while x < br.right() + spacing:
        painter.drawLine(QPointF(x, br.bottom()), QPointF(x + h, br.top()))
        x += spacing
    if cross:
        painter.setPen(QPen(line_color, max(0.6, line_width * 0.75)))
        x2 = br.left() - h + spacing * 0.5
        while x2 < br.right() + spacing:
            painter.drawLine(QPointF(x2, br.top()), QPointF(x2 + h, br.bottom()))
            x2 += spacing * 2.0
    painter.restore()


def _fmt_kg_trim(v: float) -> str:
    """Формат кг как на диаграммах: до 3 знаков, без лишних нулей."""
    return f"{v:.3f}".rstrip("0").rstrip(".")


_OPENING_COLORS = (
    QColor(224, 242, 254),
    QColor(220, 252, 231),
    QColor(254, 243, 199),
    QColor(243, 232, 255),
    QColor(252, 231, 243),
    QColor(226, 232, 240),
)


def opening_row_color(opening: int) -> QColor:
    return _OPENING_COLORS[(max(1, int(opening)) - 1) % len(_OPENING_COLORS)]


def run_app() -> None:
    logger.info("Qt app bootstrap")
    app = QApplication([])
    app.setStyle("Fusion")

    win = MainWindow()
    win.show()
    app.exec()


class LogEmitter(QObject):
    message = Signal(str)


class QtLogHandler(logging.Handler):
    """Прокидывает записи logging в Qt-панель логов."""

    def __init__(self, emitter: LogEmitter):
        super().__init__()
        self.emitter = emitter

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.emitter.message.emit(self.format(record))
        except Exception:
            self.handleError(record)


class RingBreakdownWidget(QWidget):
    """Универсальная кольцевая диаграмма с легендой и центром."""

    _SEGMENT_COLORS = (
        QColor(59, 130, 246),
        QColor(16, 185, 129),
        QColor(245, 158, 11),
        QColor(168, 85, 247),
        QColor(236, 72, 153),
    )
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setMinimumSize(220, 170)
        self._segments: list[tuple[str, float, QColor]] = []
        self._legend_lines: list[tuple[str, QColor]] = []
        self._title = "Диаграмма"
        self._subtitle = "Нет данных"
        self._center_top = "—"
        self._center_bottom = ""

    def clear_data(self) -> None:
        self._segments = []
        self._legend_lines = []
        self._subtitle = "Нет данных"
        self._center_top = "—"
        self._center_bottom = ""
        self.update()

    def set_data(
        self,
        *,
        title: str,
        subtitle: str,
        center_top: str,
        center_bottom: str,
        values_kg: dict[str, float],
        color_map: dict[str, QColor] | None = None,
    ) -> None:
        self._title = title
        self._subtitle = subtitle
        self._center_top = center_top
        self._center_bottom = center_bottom
        self._segments = []
        self._legend_lines = []
        pairs = sorted(
            ((k, v) for k, v in values_kg.items() if v > 0),
            key=lambda x: x[0],
        )
        for i, (name, kg) in enumerate(pairs):
            color = (
                color_map[name]
                if color_map is not None and name in color_map
                else self._SEGMENT_COLORS[i % len(self._SEGMENT_COLORS)]
            )
            self._segments.append((name, kg, color))
            self._legend_lines.append((f"{name}: {kg:.3f} кг", color))
        self.update()

    def paintEvent(self, _event) -> None:  # noqa: N802
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.fillRect(self.rect(), QColor(248, 250, 252))

        title_font = QFont("Segoe UI", 9, QFont.Bold)
        p.setPen(QColor(31, 41, 55))
        p.setFont(title_font)
        p.drawText(12, 20, self._title)
        p.setFont(QFont("Segoe UI", 8))
        p.setPen(QColor(75, 85, 99))
        p.drawText(12, 36, self._subtitle)

        pie_rect = QRectF(12, 46, 108, 108)
        inner_rect = QRectF(34, 68, 64, 64)
        total = sum(v for _n, v, _c in self._segments)
        start = 90.0
        for _name, val, color in self._segments:
            if total <= 0 or val <= 0:
                continue
            span = 360.0 * (val / total)
            p.setPen(QPen(QColor(255, 255, 255), 1.2))
            p.setBrush(color)
            p.drawPie(pie_rect, int(start * 16), int(-span * 16))
            start -= span

        p.setPen(Qt.NoPen)
        p.setBrush(QColor(248, 250, 252))
        p.drawEllipse(inner_rect)

        p.setPen(QColor(17, 24, 39))
        p.setFont(QFont("Segoe UI", 8, QFont.Bold))
        p.drawText(inner_rect, int(Qt.AlignCenter), f"{self._center_top}\n{self._center_bottom}")

        p.setFont(QFont("Segoe UI", 8))
        y = 54
        for text, color in self._legend_lines:
            p.setPen(Qt.NoPen)
            p.setBrush(color)
            p.drawRect(132, y - 8, 10, 10)
            p.setPen(QColor(31, 41, 55))
            p.drawText(146, y, text)
            y += 18


class CuttingLayoutWidget(QWidget):
    """Визуальная схема раскроя по пруткам (прямоугольники с сегментами)."""

    overflowLegendChanged = Signal(list)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._rows: list[dict[str, object]] = []
        self._color_map: dict[str, QColor] = {}
        self._zoom = 1.0
        self._overflow_entries: list[dict[str, object]] = []
        self._overflow_index_by_key: dict[tuple[int, int], int] = {}
        self.setMinimumHeight(240)

    def set_plan(self, rows: list[dict[str, object]], color_map: dict[str, QColor]) -> None:
        self._rows = rows
        self._color_map = color_map
        row_h = int(52 * self._zoom)
        self.setMinimumHeight(max(240, 48 + len(rows) * row_h))
        self._rebuild_overflow_index()
        self.update()

    def clear_plan(self) -> None:
        self._rows = []
        self._overflow_entries = []
        self._overflow_index_by_key = {}
        self.overflowLegendChanged.emit([])
        self.setMinimumHeight(240)
        self.update()

    def set_zoom(self, zoom: float) -> None:
        self._zoom = max(1.0, min(2.0, float(zoom)))
        row_h = int(52 * self._zoom)
        self.setMinimumHeight(max(240, 48 + len(self._rows) * row_h))
        self._rebuild_overflow_index()
        self.update()

    def resizeEvent(self, event) -> None:  # noqa: N802
        self._rebuild_overflow_index()
        super().resizeEvent(event)

    def _rebuild_overflow_index(self) -> None:
        z = self._zoom
        xw = max(400.0, float(self.width() - 180))
        bar_h = 22.0 * z
        min_font = 4
        max_font = max(6, int(round(6 * z)))
        entries: list[dict[str, object]] = []
        idx_by_key: dict[tuple[int, int], int] = {}
        for row in self._rows:
            opening = int(row.get("opening", 0))
            bar_len = int(row.get("bar_len", 0))
            segs = row.get("segments", [])
            if bar_len <= 0 or not isinstance(segs, list):
                continue
            profile_idx = 0
            for seg in segs:
                if not isinstance(seg, dict):
                    continue
                if str(seg.get("kind", "")) != "profile":
                    continue
                length_mm = float(seg.get("length_mm", 0.0))
                if length_mm <= 0:
                    profile_idx += 1
                    continue
                sw = xw * (length_mm / bar_len)
                label = str(seg.get("label", ""))
                left_angle = str(seg.get("left_angle", ""))
                right_angle = str(seg.get("right_angle", ""))
                fit_label = f"{label} L{left_angle} R{right_angle}".strip()
                full_fits = False
                if sw >= 14:
                    max_w = max(8.0, sw - 4.0)
                    max_h = max(8.0, bar_h - 4.0)
                    for s in range(max_font, min_font - 1, -1):
                        f = QFont("Segoe UI", s)
                        m = QFontMetrics(f)
                        if m.horizontalAdvance(fit_label) <= max_w and m.height() <= max_h:
                            full_fits = True
                            break
                if (sw < 14) or (not full_fits):
                    key = (opening, profile_idx)
                    idx = len(entries) + 1
                    idx_by_key[key] = idx
                    entries.append(
                        {
                            "id": idx,
                            "opening": opening,
                            "module": str(seg.get("module_name", "")),
                            "profile": str(seg.get("profile_code", "")),
                            "length_mm": int(seg.get("part_length_mm", int(length_mm))),
                            "left_angle": str(seg.get("left_angle", "")),
                            "right_angle": str(seg.get("right_angle", "")),
                            "source": str(seg.get("source_label", "")),
                            "color_rgb": tuple(seg.get("opening_color", (241, 245, 249))),
                        }
                    )
                profile_idx += 1
        if entries != self._overflow_entries:
            self._overflow_entries = entries
            self.overflowLegendChanged.emit(entries)
        self._overflow_index_by_key = idx_by_key

    def paintEvent(self, _event) -> None:  # noqa: N802
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.fillRect(self.rect(), QColor(248, 250, 252))
        z = self._zoom
        p.setPen(QColor(51, 65, 85))
        p.setFont(QFont("Segoe UI", max(7, int(round(8 * z)))))

        if not self._rows:
            p.drawText(16, 28, "Нет данных для схемы раскроя")
            return

        x0 = 160.0
        xw = max(400.0, float(self.width() - 180))
        y = 20.0
        bar_h = 22.0 * z
        row_gap_above_bar = 14.0 * z
        row_gap_below_bar = 16.0 * z
        tech_label_h = 11.0 * z
        for row in self._rows:
            opening = int(row.get("opening", 0))
            bar_len = int(row.get("bar_len", 0))
            remainder = int(row.get("remainder", 0))
            segs = row.get("segments", [])
            if not isinstance(segs, list) or bar_len <= 0:
                continue

            bar_y = y + row_gap_above_bar

            p.setPen(QColor(30, 41, 59))
            p.drawText(12, int(y + 10 * z), f"Пруток {opening} ({bar_len} мм)")
            oc = opening_row_color(opening)
            p.setPen(Qt.NoPen)
            p.setBrush(oc)
            p.drawRect(QRectF(8, bar_y + 1, 3, bar_h - 2))
            scrap_marks = {
                str(seg.get("origin", ""))
                for seg in segs
                if isinstance(seg, dict) and str(seg.get("kind", "")) == "profile" and bool(seg.get("is_scrap"))
            }
            if scrap_marks:
                p.setPen(QColor(71, 85, 105))
                p.setFont(QFont("Segoe UI", max(5, int(round(6 * z)))))
                p.drawText(12, int(bar_y + bar_h + 4 * z), "Обрезки: " + ", ".join(sorted(scrap_marks)))

            p.setPen(QPen(QColor(148, 163, 184), 1.0))
            p.setBrush(QColor(255, 255, 255))
            p.drawRect(QRectF(x0, bar_y, xw, bar_h))

            cursor_mm = 0.0
            profile_idx = 0
            for seg in segs:
                if not isinstance(seg, dict):
                    continue
                length_mm = float(seg.get("length_mm", 0.0))
                if length_mm <= 0:
                    continue
                kind = str(seg.get("kind", ""))
                label = str(seg.get("label", ""))
                if kind == "profile":
                    profile_name = str(seg.get("profile_name", ""))
                    color = self._color_map.get(profile_name, QColor(96, 165, 250))
                elif kind == "tech":
                    color = _TECH_VIS_FILL
                else:
                    color = QColor(255, 255, 255)

                sx = x0 + xw * (cursor_mm / bar_len)
                sw = xw * (length_mm / bar_len)
                p.setPen(QPen(QColor(255, 255, 255), 0.8))
                p.setBrush(color)
                if kind == "tech":
                    tech_show = int(seg.get("tech_mm", length_mm))
                    tech_rect = QRectF(sx, bar_y, sw, bar_h)
                    tech_poly = QPolygonF(tech_rect)
                    _paint_diagonal_hatch(
                        p,
                        tech_poly,
                        fill=_TECH_VIS_FILL,
                        line_color=_TECH_VIS_LINE,
                        spacing=max(3.0 * z, sw / 10.0),
                        line_width=max(0.8, 1.0 * z),
                        cross=True,
                    )
                    p.setPen(QPen(_TECH_VIS_LINE, 1.0))
                    p.setBrush(Qt.NoBrush)
                    p.drawRect(tech_rect)
                    p.setPen(_TECH_VIS_LINE)
                    p.setFont(QFont("Segoe UI", max(6, int(round(7 * z))), QFont.Bold))
                    p.drawText(
                        QRectF(sx, bar_y - tech_label_h - 1, sw, tech_label_h),
                        int(Qt.AlignHCenter | Qt.AlignBottom),
                        f"{tech_show}",
                    )
                elif kind == "kerf":
                    if sw >= 1.0:
                        kerf_rect = QRectF(sx, bar_y, sw, bar_h)
                        kerf_poly = QPolygonF(kerf_rect)
                        _paint_diagonal_hatch(
                            p,
                            kerf_poly,
                            fill=_KERF_VIS_FILL,
                            line_color=_KERF_VIS_LINE,
                            spacing=max(2.5 * z, sw / 5.0),
                            line_width=max(0.85, 1.0 * z),
                            cross=True,
                        )
                        p.setPen(QPen(_KERF_VIS_LINE, 1.0))
                        p.setBrush(Qt.NoBrush)
                        p.drawRect(kerf_rect)
                        p.setPen(QPen(_KERF_VIS_LINE, 1.25))
                        p.drawLine(QPointF(sx, bar_y + bar_h), QPointF(sx, bar_y))
                        p.drawLine(QPointF(sx + sw, bar_y + bar_h), QPointF(sx + sw, bar_y))
                        if sw >= 6:
                            p.setPen(_KERF_VIS_LINE.darker(120))
                            p.setFont(QFont("Segoe UI", max(4, int(round(5 * z))), QFont.Bold))
                            p.drawText(
                                QRectF(sx + 1, bar_y + 1, sw - 2, bar_h - 2),
                                int(Qt.AlignCenter),
                                str(int(seg.get("kerf_mm", 4))),
                            )
                else:
                    p.drawRect(QRectF(sx, bar_y, sw, bar_h))
                if kind == "profile" and sw >= 14:
                    p.setPen(QColor(15, 23, 42))
                    max_w = max(8.0, sw - 4.0)
                    max_h = max(8.0, bar_h - 4.0)
                    chosen: QFont | None = None
                    text = label
                    for s in range(max(6, int(round(6 * z))), 3, -1):
                        f = QFont("Segoe UI", s)
                        m = QFontMetrics(f)
                        t = m.elidedText(label, Qt.TextElideMode.ElideRight, int(max_w))
                        if m.horizontalAdvance(t) <= max_w and m.height() <= max_h:
                            chosen = f
                            text = t
                            break
                    if chosen is None:
                        chosen = QFont("Segoe UI", 4)
                        text = QFontMetrics(chosen).elidedText(
                            label,
                            Qt.TextElideMode.ElideRight,
                            int(max_w),
                        )
                    p.setFont(chosen)
                    p.drawText(QRectF(sx + 2, bar_y + 2, sw - 4, bar_h - 4), int(Qt.AlignCenter), text)
                    left_a = int(seg.get("left_angle", 90))
                    right_a = int(seg.get("right_angle", left_a))
                    if sw >= 28:
                        p.setPen(QColor(30, 41, 59))
                        p.setFont(QFont("Segoe UI", max(4, int(round(5 * z)))))
                        p.drawText(
                            QRectF(sx + 1, bar_y + bar_h - 10 * z, max(sw * 0.48, 8), 9 * z),
                            int(Qt.AlignLeft | Qt.AlignVCenter),
                            f"L{left_a}°",
                        )
                        p.drawText(
                            QRectF(sx + sw * 0.52, bar_y + bar_h - 10 * z, max(sw * 0.46, 8), 9 * z),
                            int(Qt.AlignRight | Qt.AlignVCenter),
                            f"R{right_a}°",
                        )
                    else:
                        p.setPen(QColor(51, 65, 85))
                        p.setFont(QFont("Segoe UI", max(4, int(round(5 * z)))))
                        p.drawText(
                            QRectF(sx - 18 * z, bar_y - tech_label_h - 1, 24 * z, tech_label_h),
                            int(Qt.AlignLeft | Qt.AlignBottom),
                            f"L{left_a}°",
                        )
                        p.drawText(
                            QRectF(sx + sw - 6 * z, bar_y - tech_label_h - 1, 24 * z, tech_label_h),
                            int(Qt.AlignRight | Qt.AlignBottom),
                            f"R{right_a}°",
                        )
                if kind == "profile":
                    oid = self._overflow_index_by_key.get((opening, profile_idx))
                    if oid is not None:
                        p.setPen(QColor(15, 23, 42))
                        p.setFont(QFont("Segoe UI", max(5, int(round(6 * z))), QFont.Bold))
                        p.drawText(
                            QRectF(sx + 1, bar_y + 1, max(sw - 2, 8), bar_h - 2),
                            int(Qt.AlignCenter),
                            f"#{oid}",
                        )
                    profile_idx += 1
                if kind == "profile" and bool(seg.get("is_scrap")):
                    p.setPen(QColor(100, 116, 139))
                    p.setBrush(Qt.NoBrush)
                    p.drawRect(QRectF(sx + 0.8, bar_y + 0.8, max(sw - 1.6, 1.0), max(bar_h - 1.6, 1.0)))
                cursor_mm += length_mm

            if remainder > 0:
                rx = x0 + xw * ((bar_len - remainder) / bar_len)
                rw = xw * (remainder / bar_len)
                p.setPen(QPen(QColor(148, 163, 184), 1.0, Qt.PenStyle.DashLine))
                p.setBrush(Qt.NoBrush)
                p.drawRect(QRectF(rx, bar_y, rw, bar_h))
                p.setPen(QColor(71, 85, 105))
                p.setFont(QFont("Segoe UI", max(6, int(round(6 * z)))))
                p.drawText(
                    QRectF(rx + 2, bar_y + 2, max(rw - 4, 12), bar_h - 4),
                    int(Qt.AlignCenter),
                    f"остаток {remainder}",
                )
            y += row_gap_above_bar + bar_h + row_gap_below_bar


class JointAlbumWidget(QWidget):
    """Крупный альбом стыков соседних деталей."""

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._rows: list[dict[str, object]] = []
        self._color_map: dict[str, QColor] = {}
        self.setMinimumHeight(280)

    def set_rows(self, rows: list[dict[str, object]], color_map: dict[str, QColor]) -> None:
        self._rows = rows
        self._color_map = color_map
        self.setMinimumHeight(max(280, 24 + len(rows) * 126))
        self.setMinimumWidth(720)
        self.updateGeometry()
        self.update()

    def clear_rows(self) -> None:
        self._rows = []
        self.setMinimumHeight(280)
        self.update()

    def paintEvent(self, _event) -> None:  # noqa: N802
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.fillRect(self.rect(), QColor(248, 250, 252))
        if not self._rows:
            p.setPen(QColor(51, 65, 85))
            p.setFont(QFont("Segoe UI", 9))
            p.drawText(14, 28, "Нет данных для альбома стыков")
            return

        y = 12.0
        row_h = 116.0
        left_x = 140.0
        detail_w = max(180.0, (self.width() - 220.0) / 2.0)
        detail_h = 56.0

        def _draw_detail(
            rect: QRectF,
            fill_color: QColor,
            title: str,
            left_angle_deg: int,
            right_angle_deg: int,
            tech_mm: int,
            kerf_mm: int,
        ) -> None:
            yt = rect.top()
            yb = rect.bottom()
            # Левый техотступ и вертикальный пропил; справа — только пропил (схема 1D).
            ltw = 16.0 if tech_mm == 50 else 10.0
            kw = 6.0 if kerf_mm >= 4 else 4.0
            left_zone_end = rect.left() + ltw + kw
            right_zone_start = rect.right() - kw

            tech_rect = QRectF(rect.left(), yt, ltw, rect.height())
            kerf_l = QRectF(rect.left() + ltw, yt, kw, rect.height())
            kerf_r = QRectF(right_zone_start, yt, rect.right() - right_zone_start, rect.height())
            body = QRectF(left_zone_end, yt, right_zone_start - left_zone_end, rect.height())

            p.setPen(QPen(QColor(241, 245, 249), 1.0))
            p.setBrush(fill_color)
            p.drawRect(rect)

            _paint_diagonal_hatch(
                p,
                QPolygonF(tech_rect),
                fill=_TECH_VIS_FILL,
                line_color=_TECH_VIS_LINE,
                spacing=4.0,
                line_width=1.0,
                cross=True,
            )
            p.setPen(QPen(_TECH_VIS_LINE, 1.0))
            p.setBrush(Qt.NoBrush)
            p.drawRect(tech_rect)
            p.setPen(_TECH_VIS_LINE)
            p.setFont(QFont("Segoe UI", 8, QFont.Bold))
            p.drawText(
                QRectF(tech_rect.left(), yt - 16, tech_rect.width(), 14),
                int(Qt.AlignHCenter | Qt.AlignBottom),
                str(int(tech_mm)),
            )

            _paint_diagonal_hatch(
                p,
                QPolygonF(kerf_l),
                fill=_KERF_VIS_FILL,
                line_color=_KERF_VIS_LINE,
                spacing=max(3.0, kw / 2.0),
                line_width=1.0,
                cross=True,
            )
            p.setPen(QPen(_KERF_VIS_LINE, 1.0))
            p.drawRect(kerf_l)
            p.drawLine(QPointF(kerf_l.left(), yb), QPointF(kerf_l.left(), yt))
            p.drawLine(QPointF(kerf_l.right(), yb), QPointF(kerf_l.right(), yt))
            p.setPen(_KERF_VIS_LINE.darker(120))
            p.setFont(QFont("Segoe UI", 7, QFont.Bold))
            p.drawText(kerf_l.adjusted(0, 0, 0, 0), int(Qt.AlignCenter), str(int(kerf_mm)))

            _paint_diagonal_hatch(
                p,
                QPolygonF(kerf_r),
                fill=_KERF_VIS_FILL,
                line_color=_KERF_VIS_LINE,
                spacing=max(3.0, kw / 2.0),
                line_width=1.0,
                cross=True,
            )
            p.setPen(QPen(_KERF_VIS_LINE, 1.0))
            p.drawRect(kerf_r)
            p.drawLine(QPointF(kerf_r.left(), yb), QPointF(kerf_r.left(), yt))
            p.drawLine(QPointF(kerf_r.right(), yb), QPointF(kerf_r.right(), yt))
            p.setPen(_KERF_VIS_LINE.darker(120))
            p.drawText(kerf_r, int(Qt.AlignCenter), str(int(kerf_mm)))

            p.setPen(QPen(fill_color.darker(115), 1.0))
            p.setBrush(fill_color)
            p.drawRect(body)

            p.setPen(QColor(15, 23, 42))
            p.setFont(QFont("Segoe UI", 8))
            fm = QFontMetrics(p.font())
            p.drawText(
                body.adjusted(4, 2, -4, -2),
                int(Qt.AlignCenter),
                fm.elidedText(title, Qt.TextElideMode.ElideRight, int(body.width() - 8)),
            )

            p.setPen(QColor(51, 65, 85))
            p.setFont(QFont("Segoe UI", 8, QFont.Bold))
            p.drawText(int(rect.left() + 2), int(rect.bottom() + 16), f"L: {left_angle_deg}°")
            p.drawText(int(rect.right() - 56), int(rect.bottom() + 16), f"R: {right_angle_deg}°")

        for r in self._rows:
            opening = int(r.get("opening", 0))
            kind = str(r.get("kind", "joint"))
            left_title = str(r.get("left_title", ""))
            right_title = str(r.get("right_title", ""))
            left_profile = str(r.get("left_profile_name", ""))
            right_profile = str(r.get("right_profile_name", ""))
            left_angle = int(r.get("left_right_angle", 90))
            right_angle = int(r.get("right_left_angle", 90))
            left_left_angle = int(r.get("left_left_angle", left_angle))
            right_right_angle = int(r.get("right_right_angle", right_angle))

            block = QRectF(8, y, self.width() - 16, row_h)
            p.setPen(QPen(QColor(203, 213, 225), 1.0))
            p.setBrush(QColor(255, 255, 255))
            p.drawRoundedRect(block, 6, 6)
            p.setPen(QColor(30, 41, 59))
            p.setFont(QFont("Segoe UI", 9, QFont.Bold))
            p.drawText(
                16,
                int(y + 18),
                f"Пруток {opening}: {'стык' if kind == 'joint' else 'деталь'}",
            )

            ly = y + 36.0
            left_rect = QRectF(left_x, ly, detail_w, detail_h)
            seam_x = left_rect.right()
            right_rect = QRectF(seam_x, ly, detail_w, detail_h)
            left_tech = int(r.get("left_tech_mm", 30))
            right_tech = int(r.get("right_tech_mm", 30))
            kerf_mm = int(r.get("kerf_mm", 4))

            lc = self._color_map.get(left_profile, QColor(147, 197, 253))
            rc = self._color_map.get(right_profile, QColor(134, 239, 172))
            _draw_detail(left_rect, lc, left_title, left_left_angle, left_angle, left_tech, kerf_mm)
            _draw_detail(right_rect, rc, right_title, right_angle, right_right_angle, right_tech, kerf_mm)

            p.setPen(QPen(QColor(15, 23, 42), 1.2, Qt.PenStyle.DashLine))
            p.drawLine(QPointF(seam_x, ly - 3), QPointF(seam_x, ly + detail_h + 3))
            y += row_h + 10.0


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"NordFox — раскрой v{__version__}")
        self.resize(1180, 720)
        self.setStyleSheet(
            """
            QWidget { background: #f8fafc; color: #0f172a; }
            QGroupBox {
                border: 1px solid #dbe3ee;
                border-radius: 10px;
                margin-top: 10px;
                padding-top: 8px;
                font-weight: 600;
            }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
            QLineEdit, QComboBox, QTextEdit, QTableWidget {
                background: #ffffff;
                border: 1px solid #dbe3ee;
                border-radius: 8px;
                padding: 4px;
            }
            QPushButton {
                background: #ffffff;
                border: 1px solid #cfd8e3;
                border-radius: 8px;
                padding: 6px 10px;
            }
            QPushButton:hover { background: #f1f5f9; }
            """
        )

        self._last_sorted_cuts: list[CutEvent] | None = None
        self._optimizer_cuts: list[CutEvent] | None = None
        self._last_summary_text: str = ""
        self._recommended_bars: tuple[int, ...] | None = None
        self._footer_row_index: int | None = None
        self._data_row_count: int = 0
        self._selected_bars_mm: tuple[int, ...] | None = None
        self._last_kerf_mm: int = 0
        self._layout_rows: list[dict[str, object]] = []
        self._album_rows: list[dict[str, object]] = []
        self._table_zoom = 1.0
        self._layout_zoom = 1.0
        self.log_emitter = LogEmitter()
        logger.info("MainWindow init")

        root = QWidget()
        self.setCentralWidget(root)
        layout = QHBoxLayout(root)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        layout.addWidget(splitter, 1)

        left_wrap = QWidget()
        left_wrap.setMinimumWidth(280)
        left_wrap.setMaximumWidth(420)
        left_wrap.setSizePolicy(
            QSizePolicy.Policy.Preferred,
            QSizePolicy.Policy.Expanding,
        )
        left_layout = QVBoxLayout(left_wrap)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        right_wrap = QWidget()
        right_layout = QVBoxLayout(right_wrap)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(8)

        logs_wrap = QWidget()
        logs_wrap.setMinimumWidth(260)
        logs_layout = QVBoxLayout(logs_wrap)
        logs_layout.setContentsMargins(0, 0, 0, 0)
        logs_layout.setSpacing(6)

        splitter.addWidget(left_wrap)
        splitter.addWidget(right_wrap)
        splitter.addWidget(logs_wrap)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 6)
        splitter.setStretchFactor(2, 2)
        splitter.setChildrenCollapsible(False)
        splitter.setSizes([320, 980, 360])

        top = QHBoxLayout()
        left_layout.addLayout(top)
        top.addWidget(QLabel("Спецификация (.xlsx):"))
        self.path_edit = QLineEdit()
        self.path_edit.setMinimumWidth(120)
        self._bind_line_edit_undo_redo(self.path_edit)
        top.addWidget(self.path_edit, stretch=1)
        browse = QPushButton("Обзор…")
        browse.clicked.connect(self._browse)
        top.addWidget(browse)

        scrap_row = QHBoxLayout()
        left_layout.addLayout(scrap_row)
        scrap_row.addWidget(QLabel("Склад обрезков (.xlsx):"))
        self.scrap_path_edit = QLineEdit()
        self.scrap_path_edit.setMinimumWidth(120)
        self._bind_line_edit_undo_redo(self.scrap_path_edit)
        scrap_row.addWidget(self.scrap_path_edit, stretch=1)
        sb = QPushButton("Обзор…")
        sb.clicked.connect(self._browse_scrap)
        scrap_row.addWidget(sb)

        prof = QGroupBox(
            "Профили в раскрое (СК-/СС-/Р-: 0=Н20, 1=Н21, 2=Н22, 3=Н23)"
        )
        prof.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        prof.setMaximumHeight(130)
        pg = QGridLayout(prof)
        pg.setHorizontalSpacing(10)
        pg.setVerticalSpacing(4)
        self.profile_checks: dict[int, QCheckBox] = {}
        for i, name in PROFILE_DIGIT_TO_NAME.items():
            cb = QCheckBox(f"{name} (цифра {i})")
            cb.setChecked(True)
            self.profile_checks[i] = cb
            pg.addWidget(cb, i // 4, i % 4)
        left_layout.addWidget(prof)

        bar_fr = QGroupBox("Длина заготовки (мм)")
        bar_fr.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        bar_fr.setMaximumHeight(72)
        bl = QVBoxLayout(bar_fr)
        bl.setSpacing(6)
        bar_row = QHBoxLayout()
        bl.addLayout(bar_row)
        bar_row.addWidget(QLabel("Базовая длина:"))
        self.base_bar_edit = QLineEdit("6000")
        self.base_bar_edit.setMaximumWidth(100)
        self._bind_line_edit_undo_redo(self.base_bar_edit)
        self.base_bar_edit.setToolTip(
            "Длина заготовки для расчёта (не более 12000 мм). "
            "Число можно менять сколько угодно и снова запускать расчёт. "
            "Ctrl+Z / Ctrl+Y — отмена и повтор ввода в этом поле."
        )
        bar_row.addWidget(self.base_bar_edit)
        bar_row.addStretch(1)
        left_layout.addWidget(bar_fr)

        opt = QHBoxLayout()
        left_layout.addLayout(opt)
        opt.addWidget(QLabel("Пропил (kerf), мм:"))
        self.kerf_edit = QLineEdit("0")
        self.kerf_edit.setMaximumWidth(72)
        self._bind_line_edit_undo_redo(self.kerf_edit)
        opt.addWidget(self.kerf_edit)
        opt.addStretch()
        tech_rule = QLabel("Тех. отступ: по 1-му углу (90°=30 мм, иначе 50 мм), пропил = 2×kerf")
        tech_rule.setStyleSheet("color: #475569;")
        tech_rule.setWordWrap(True)
        tech_rule.setToolTip(
            "Отступ берётся только по первому углу (90°=30 мм, иначе 50 мм). "
            "Пропил учитывается как 2×kerf на деталь."
        )
        left_layout.addWidget(tech_rule)

        btn_row = QHBoxLayout()
        left_layout.addLayout(btn_row)
        run_btn = QPushButton("Рассчитать раскрой")
        run_btn.clicked.connect(self._compute)
        btn_row.addWidget(run_btn, 1)
        self.btn_xlsx = QPushButton("Экспорт Excel…")
        self.btn_xlsx.clicked.connect(self._export_excel)
        self.btn_xlsx.setEnabled(False)
        btn_row.addWidget(self.btn_xlsx, 1)
        self.btn_pdf = QPushButton("Экспорт PDF…")
        self.btn_pdf.clicked.connect(self._export_pdf)
        self.btn_pdf.setEnabled(False)
        btn_row.addWidget(self.btn_pdf, 1)

        btn_row2 = QHBoxLayout()
        left_layout.addLayout(btn_row2)
        self.btn_recalc = QPushButton("Пересчитать по таблице")
        self.btn_recalc.setToolTip(
            "Учитываются колонки: модуль, тип профиля, длина, угол. "
            "Колонка «Масса» только для отображения. Остальные после пересчёта обновятся."
        )
        self.btn_recalc.clicked.connect(self._recalc_from_table)
        self.btn_recalc.setEnabled(False)
        btn_row2.addWidget(self.btn_recalc, 1)
        btn_adv = QPushButton("Подбор заготовок")
        btn_adv.setToolTip(
            "Сравнивает типовые наборы длин (только 6 м, только 7,5 м, комбинации…); "
            "учитывает склад обрезков, если указан файл."
        )
        btn_adv.clicked.connect(self._run_bar_advisor)
        btn_row2.addWidget(btn_adv, 1)
        btn_apply_rec = QPushButton("Отметить рекоменд. длины")
        btn_apply_rec.setToolTip("Ставит галочки по последней рекомендации из блока ниже")
        btn_apply_rec.clicked.connect(self._apply_recommended_bars)
        btn_row2.addWidget(btn_apply_rec, 1)

        chart_row = QHBoxLayout()
        left_layout.addLayout(chart_row)
        self.mass_chart = RingBreakdownWidget()
        self.mass_chart.setMaximumHeight(180)
        chart_row.addWidget(self.mass_chart, 1)
        self.waste_chart = RingBreakdownWidget()
        self.waste_chart.setMaximumHeight(180)
        chart_row.addWidget(self.waste_chart, 1)

        self.advisor_text = QTextEdit()
        self.advisor_text.setReadOnly(True)
        self.advisor_text.setMinimumHeight(100)
        self.advisor_text.setSizePolicy(
            QSizePolicy.Policy.Preferred,
            QSizePolicy.Policy.Expanding,
        )
        self.advisor_text.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
            | Qt.TextInteractionFlag.TextSelectableByKeyboard
        )
        self.advisor_text.setPlaceholderText(
            "Нажмите «Подбор заготовок» для сравнения вариантов 6 / 7,5 / 12 м по текущей спецификации."
        )
        self.advisor_text.setFont(QFont("Segoe UI", 9))
        left_layout.addWidget(self.advisor_text, stretch=1)

        self.right_tabs = QTabWidget()
        right_layout.addWidget(self.right_tabs, 1)

        table_tab = QWidget()
        table_layout = QVBoxLayout(table_tab)
        table_layout.setContentsMargins(4, 4, 4, 4)
        table_layout.setSpacing(6)
        sort_row = QHBoxLayout()
        table_layout.addLayout(sort_row)
        sort_row.addWidget(QLabel("Сортировка таблицы:"))
        self.sort_combo = QComboBox()
        for mode_id, label in SORT_MODES:
            self.sort_combo.addItem(label, mode_id)
        self.sort_combo.currentIndexChanged.connect(self._apply_sort)
        sort_row.addWidget(self.sort_combo, stretch=1)

        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels(
            [
                "Пруток №",
                "Модуль",
                "Тип профиля",
                "Серия",
                "Длина",
                "Угол",
                "Источник",
                "Заготовка мм",
                "Остаток мм",
                "Масса профиля, кг",
            ]
        )
        self.table.setAlternatingRowColors(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.SelectedClicked
        )
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_table_context_menu)
        hdr = self.table.horizontalHeader()
        hdr.setStretchLastSection(False)
        for i in range(self.table.columnCount()):
            hdr.setSectionResizeMode(i, QHeaderView.Interactive)
        hdr.setSectionsClickable(True)
        hdr.setCascadingSectionResizes(False)
        hdr.setMinimumSectionSize(70)
        self.table.setFont(QFont("Segoe UI", 10))
        self.table.viewport().installEventFilter(self)
        table_layout.addWidget(self.table, 1)
        self.right_tabs.addTab(table_tab, "Таблица")

        layout_tab = QWidget()
        layout_tab_l = QVBoxLayout(layout_tab)
        layout_tab_l.setContentsMargins(4, 4, 4, 4)
        layout_tab_l.setSpacing(6)
        plan_btns = QHBoxLayout()
        layout_tab_l.addLayout(plan_btns)
        self.btn_layout_xlsx = QPushButton("Экспорт схемы Excel…")
        self.btn_layout_xlsx.setEnabled(False)
        self.btn_layout_xlsx.clicked.connect(self._export_layout_excel)
        plan_btns.addWidget(self.btn_layout_xlsx)
        self.btn_layout_pdf = QPushButton("Экспорт схемы PDF…")
        self.btn_layout_pdf.setEnabled(False)
        self.btn_layout_pdf.clicked.connect(self._export_layout_pdf)
        plan_btns.addWidget(self.btn_layout_pdf)
        plan_btns.addStretch(1)
        self.layout_hint = QLabel(
            "Сегменты: техотступ (оранж.) — мм над полосой; пропил (бирюза), вертикальная полоса; профиль — цвет серии. "
            "Углы только подписью (L/R). Пунктир: деталь из обрезка. Узкие сегменты — метки #N в таблице ниже."
        )
        self.layout_hint.setStyleSheet("color: #475569;")
        layout_tab_l.addWidget(self.layout_hint)
        self.layout_widget = CuttingLayoutWidget()
        self.layout_widget.overflowLegendChanged.connect(self._populate_layout_overflow_table)
        self.layout_scroll = QScrollArea()
        self.layout_scroll.setWidgetResizable(True)
        self.layout_scroll.setWidget(self.layout_widget)
        self.layout_scroll.viewport().installEventFilter(self)
        self.layout_widget.installEventFilter(self)
        layout_tab_l.addWidget(self.layout_scroll, 1)
        self.layout_overflow_table = QTableWidget()
        self.layout_overflow_table.setColumnCount(8)
        self.layout_overflow_table.setHorizontalHeaderLabels(
            ["#ID", "Пруток №", "Модуль", "Профиль", "Длина, мм", "L, °", "R, °", "Источник"]
        )
        self.layout_overflow_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.layout_overflow_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.layout_overflow_table.setAlternatingRowColors(False)
        self.layout_overflow_table.setShowGrid(True)
        self.layout_overflow_table.setGridStyle(Qt.PenStyle.SolidLine)
        self.layout_overflow_table.setMaximumHeight(180)
        oh = self.layout_overflow_table.horizontalHeader()
        for i in range(self.layout_overflow_table.columnCount()):
            oh.setSectionResizeMode(i, QHeaderView.Interactive)
        layout_tab_l.addWidget(self.layout_overflow_table)
        self.right_tabs.addTab(layout_tab, "Схема раскроя")

        album_tab = QWidget()
        album_l = QVBoxLayout(album_tab)
        album_l.setContentsMargins(4, 4, 4, 4)
        album_l.setSpacing(6)
        album_btns = QHBoxLayout()
        album_l.addLayout(album_btns)
        album_btns.addWidget(QLabel("Режим:"))
        self.album_mode_combo = QComboBox()
        self.album_mode_combo.addItem("Стыки", "joints")
        self.album_mode_combo.addItem("Детали", "details")
        self.album_mode_combo.currentIndexChanged.connect(self._refresh_album)
        album_btns.addWidget(self.album_mode_combo)
        self.btn_album_pdf = QPushButton("Экспорт альбома PDF…")
        self.btn_album_pdf.setEnabled(False)
        self.btn_album_pdf.clicked.connect(self._export_album_pdf)
        album_btns.addWidget(self.btn_album_pdf)
        album_btns.addStretch(1)
        album_hint = QLabel(
            "Крупный план стыков: каждая строка показывает соседние детали и геометрию углов на общем резе."
        )
        album_hint.setStyleSheet("color: #475569;")
        album_l.addWidget(album_hint)
        self.album_widget = JointAlbumWidget()
        self.album_scroll = QScrollArea()
        # False: иначе высота альбома сжимается до окна и видна только одна строка.
        self.album_scroll.setWidgetResizable(False)
        self.album_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.album_scroll.setWidget(self.album_widget)
        album_l.addWidget(self.album_scroll, 1)
        self.right_tabs.addTab(album_tab, "Альбом стыков")
        self._copy_shortcut = QShortcut(QKeySequence.Copy, self)
        self._copy_shortcut.setContext(Qt.ShortcutContext.ApplicationShortcut)
        self._copy_shortcut.activated.connect(self._copy_from_focused_widget)

        logs_layout.addWidget(QLabel("Логи в реальном времени"))
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.log_view.setFont(QFont("Consolas", 9))
        self.log_view.setPlaceholderText("Логи расчётов и ошибок появятся здесь…")
        logs_layout.addWidget(self.log_view, 1)
        logs_btn_row = QHBoxLayout()
        self.btn_save_log = QPushButton("Сохранить лог")
        self.btn_save_log.clicked.connect(self._save_live_log)
        logs_btn_row.addWidget(self.btn_save_log)
        self.btn_clear_log = QPushButton("Очистить лог")
        self.btn_clear_log.clicked.connect(self._clear_live_log)
        logs_btn_row.addWidget(self.btn_clear_log)
        logs_layout.addLayout(logs_btn_row)

        self._setup_live_log_handler()
        self._enable_text_copy(self.advisor_text)
        self._enable_text_copy(self.log_view)

        proj_root = Path(__file__).resolve().parents[2]
        for name in (
            "spec_20_modules_block.xlsx",
            "spec_20x5_modules.xlsx",
            "spec_10x5_modules.xlsx",
        ):
            p = proj_root / "test" / name
            if p.is_file():
                self.path_edit.setText(str(p))
                break

    def _browse(self) -> None:
        base = self.path_edit.text().strip()
        p, _ = QFileDialog.getOpenFileName(
            self,
            "Спецификация",
            str(Path(base).parent) if base else "",
            "Excel (*.xlsx);;Все файлы (*.*)",
        )
        if p:
            self.path_edit.setText(p)

    def _setup_live_log_handler(self) -> None:
        self.log_emitter.message.connect(self._append_live_log)
        app_logger = logging.getLogger("nordfox_raskroy")
        # Избегаем повторного подключения при повторном создании окна.
        for h in app_logger.handlers:
            if isinstance(h, QtLogHandler):
                h.emitter = self.log_emitter
                return
        h = QtLogHandler(self.log_emitter)
        h.setFormatter(
            logging.Formatter(
                "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
                datefmt="%H:%M:%S",
            )
        )
        h.setLevel(logging.INFO)
        app_logger.addHandler(h)

    def _bind_line_edit_undo_redo(self, line_edit: QLineEdit) -> None:
        """Ctrl+Z / Ctrl+Y на самом поле (WidgetShortcut), чтобы отмена ввода работала стабильно."""
        us = QShortcut(QKeySequence.Undo, line_edit)
        us.setContext(Qt.ShortcutContext.WidgetShortcut)
        us.activated.connect(line_edit.undo)
        rs = QShortcut(QKeySequence.Redo, line_edit)
        rs.setContext(Qt.ShortcutContext.WidgetShortcut)
        rs.activated.connect(line_edit.redo)

    def _enable_text_copy(self, widget: QWidget) -> None:
        """Контекстное меню для текстовых панелей (копирование — через ApplicationShortcut)."""
        if hasattr(widget, "setContextMenuPolicy"):
            widget.setContextMenuPolicy(Qt.ContextMenuPolicy.DefaultContextMenu)

    def _ancestor_text_edit(self, w: QWidget | None) -> QTextEdit | QPlainTextEdit | None:
        """Фокус часто на дочернем виджете QTextEdit — поднимаемся к редактору."""
        cur: QWidget | None = w
        for _ in range(24):
            if isinstance(cur, (QTextEdit, QPlainTextEdit)):
                return cur
            if cur is None:
                return None
            cur = cur.parentWidget()
        return None

    def _append_live_log(self, text: str) -> None:
        self.log_view.appendPlainText(text)
        self.log_view.verticalScrollBar().setValue(
            self.log_view.verticalScrollBar().maximum()
        )

    def _save_live_log(self) -> None:
        text = self.log_view.toPlainText().strip()
        if not text:
            QMessageBox.information(self, "Логи", "Панель логов пуста.")
            return
        base = self.path_edit.text().strip()
        default_dir = str(Path(base).parent) if base else ""
        p, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить лог",
            str(Path(default_dir) / "nordfox_raskroy_runtime.log") if default_dir else "",
            "Лог (*.log *.txt);;Все файлы (*.*)",
        )
        if not p:
            return
        try:
            Path(p).write_text(text + "\n", encoding="utf-8")
        except Exception as e:  # noqa: BLE001
            logger.exception("Save live log failed")
            QMessageBox.critical(self, "Логи", str(e))
            return
        QMessageBox.information(self, "Логи", f"Сохранено:\n{p}")

    def _clear_live_log(self) -> None:
        self.log_view.clear()
        logger.info("Live log view cleared by user")

    def eventFilter(self, obj: QObject, event: QEvent) -> bool:  # noqa: N802
        if event.type() == QEvent.Type.Wheel:
            tab_idx = self.right_tabs.currentIndex() if hasattr(self, "right_tabs") else 0
            dy = event.angleDelta().y() if hasattr(event, "angleDelta") else 0
            mods = event.modifiers() if hasattr(event, "modifiers") else Qt.KeyboardModifier.NoModifier
            ctrl_pressed = bool(mods & Qt.KeyboardModifier.ControlModifier)
            if dy != 0 and ctrl_pressed and obj is self.table.viewport() and tab_idx == 0:
                step = 0.1 if dy > 0 else -0.1
                self._set_table_zoom(self._table_zoom + step)
                return True
            if (
                dy != 0
                and ctrl_pressed
                and tab_idx == 1
                and (obj is self.layout_scroll.viewport() or obj is self.layout_widget)
            ):
                step = 0.1 if dy > 0 else -0.1
                self._set_layout_zoom(self._layout_zoom + step)
                return True
        return super().eventFilter(obj, event)

    def _set_table_zoom(self, zoom: float) -> None:
        self._table_zoom = max(1.0, min(2.0, float(zoom)))
        body_sz = max(9, int(round(10 * self._table_zoom)))
        head_sz = max(8, int(round(9 * self._table_zoom)))
        self.table.setFont(QFont("Segoe UI", body_sz))
        self.table.horizontalHeader().setFont(QFont("Segoe UI", head_sz, QFont.Bold))
        for r in range(self.table.rowCount()):
            self.table.setRowHeight(r, max(22, int(round(22 * self._table_zoom))))

    def _set_layout_zoom(self, zoom: float) -> None:
        self._layout_zoom = max(1.0, min(2.0, float(zoom)))
        self.layout_widget.set_zoom(self._layout_zoom)

    def _apply_optimization_result(
        self,
        *,
        result: OptimizationResult,
        kerf_mm: int,
        chart_metrics: dict[str, object] | None,
    ) -> None:
        """
        Единая точка обновления всех представлений после любого расчёта.
        Обновляет раскрой, схему, альбом и массовые метрики синхронно.
        """
        self._optimizer_cuts = list(result.cuts)
        self._last_kerf_mm = kerf_mm
        self.btn_xlsx.setEnabled(True)
        self.btn_pdf.setEnabled(True)
        self.btn_recalc.setEnabled(True)
        self.btn_layout_xlsx.setEnabled(True)
        self.btn_layout_pdf.setEnabled(True)
        self.btn_album_pdf.setEnabled(True)
        self._update_layout_plan(result.cuts, kerf_mm=kerf_mm)
        self._update_album_plan(result.cuts)
        self._apply_chart_metrics(chart_metrics)
        self._apply_sort()

    def _populate_layout_overflow_table(self, rows: list[dict[str, object]]) -> None:
        self.layout_overflow_table.setRowCount(0)
        for r in rows:
            row = self.layout_overflow_table.rowCount()
            self.layout_overflow_table.insertRow(row)
            vals = [
                str(r.get("id", "")),
                str(r.get("opening", "")),
                str(r.get("module", "")),
                str(r.get("profile", "")),
                str(r.get("length_mm", "")),
                str(r.get("left_angle", "")),
                str(r.get("right_angle", "")),
                str(r.get("source", "")),
            ]
            rgb = r.get("color_rgb", (241, 245, 249))
            bg = QColor(int(rgb[0]), int(rgb[1]), int(rgb[2])) if isinstance(rgb, tuple) and len(rgb) == 3 else opening_row_color(int(r.get("opening", 0) or 0))
            for c, v in enumerate(vals):
                it = QTableWidgetItem(v)
                it.setBackground(bg)
                self.layout_overflow_table.setItem(row, c, it)

    def _refresh_album(self) -> None:
        if not self._optimizer_cuts:
            self.album_widget.clear_rows()
            self._album_rows = []
            return
        self._update_album_plan(self._optimizer_cuts)

    def _browse_scrap(self) -> None:
        base = self.scrap_path_edit.text().strip()
        p, _ = QFileDialog.getOpenFileName(
            self,
            "Склад обрезков",
            str(Path(base).parent) if base else "",
            "Excel (*.xlsx);;Все файлы (*.*)",
        )
        if p:
            self.scrap_path_edit.setText(p)

    def _load_initial_scraps(self) -> tuple[list[int], list[str]]:
        p = self.scrap_path_edit.text().strip()
        if not p or not Path(p).is_file():
            return [], []
        try:
            return parse_scrap_inventory(Path(p))
        except Exception as e:  # noqa: BLE001
            return [], [str(e)]

    def _selected_bar_lengths(
        self,
        demands: list[PartDemand],
        kerf_mm: int,
    ) -> tuple[list[int], str]:
        try:
            base_len = int(self.base_bar_edit.text().strip() or "0")
        except ValueError:
            return [], "Базовая длина должна быть целым числом"
        if base_len <= 0:
            return [], "Базовая длина должна быть > 0"
        if base_len > 12000:
            return [], "Базовая длина не должна превышать 12000 мм"

        required = [demand_cut_length_mm(d, kerf_mm) for d in demands]
        if required and max(required) > 12000:
            return [], (
                "Есть детали, требующие заготовку больше 12000 мм "
                f"(максимум требуется: {max(required)} мм)"
            )
        if required and max(required) > base_len:
            return [], (
                f"Базовой длины {base_len} мм недостаточно: "
                f"нужно минимум {max(required)} мм"
            )

        self._selected_bars_mm = (base_len,)
        return [base_len], ""

    def _demands_from_spec(self) -> tuple[list[PartDemand] | None, str, list[str]]:
        path = self.path_edit.text().strip()
        if not path or not Path(path).is_file():
            return None, "Укажите файл спецификации", []
        allowed: set[int] = {d for d, cb in self.profile_checks.items() if cb.isChecked()}
        if not allowed:
            return None, "Отметьте хотя бы один профиль Н20–Н23", []
        try:
            rows_all = parse_specification(path)
            rows, warns = filter_spec_by_profiles(rows_all, allowed)
        except Exception as e:  # noqa: BLE001
            return None, str(e), []
        if not rows:
            return None, "После фильтра по профилям нет строк", warns
        return spec_rows_to_demands(rows), "", warns

    def _run_bar_advisor(self) -> None:
        logger.info("Bar advisor requested")
        demands, err, fw = self._demands_from_spec()
        if demands is None:
            QMessageBox.warning(self, "Подбор заготовок", err)
            return
        try:
            kerf = int(self.kerf_edit.text().strip() or "0")
        except ValueError:
            QMessageBox.critical(self, "Ошибка", "Пропил должен быть целым числом")
            return
        try:
            base_len = int(self.base_bar_edit.text().strip() or "0")
        except ValueError:
            QMessageBox.critical(self, "Ошибка", "Базовая длина должна быть целым числом")
            return
        if base_len <= 0:
            QMessageBox.critical(self, "Ошибка", "Базовая длина должна быть > 0")
            return
        if base_len > 12000:
            QMessageBox.critical(self, "Ошибка", "Базовая длина не должна превышать 12000 мм")
            return
        min_scrap = 0
        initial, scrap_warns = self._load_initial_scraps()
        required_max = max((demand_cut_length_mm(d, kerf) for d in demands), default=0)
        if required_max > 12000:
            QMessageBox.critical(
                self,
                "Ошибка",
                "Есть детали, требующие заготовку больше 12000 мм "
                f"(максимум требуется: {required_max} мм)",
            )
            return
        scan_from = max(required_max, base_len - 500)
        scan_to = min(12000, max(base_len + 1000, required_max + 1000))
        candidates = set(range(scan_from, scan_to + 1, 50))
        candidates.add(base_len)
        candidates.add(required_max)
        candidates = {x for x in candidates if required_max <= x <= 12000}
        scenarios = [(f"Только {b} мм", (b,)) for b in sorted(candidates)]
        outcomes = compare_bar_scenarios(
            demands,
            kerf_mm=kerf,
            min_scrap_mm=min_scrap,
            initial_scraps_mm=initial if initial else None,
            scenarios=scenarios,
        )
        mode = "waste_first"
        text = format_scenario_report(outcomes, mode=mode)
        extra = list(fw) + list(scrap_warns)
        if extra:
            text += "\n\n" + "\n".join(extra[:15])
        rec = pick_recommended(outcomes, mode=mode)
        self._recommended_bars = rec.bars_mm if rec else None
        if rec and rec.bars_mm:
            recommended_len = rec.bars_mm[0]
            self.base_bar_edit.setText(str(recommended_len))
            text += (
                "\n\n"
                f"Автоприменение: базовая длина установлена в {recommended_len} мм. "
                "Выполняю расчёт раскроя..."
            )
        self.advisor_text.setPlainText(text)
        logger.info(
            "Bar advisor completed: scenarios=%d recommended=%s mode=%s",
            len(outcomes),
            rec.name if rec else "none",
            mode,
        )
        if rec and rec.bars_mm:
            self._compute()

    def _apply_recommended_bars(self) -> None:
        if not self._recommended_bars:
            QMessageBox.information(
                self,
                "Автовыбор длин",
                "Сначала нажмите «Подбор заготовок», чтобы получить рекомендацию.",
            )
            return
        recommended_len = self._recommended_bars[0]
        self.base_bar_edit.setText(str(recommended_len))
        logger.info("Recommended bar applied manually: %d", recommended_len)
        self._compute()

    def _export_excel(self) -> None:
        if not self._last_sorted_cuts:
            QMessageBox.information(self, "Экспорт", "Сначала выполните расчёт.")
            return
        p, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить Excel",
            str(Path(self.path_edit.text()).with_name("raskroy_result.xlsx")),
            "Excel (*.xlsx)",
        )
        if not p:
            return
        try:
            logger.info("Export Excel start: %s rows=%d", p, len(self._last_sorted_cuts))
            export_cuts_excel(
                self._last_sorted_cuts,
                p,
                summary=self._last_summary_text,
            )
        except Exception as e:  # noqa: BLE001
            logger.exception("Export Excel failed")
            QMessageBox.critical(self, "Экспорт Excel", str(e))
            return
        logger.info("Export Excel done: %s", p)
        QMessageBox.information(self, "Экспорт", f"Сохранено:\n{p}")

    def _export_pdf(self) -> None:
        if not self._last_sorted_cuts:
            QMessageBox.information(self, "Экспорт", "Сначала выполните расчёт.")
            return
        p, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить PDF",
            str(Path(self.path_edit.text()).with_name("raskroy_result.pdf")),
            "PDF (*.pdf)",
        )
        if not p:
            return
        try:
            logger.info("Export PDF start: %s rows=%d", p, len(self._last_sorted_cuts))
            export_cuts_pdf(
                self._last_sorted_cuts,
                p,
                summary=self._last_summary_text,
            )
        except Exception as e:  # noqa: BLE001
            logger.exception("Export PDF failed")
            QMessageBox.critical(self, "Экспорт PDF", str(e))
            return
        logger.info("Export PDF done: %s", p)
        QMessageBox.information(self, "Экспорт", f"Сохранено:\n{p}")

    def _export_layout_excel(self) -> None:
        if not self._layout_rows:
            QMessageBox.information(self, "Экспорт схемы", "Сначала выполните расчёт.")
            return
        default = Path(self.path_edit.text() or "layout").with_name("raskroy_layout.xlsx")
        p, _ = QFileDialog.getSaveFileName(self, "Сохранить Excel-схему", str(default), "Excel (*.xlsx)")
        if not p:
            return
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font
        except ImportError as e:  # pragma: no cover
            QMessageBox.critical(self, "Экспорт схемы", f"Нужен openpyxl: {e}")
            return
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Схема раскроя"
        ws.append(
            (
                "Пруток №",
                "Длина заготовки",
                "Старт мм",
                "Длина сегмента",
                "Тип",
                "Подпись",
                "Происхождение",
            )
        )
        for c in range(1, 8):
            ws.cell(1, c).font = Font(bold=True)
        for row in self._layout_rows:
            opening = int(row["opening"])
            bar_len = int(row["bar_len"])
            segs = row["segments"]
            if not isinstance(segs, list):
                continue
            pos = 0
            for seg in segs:
                if not isinstance(seg, dict):
                    continue
                seg_len = int(seg.get("length_mm", 0))
                if seg_len <= 0:
                    continue
                ws.append(
                    (
                        opening,
                        bar_len,
                        pos,
                        seg_len,
                        str(seg.get("kind", "")),
                        str(seg.get("label", "")),
                        str(seg.get("origin", "")),
                    )
                )
                pos += seg_len
            rem = int(row.get("remainder", 0))
            ws.append((opening, bar_len, bar_len - rem, rem, "remainder", f"остаток {rem}", ""))
        wb.save(Path(p))
        QMessageBox.information(self, "Экспорт схемы", f"Сохранено:\n{p}")

    def _export_layout_pdf(self) -> None:
        if not self._layout_rows:
            QMessageBox.information(self, "Экспорт схемы", "Сначала выполните расчёт.")
            return
        default = Path(self.path_edit.text() or "layout").with_name("raskroy_layout.pdf")
        p, _ = QFileDialog.getSaveFileName(self, "Сохранить PDF-схему", str(default), "PDF (*.pdf)")
        if not p:
            return
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A3, A4, landscape
            from reportlab.lib.units import mm
            from reportlab.pdfgen import canvas
        except ImportError as e:  # pragma: no cover
            QMessageBox.critical(self, "Экспорт схемы", f"Нужен reportlab: {e}")
            return
        fmt, ok = QInputDialog.getItem(
            self,
            "Формат PDF",
            "Выберите формат:",
            ["A3 (горизонтально)", "A4 (горизонтально)"],
            0,
            False,
        )
        if not ok:
            return
        base_page = A3 if "A3" in fmt else A4
        font_regular, font_bold = _reportlab_cyrillic_fonts()
        c = canvas.Canvas(str(Path(p)), pagesize=landscape(base_page))
        pw, ph = landscape(base_page)
        y = ph - 18 * mm
        c.setFont(font_bold, 11)
        c.drawString(12 * mm, y, "Схема раскроя")
        y -= 10 * mm
        x0 = 40 * mm
        w = pw - 52 * mm
        bar_h = 6 * mm
        overflow_entries: list[dict[str, object]] = []
        overflow_idx = 0
        for row in self._layout_rows:
            opening = int(row["opening"])
            bar_len = int(row["bar_len"])
            rem = int(row.get("remainder", 0))
            segs = row.get("segments", [])
            if y < 20 * mm:
                c.showPage()
                y = ph - 18 * mm
                c.setFont(font_bold, 11)
                c.drawString(12 * mm, y, "Схема раскроя")
                y -= 10 * mm
            c.setFillColor(colors.black)
            c.setFont(font_regular, 7)
            c.drawString(12 * mm, y + 2 * mm, f"Пруток {opening} ({bar_len} мм)")
            oc = opening_row_color(opening)
            c.setFillColor(colors.Color(oc.red() / 255.0, oc.green() / 255.0, oc.blue() / 255.0))
            c.rect(9 * mm, y + 0.5 * mm, 1.8 * mm, bar_h - 1.0 * mm, stroke=0, fill=1)
            scrap_marks = {
                str(seg.get("origin", ""))
                for seg in segs
                if isinstance(seg, dict) and str(seg.get("kind", "")) == "profile" and bool(seg.get("is_scrap"))
            }
            if scrap_marks:
                c.setFillColor(colors.HexColor("#475569"))
                c.setFont(font_regular, 5)
                c.drawString(12 * mm, y - 1.2 * mm, "Обрезки: " + ", ".join(sorted(scrap_marks)))
            c.setStrokeColor(colors.HexColor("#94a3b8"))
            c.rect(x0, y, w, bar_h, stroke=1, fill=0)
            pos = 0.0
            profile_idx = 0
            if isinstance(segs, list):
                for seg in segs:
                    if not isinstance(seg, dict):
                        continue
                    seg_len = float(seg.get("length_mm", 0.0))
                    if seg_len <= 0:
                        continue
                    sx = x0 + w * (pos / bar_len)
                    sw = w * (seg_len / bar_len)
                    kind = str(seg.get("kind", ""))
                    if kind == "profile":
                        profile_name = str(seg.get("profile_name", ""))
                        qc = self._profile_color_map([profile_name]).get(profile_name, QColor(96, 165, 250))
                        c.setFillColor(colors.Color(qc.red() / 255.0, qc.green() / 255.0, qc.blue() / 255.0))
                        c.rect(sx, y, sw, bar_h, stroke=0, fill=1)
                    elif kind == "tech":
                        c.setFillColor(colors.HexColor("#ffedd5"))
                        c.setStrokeColor(colors.HexColor("#ea580c"))
                        c.setLineWidth(0.35)
                        c.rect(sx, y, sw, bar_h, stroke=1, fill=1)
                        if sw >= 1.5 * mm:
                            c.setFillColor(colors.HexColor("#c2410c"))
                            c.setFont(font_bold, 4)
                            c.drawCentredString(
                                sx + sw / 2.0,
                                y + bar_h + 2.2 * mm,
                                str(int(seg.get("tech_mm", 30))),
                            )
                    else:
                        c.setFillColor(colors.white)
                        c.rect(sx, y, sw, bar_h, stroke=0, fill=1)
                    if kind == "profile":
                        left_a = int(seg.get("left_angle", 90))
                        right_a = int(seg.get("right_angle", left_a))
                        draw_label = sw >= 14 * mm
                        draw_angles = sw >= 18 * mm
                        if draw_label:
                            txt = str(seg.get("label", ""))
                            c.setFillColor(colors.black)
                            c.setFont(font_regular, 5)
                            c.drawCentredString(sx + sw / 2.0, y + (bar_h / 2.0) - 2, txt)
                        if draw_angles:
                            c.setFillColor(colors.HexColor("#334155"))
                            c.setFont(font_regular, 4)
                            c.drawString(sx + 0.7 * mm, y + 0.5 * mm, f"L{left_a}°")
                            c.drawRightString(sx + sw - 0.7 * mm, y + 0.5 * mm, f"R{right_a}°")
                        # Если углы не влезли, добавляем деталь в легенду.
                        if not draw_angles:
                            overflow_idx += 1
                            mark = f"#{overflow_idx}"
                            c.setFillColor(colors.black)
                            c.setFont(font_bold, 5)
                            c.drawRightString(sx + sw - 0.4 * mm, y + (bar_h / 2.0) - 2, mark)
                            overflow_entries.append(
                                {
                                    "id": overflow_idx,
                                    "opening": opening,
                                    "module": str(seg.get("module_name", "")),
                                    "profile": str(seg.get("profile_code", "")),
                                    "length_mm": int(seg.get("part_length_mm", int(seg_len))),
                                    "left_angle": str(seg.get("left_angle", "")),
                                    "right_angle": str(seg.get("right_angle", "")),
                                    "source": str(seg.get("source_label", "")),
                                    "profile_idx": profile_idx,
                                }
                            )
                    if kind == "profile":
                        profile_idx += 1
                    if kind == "kerf" and sw >= 1.0 * mm:
                        c.setFillColor(colors.HexColor("#cffafe"))
                        c.setStrokeColor(colors.HexColor("#0e7490"))
                        c.setLineWidth(0.35)
                        c.rect(sx, y, sw, bar_h, stroke=1, fill=1)
                        c.setLineWidth(0.45)
                        c.line(sx, y, sx, y + bar_h)
                        c.line(sx + sw, y, sx + sw, y + bar_h)
                        if sw >= 2.0 * mm:
                            c.setFillColor(colors.HexColor("#155e75"))
                            c.setFont(font_regular, 4)
                            c.drawCentredString(
                                sx + sw / 2.0,
                                y + (bar_h / 2.0) - 1,
                                str(int(seg.get("kerf_mm", 4))),
                            )
                    pos += seg_len
            if rem > 0:
                rx = x0 + w * ((bar_len - rem) / bar_len)
                rw = w * (rem / bar_len)
                c.setStrokeColor(colors.HexColor("#64748b"))
                c.setDash(2, 2)
                c.rect(rx, y, rw, bar_h, stroke=1, fill=0)
                c.setDash()
                if rw >= 12 * mm:
                    c.setFillColor(colors.HexColor("#475569"))
                    c.setFont(font_regular, 5)
                    c.drawCentredString(rx + rw / 2.0, y + (bar_h / 2.0) - 2, f"остаток {rem}")
            y -= 11.5 * mm
        if overflow_entries:
            if y < 35 * mm:
                c.showPage()
                y = ph - 18 * mm
            c.setFillColor(colors.black)
            c.setFont(font_bold, 8)
            c.drawString(12 * mm, y, "Легенда коротких сегментов (#)")
            y -= 6 * mm
            c.setFont(font_bold, 6)
            c.drawString(12 * mm, y, "#")
            c.drawString(20 * mm, y, "Пруток")
            c.drawString(34 * mm, y, "Модуль")
            c.drawString(68 * mm, y, "Профиль")
            c.drawString(98 * mm, y, "Длина")
            c.drawString(112 * mm, y, "L")
            c.drawString(122 * mm, y, "R")
            c.drawString(136 * mm, y, "Источник")
            y -= 4 * mm
            c.setFont(font_regular, 6)
            for e in overflow_entries:
                if y < 14 * mm:
                    c.showPage()
                    y = ph - 18 * mm
                    c.setFont(font_bold, 8)
                    c.drawString(12 * mm, y, "Легенда коротких сегментов (#)")
                    y -= 6 * mm
                    c.setFont(font_bold, 6)
                    c.drawString(12 * mm, y, "#")
                    c.drawString(20 * mm, y, "Пруток")
                    c.drawString(34 * mm, y, "Модуль")
                    c.drawString(68 * mm, y, "Профиль")
                    c.drawString(98 * mm, y, "Длина")
                    c.drawString(112 * mm, y, "L")
                    c.drawString(122 * mm, y, "R")
                    c.drawString(136 * mm, y, "Источник")
                    y -= 4 * mm
                    c.setFont(font_regular, 6)
                oe = opening_row_color(int(e.get("opening", 0) or 0))
                c.setFillColor(colors.Color(oe.red() / 255.0, oe.green() / 255.0, oe.blue() / 255.0))
                c.rect(11.5 * mm, y - 1.2 * mm, (pw - 24 * mm), 4.0 * mm, stroke=0, fill=1)
                c.setStrokeColor(colors.HexColor("#cbd5e1"))
                c.rect(11.5 * mm, y - 1.2 * mm, (pw - 24 * mm), 4.0 * mm, stroke=1, fill=0)
                c.setFillColor(colors.black)
                c.drawString(12 * mm, y, str(e.get("id", "")))
                c.drawString(20 * mm, y, str(e.get("opening", "")))
                c.drawString(34 * mm, y, str(e.get("module", ""))[:22])
                c.drawString(68 * mm, y, str(e.get("profile", ""))[:12])
                c.drawString(98 * mm, y, str(e.get("length_mm", "")))
                c.drawString(112 * mm, y, str(e.get("left_angle", ""))[:5])
                c.drawString(122 * mm, y, str(e.get("right_angle", ""))[:5])
                c.drawString(136 * mm, y, str(e.get("source", ""))[:16])
                y -= 4 * mm
        c.save()
        QMessageBox.information(self, "Экспорт схемы", f"Сохранено:\n{p}")

    def _export_album_pdf(self) -> None:
        if not self._album_rows:
            QMessageBox.information(self, "Экспорт альбома", "Сначала выполните расчёт.")
            return
        default = Path(self.path_edit.text() or "layout").with_name("raskroy_album_a4.pdf")
        p, _ = QFileDialog.getSaveFileName(self, "Сохранить PDF-альбом", str(default), "PDF (*.pdf)")
        if not p:
            return
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm
            from reportlab.pdfgen import canvas
        except ImportError as e:  # pragma: no cover
            QMessageBox.critical(self, "Экспорт альбома", f"Нужен reportlab: {e}")
            return
        font_regular, font_bold = _reportlab_cyrillic_fonts()
        c = canvas.Canvas(str(Path(p)), pagesize=A4)
        pw, ph = A4
        y = ph - 16 * mm
        c.setFont(font_bold, 11)
        c.drawString(12 * mm, y, "Альбом стыков (A4, вертикально)")
        y -= 8 * mm
        left_x = 38 * mm
        detail_w = (pw - 2 * left_x) / 2
        detail_h = 18 * mm
        for r in self._album_rows:
            if y < 34 * mm:
                c.showPage()
                y = ph - 16 * mm
                c.setFont(font_bold, 11)
                c.drawString(12 * mm, y, "Альбом стыков (A4, вертикально)")
                y -= 8 * mm
            opening = int(r.get("opening", 0))
            kind = str(r.get("kind", "joint"))
            left_title = str(r.get("left_title", ""))
            right_title = str(r.get("right_title", ""))
            left_angle = int(r.get("left_right_angle", 90))
            right_angle = int(r.get("right_left_angle", 90))
            left_left_angle = int(r.get("left_left_angle", left_angle))
            right_right_angle = int(r.get("right_right_angle", right_angle))
            left_tech = int(r.get("left_tech_mm", 30))
            right_tech = int(r.get("right_tech_mm", 30))
            kerf_mm = int(r.get("kerf_mm", 4))
            c.setFont(font_bold, 8)
            c.setFillColor(colors.black)
            c.drawString(
                12 * mm,
                y + 2 * mm,
                f"Пруток {opening}: {'стык' if kind == 'joint' else 'деталь'}",
            )
            ly = y - detail_h
            seam_x = left_x + detail_w
            ltw = 4 * mm if left_tech == 50 else 2.5 * mm
            rtw = 4 * mm if right_tech == 50 else 2.5 * mm
            kw = 1.8 * mm if kerf_mm >= 4 else 1.2 * mm
            yb = ly + detail_h
            yt = ly

            def _pdf_strip_card(
                x0: float,
                tech_w: float,
                tech_mm_val: int,
                title: str,
                base_fill: colors.Color,
            ) -> None:
                ze = x0 + detail_w - kw
                c.setStrokeColor(colors.HexColor("#e2e8f0"))
                c.setFillColor(base_fill)
                c.rect(x0, ly, detail_w, detail_h, stroke=1, fill=1)
                c.setFillColor(colors.HexColor("#ffedd5"))
                c.setStrokeColor(colors.HexColor("#ea580c"))
                c.setLineWidth(0.3)
                c.rect(x0, ly, tech_w, detail_h, stroke=1, fill=1)
                c.setFillColor(colors.HexColor("#c2410c"))
                c.setFont(font_bold, 6)
                c.drawCentredString(x0 + tech_w / 2.0, yt + detail_h + 1.8 * mm, str(tech_mm_val))
                kx0 = x0 + tech_w
                c.setFillColor(colors.HexColor("#cffafe"))
                c.setStrokeColor(colors.HexColor("#0e7490"))
                c.rect(kx0, ly, kw, detail_h, stroke=1, fill=1)
                c.setLineWidth(0.35)
                c.line(kx0, yt, kx0, yb)
                c.line(kx0 + kw, yt, kx0 + kw, yb)
                c.setFillColor(colors.HexColor("#155e75"))
                c.setFont(font_regular, 6)
                c.drawCentredString(kx0 + kw / 2.0, ly + detail_h / 2.0 - 1.0, str(kerf_mm))
                body_x = x0 + tech_w + kw
                body_w = ze - body_x
                c.setFillColor(base_fill)
                c.setStrokeColor(base_fill)
                c.rect(body_x, ly, body_w, detail_h, stroke=0, fill=1)
                c.setStrokeColor(colors.HexColor("#0e7490"))
                c.setFillColor(colors.HexColor("#cffafe"))
                c.rect(ze, ly, kw, detail_h, stroke=1, fill=1)
                c.line(ze, yt, ze, yb)
                c.line(ze + kw, yt, ze + kw, yb)
                c.setFillColor(colors.HexColor("#155e75"))
                c.drawCentredString(ze + kw / 2.0, ly + detail_h / 2.0 - 1.0, str(kerf_mm))
                c.setFillColor(colors.black)
                c.setFont(font_regular, 7)
                c.drawCentredString(body_x + body_w / 2.0, ly + detail_h / 2.0 - 2, title[:42])

            _pdf_strip_card(
                left_x,
                ltw,
                left_tech,
                left_title,
                colors.HexColor("#dbeafe"),
            )
            _pdf_strip_card(
                seam_x,
                rtw,
                right_tech,
                right_title,
                colors.HexColor("#dcfce7"),
            )
            c.setStrokeColor(colors.HexColor("#0f172a"))
            c.setLineWidth(0.8)
            c.setDash(2, 2)
            c.line(seam_x, ly - 1.0 * mm, seam_x, ly + detail_h + 1.0 * mm)
            c.setDash()
            c.setFont(font_bold, 7)
            c.drawString(left_x + 1.0 * mm, ly - 4 * mm, f"L: {left_left_angle}°")
            c.drawString(left_x + detail_w - 16 * mm, ly - 4 * mm, f"R: {left_angle}°")
            c.drawString(seam_x + 1.0 * mm, ly - 4 * mm, f"L: {right_angle}°")
            c.drawString(seam_x + detail_w - 16 * mm, ly - 4 * mm, f"R: {right_right_angle}°")
            y = ly - 11 * mm
        c.save()
        QMessageBox.information(self, "Экспорт альбома", f"Сохранено:\n{p}")

    def _compute(self) -> None:
        logger.info("Compute requested")
        path = self.path_edit.text().strip()
        if not path or not Path(path).is_file():
            QMessageBox.critical(self, "Ошибка", "Укажите существующий файл .xlsx")
            return

        allowed: set[int] = {d for d, cb in self.profile_checks.items() if cb.isChecked()}
        if not allowed:
            QMessageBox.critical(self, "Ошибка", "Отметьте хотя бы один профиль Н20–Н23")
            return

        try:
            kerf = int(self.kerf_edit.text().strip() or "0")
        except ValueError:
            QMessageBox.critical(self, "Ошибка", "Пропил должен быть целым числом")
            return
        min_scrap = 0

        initial_scraps, scrap_warns = self._load_initial_scraps()
        try:
            rows_all, parse_stats = parse_specification_with_stats(path)
            rows, warns = filter_spec_by_profiles(rows_all, allowed)
            logger.info(
                "Specification loaded: total_rows=%d filtered_rows=%d filtered_out=%d",
                len(rows_all),
                len(rows),
                len(warns),
            )
            warn_text = ""
            if warns:
                warn_text = "\n\n".join(warns[:40])
                if len(warns) > 40:
                    warn_text += f"\n… и ещё {len(warns) - 40} сообщений"
            if not rows:
                QMessageBox.warning(
                    self,
                    "Нет данных",
                    "После фильтра по профилям не осталось строк.\n"
                    + (warn_text or "Проверьте галочки и формат кодов СК-/СС-/Р-."),
                )
                self._last_sorted_cuts = None
                self._optimizer_cuts = None
                self._last_summary_text = ""
                self.btn_xlsx.setEnabled(False)
                self.btn_pdf.setEnabled(False)
                self.btn_recalc.setEnabled(False)
                self.btn_layout_xlsx.setEnabled(False)
                self.btn_layout_pdf.setEnabled(False)
                self.mass_chart.clear_data()
                self.waste_chart.clear_data()
                self.layout_widget.clear_plan()
                self._layout_rows = []
                self.album_widget.clear_rows()
                self._album_rows = []
                return
            demands = spec_rows_to_demands(rows)
            bars, bars_err = self._selected_bar_lengths(
                demands,
                kerf,
            )
            if bars_err:
                QMessageBox.critical(self, "Ошибка", bars_err)
                return
            logger.info("Demands prepared: count=%d", len(demands))
            result = optimize_cutting(
                demands,
                bar_lengths_mm=bars,
                kerf_mm=kerf,
                min_scrap_mm=min_scrap,
                initial_scraps_mm=initial_scraps if initial_scraps else None,
            )
        except Exception as e:  # noqa: BLE001
            logger.exception("Compute failed")
            QMessageBox.critical(self, "Расчёт", str(e))
            return

        summary_lines = [summarize(result)]
        if self._selected_bars_mm:
            summary_lines.append(
                "Используемые длины заготовок, мм: "
                + ", ".join(str(x) for x in self._selected_bars_mm)
            )
        summary_lines.extend(self._mass_summary_lines(result.cuts))
        summary_lines.append(
            "Импорт спецификации: "
            f"прочитано {parse_stats.parsed_rows}, "
            f"пропущено пустых {parse_stats.skipped_blank_rows}, "
            f"некорректных {parse_stats.skipped_invalid_rows}"
        )
        if warns:
            summary_lines.append(f"Исключено при фильтре: {len(warns)} строк(и)")
        if scrap_warns:
            summary_lines.append("Склад обрезков: " + "; ".join(scrap_warns[:5]))
        full_summary = "\n".join(summary_lines) + (f"\n\n{warn_text}" if warn_text else "")
        chart_m = self._compute_chart_metrics(result)
        if chart_m is not None:
            full_summary += "\n\n" + "\n".join(self._chart_summary_lines(chart_m))
        self.advisor_text.setPlainText(full_summary)
        self._last_summary_text = full_summary

        self._apply_optimization_result(
            result=result,
            kerf_mm=kerf,
            chart_metrics=chart_m,
        )
        logger.info(
            "Compute completed: cuts=%d new_bars=%d",
            len(result.cuts),
            sum(result.bars_used.values()),
        )

        if warns:
            QMessageBox.information(
                self,
                "Фильтр профилей",
                f"Исключено строк: {len(warns)}. Подробности — в сводке ниже.",
            )

    def _apply_sort(self) -> None:
        if not self._optimizer_cuts:
            return
        mode = self.sort_combo.currentData()
        if not isinstance(mode, str):
            mode = "opening"
        cuts = sort_cuts(self._optimizer_cuts, mode)
        self._last_sorted_cuts = cuts
        self._populate_table(cuts)

    def _populate_table(self, cuts: list[CutEvent]) -> None:
        logger.info("populate_table start: cuts=%d", len(cuts))
        self._footer_row_index = None
        self._data_row_count = len(cuts)
        self.table.setRowCount(0)
        total_kg = 0.0
        mass_by_profile: dict[str, float] = defaultdict(float)
        any_mass = False
        for cut in cuts:
            d = cut.demand
            row = self.table.rowCount()
            self.table.insertRow(row)
            series = profile_label_for_code(d.profile_code)
            kg, mtxt = row_mass_kg_display(d.profile_code, float(d.length_mm), 1.0)
            if kg is not None:
                total_kg += kg
                prof_name = profile_label_for_code(d.profile_code)
                if prof_name == "—":
                    prof_name = d.profile_code
                mass_by_profile[prof_name] += kg
                any_mass = True
            op_txt = (
                "склад" if cut.stock_opening_id == 0 else str(cut.stock_opening_id)
            )
            vals = [
                op_txt,
                d.module_name,
                d.profile_code,
                series,
                str(d.length_mm),
                format_cut_angles(d),
                "Обрезок" if cut.stock_source == "scrap" else "Новая",
                str(cut.stock_length_mm),
                str(cut.remainder_mm),
                mtxt,
            ]
            rgb = module_row_rgb(
                d.module_name,
                is_scrap=cut.stock_source == "scrap",
            )
            bg = QColor(*rgb)
            for col, text in enumerate(vals):
                it = QTableWidgetItem(text)
                flags = Qt.ItemIsSelectable | Qt.ItemIsEnabled
                if col != 9:
                    flags |= Qt.ItemIsEditable
                it.setFlags(flags)
                it.setBackground(bg)
                self.table.setItem(row, col, it)
            logger.info(
                "table row added: module=%s profile=%s len=%d source=%s mass=%s",
                d.module_name,
                d.profile_code,
                d.length_mm,
                cut.stock_source,
                mtxt,
            )

        if cuts:
            # Итог первой таблицы должен стоять сразу после строк раскроя.
            fr = self.table.rowCount()
            self.table.insertRow(fr)
            self._footer_row_index = fr
            foot_bg = QColor(226, 232, 240)
            lab = QTableWidgetItem("Итого масса профиля, кг")
            lab.setFlags(Qt.ItemIsEnabled)
            lab.setBackground(foot_bg)
            lab.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
            self.table.setItem(fr, 0, lab)
            self.table.setSpan(fr, 0, 1, 9)
            tot_txt = (
                f"{total_kg:.3f}".rstrip("0").rstrip(".")
                if any_mass
                else "—"
            )
            tot_it = QTableWidgetItem(tot_txt if tot_txt else "0")
            tot_it.setFlags(Qt.ItemIsEnabled)
            tot_it.setBackground(foot_bg)
            self.table.setItem(fr, 9, tot_it)
            logger.info("table total mass row: total_kg=%s", tot_txt if tot_txt else "0")

            # Дополнительные строки по массе каждого профиля в первой таблице.
            std_names = [PROFILE_DIGIT_TO_NAME[d] for d in sorted(PROFILE_DIGIT_TO_NAME)]
            ordered_profiles = [n for n in std_names if n in mass_by_profile] + [
                n for n in sorted(mass_by_profile) if n not in std_names
            ]
            for name in ordered_profiles:
                r = self.table.rowCount()
                self.table.insertRow(r)
                row_bg = QColor(241, 245, 249)
                lbl = QTableWidgetItem(f"Масса профиля {name}, кг")
                lbl.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                lbl.setBackground(row_bg)
                lbl.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
                self.table.setItem(r, 0, lbl)
                self.table.setSpan(r, 0, 1, 9)
                val = f"{mass_by_profile[name]:.3f}".rstrip("0").rstrip(".")
                vit = QTableWidgetItem(val if val else "0")
                vit.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                vit.setBackground(row_bg)
                self.table.setItem(r, 9, vit)
                logger.info("table mass-by-profile row: %s=%s", name, val if val else "0")

        # Вторая таблица: закупка профилей, отделена от первой.
        purchase_rows: dict[tuple[str, int], tuple[int, float | None]] = {}
        for cut in cuts:
            if cut.stock_source != "new_bar":
                continue
            code = cut.demand.profile_code
            profile_type = profile_label_for_code(code)
            if profile_type == "—":
                profile_type = code
            bar_len = int(cut.stock_length_mm)
            key = (profile_type, bar_len)
            old = purchase_rows.get(key, (0, 0.0))
            count = old[0] + 1
            lookup = kg_per_meter_from_profile_code(code)
            if lookup.source == "unknown" or lookup.kg_per_m <= 0:
                mass = None
            else:
                add_kg = total_mass_kg(float(bar_len), lookup.kg_per_m, 1.0)
                mass = (old[1] or 0.0) + add_kg
            purchase_rows[key] = (count, mass)
            logger.info(
                "purchase aggregate step: profile=%s bar=%d count=%d mass=%s",
                profile_type,
                bar_len,
                count,
                "—" if mass is None else f"{mass:.3f}",
            )

        if purchase_rows:
            sep = self.table.rowCount()
            self.table.insertRow(sep)
            sep_bg = QColor(255, 255, 255)
            sep_item = QTableWidgetItem("")
            sep_item.setFlags(Qt.ItemIsEnabled)
            sep_item.setBackground(sep_bg)
            self.table.setItem(sep, 0, sep_item)
            self.table.setSpan(sep, 0, 1, 10)

            title_row = self.table.rowCount()
            self.table.insertRow(title_row)
            head_bg = QColor(219, 234, 254)
            t = QTableWidgetItem("Таблица закупки профилей")
            t.setFlags(Qt.ItemIsEnabled)
            t.setBackground(head_bg)
            t.setTextAlignment(int(Qt.AlignLeft | Qt.AlignVCenter))
            self.table.setItem(title_row, 0, t)
            self.table.setSpan(title_row, 0, 1, 10)

            # Шапка второй таблицы (4 колонки через объединения, без пустых дублей).
            hdr_row = self.table.rowCount()
            self.table.insertRow(hdr_row)
            h_name = QTableWidgetItem("Наименование")
            h_len = QTableWidgetItem("Длина заготовки, мм")
            h_qty = QTableWidgetItem("Количество, шт")
            h_mass = QTableWidgetItem("Масса, кг")
            for c, it in ((0, h_name), (3, h_len), (5, h_qty), (7, h_mass)):
                it.setFlags(Qt.ItemIsEnabled)
                it.setBackground(head_bg)
                it.setTextAlignment(int(Qt.AlignCenter))
                self.table.setItem(hdr_row, c, it)
            self.table.setSpan(hdr_row, 0, 1, 3)
            self.table.setSpan(hdr_row, 3, 1, 2)
            self.table.setSpan(hdr_row, 5, 1, 2)
            self.table.setSpan(hdr_row, 7, 1, 2)

            purchase_total_kg = 0.0
            purchase_any_mass = False
            for (profile_type, bar_len), (count, mass_kg) in sorted(
                purchase_rows.items(),
                key=lambda x: (x[0][0], x[0][1]),
            ):
                r = self.table.rowCount()
                self.table.insertRow(r)
                mass_txt = "—" if mass_kg is None else f"{mass_kg:.3f}".rstrip("0").rstrip(".")
                row_bg = QColor(239, 246, 255)
                v_name = QTableWidgetItem(profile_type)
                v_len = QTableWidgetItem(str(bar_len))
                v_qty = QTableWidgetItem(str(count))
                v_mass = QTableWidgetItem(mass_txt if mass_txt else "0")
                for c, it in ((0, v_name), (3, v_len), (5, v_qty), (7, v_mass)):
                    it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    it.setBackground(row_bg)
                    it.setTextAlignment(int(Qt.AlignCenter))
                    self.table.setItem(r, c, it)
                self.table.setSpan(r, 0, 1, 3)
                self.table.setSpan(r, 3, 1, 2)
                self.table.setSpan(r, 5, 1, 2)
                self.table.setSpan(r, 7, 1, 2)
                if mass_kg is not None:
                    purchase_total_kg += mass_kg
                    purchase_any_mass = True
                logger.info(
                    "purchase row added: profile=%s bar=%d qty=%d mass=%s",
                    profile_type,
                    bar_len,
                    count,
                    mass_txt if mass_txt else "0",
                )

            # Итог по второй таблице одной строкой (количество + масса).
            qty_total = sum(cnt for (cnt, _m) in purchase_rows.values())
            r_tot2 = self.table.rowCount()
            self.table.insertRow(r_tot2)
            foot2_bg = QColor(219, 234, 254)
            t2 = QTableWidgetItem("Итого")
            t2.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            t2.setBackground(foot2_bg)
            t2.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
            self.table.setItem(r_tot2, 0, t2)
            self.table.setSpan(r_tot2, 0, 1, 5)
            qit = QTableWidgetItem(str(qty_total))
            qit.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            qit.setBackground(foot2_bg)
            qit.setTextAlignment(int(Qt.AlignCenter))
            self.table.setItem(r_tot2, 5, qit)
            self.table.setSpan(r_tot2, 5, 1, 2)
            t2_val = (
                f"{purchase_total_kg:.3f}".rstrip("0").rstrip(".")
                if purchase_any_mass
                else "—"
            )
            t2i = QTableWidgetItem(t2_val if t2_val else "0")
            t2i.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            t2i.setBackground(foot2_bg)
            t2i.setTextAlignment(int(Qt.AlignCenter))
            self.table.setItem(r_tot2, 7, t2i)
            self.table.setSpan(r_tot2, 7, 1, 2)
            logger.info(
                "purchase total row: qty=%d mass=%s",
                qty_total,
                t2_val if t2_val else "0",
            )
        # Ручной ресайз колонок всегда активен.
        hdr = self.table.horizontalHeader()
        for i in range(self.table.columnCount()):
            hdr.setSectionResizeMode(i, QHeaderView.Interactive)
        self._set_table_zoom(self._table_zoom)
        logger.info("populate_table done: total_rows=%d", self.table.rowCount())

    def _data_rows_for_recalc(self) -> list[list[str]]:
        """Колонки «Модуль»…«Остаток» (без «Пруток №» и массы), без строки «Итого»."""
        end = self._data_row_count
        rows: list[list[str]] = []
        for r in range(end):
            row: list[str] = []
            for c in range(1, 9):
                it = self.table.item(r, c)
                row.append(it.text().strip() if it is not None else "")
            rows.append(row)
        return rows

    def _copy_table_selection(self) -> bool:
        """Копировать выделенный фрагмент таблицы (TSV) в буфер обмена."""
        ranges = self.table.selectedRanges()
        if not ranges:
            sel_model = self.table.selectionModel()
            rows = sel_model.selectedRows() if sel_model is not None else []
            if not rows:
                logger.info("Copy table skipped: no selection")
                return False
            min_row = min(i.row() for i in rows)
            max_row = max(i.row() for i in rows)
            min_col = 0
            max_col = self.table.columnCount() - 1
            selected: set[tuple[int, int]] = {
                (r.row(), c) for r in rows for c in range(self.table.columnCount())
            }
        else:
            min_row = min(r.topRow() for r in ranges)
            max_row = max(r.bottomRow() for r in ranges)
            min_col = min(r.leftColumn() for r in ranges)
            max_col = max(r.rightColumn() for r in ranges)
            selected = set()
            for rg in ranges:
                for rr in range(rg.topRow(), rg.bottomRow() + 1):
                    for cc in range(rg.leftColumn(), rg.rightColumn() + 1):
                        selected.add((rr, cc))

        lines: list[str] = []
        for r in range(min_row, max_row + 1):
            vals: list[str] = []
            for c in range(min_col, max_col + 1):
                if (r, c) in selected:
                    it = self.table.item(r, c)
                    vals.append(it.text() if it is not None else "")
                else:
                    vals.append("")
            lines.append("\t".join(vals))

        QApplication.clipboard().setText("\n".join(lines))
        logger.info(
            "Copied table selection to clipboard: rows=%d..%d cols=%d..%d",
            min_row,
            max_row,
            min_col,
            max_col,
        )
        return True

    def _handle_copy_shortcut(self) -> bool:
        """Единый обработчик Ctrl+C для таблицы и текстовых полей."""
        fw = QApplication.focusWidget()
        if fw is None:
            return False
        if fw is self.table or self.table.isAncestorOf(fw):
            return self._copy_table_selection()
        te = self._ancestor_text_edit(fw)
        if te is not None:
            cur = te.textCursor()
            if cur.hasSelection():
                te.copy()
                logger.info("Copied from %s", type(te).__name__)
                return True
            return False
        if isinstance(fw, QLineEdit):
            sel = fw.selectedText()
            if sel:
                QApplication.clipboard().setText(sel)
                logger.info("Copied QLineEdit selection")
                return True
            return False
        if hasattr(fw, "copy"):
            try:
                fw.copy()  # type: ignore[attr-defined]
                logger.info("Copied from widget: %s", type(fw).__name__)
                return True
            except Exception:
                logger.exception("Copy from focused widget failed")
        return False

    def _copy_from_focused_widget(self) -> None:
        """Глобальный Ctrl+C: таблица или любой текстовый виджет."""
        self._handle_copy_shortcut()

    def _show_table_context_menu(self, pos) -> None:
        menu = QMenu(self.table)
        act_copy = QAction("Копировать", self.table)
        act_copy.triggered.connect(self._copy_table_selection)
        menu.addAction(act_copy)
        menu.exec(self.table.viewport().mapToGlobal(pos))

    def _recalc_from_table(self) -> None:
        logger.info("Recalc from table requested")
        if self.table.rowCount() == 0:
            QMessageBox.information(self, "Пересчёт", "Таблица пуста.")
            return

        allowed: set[int] = {d for d, cb in self.profile_checks.items() if cb.isChecked()}
        if not allowed:
            QMessageBox.critical(self, "Ошибка", "Отметьте хотя бы один профиль Н20–Н23")
            return

        try:
            kerf = int(self.kerf_edit.text().strip() or "0")
        except ValueError:
            QMessageBox.critical(self, "Ошибка", "Пропил должен быть целым числом")
            return
        min_scrap = 0

        matrix = self._data_rows_for_recalc()
        demands, err = demands_from_cut_table_rows(matrix, allowed)
        if err or demands is None:
            QMessageBox.critical(self, "Таблица", err or "Ошибка разбора таблицы")
            return
        initial_scraps, _sw = self._load_initial_scraps()
        bars, bars_err = self._selected_bar_lengths(
            demands,
            kerf,
        )
        if bars_err:
            QMessageBox.critical(self, "Ошибка", bars_err)
            return
        logger.info("Recalc demands parsed: count=%d", len(demands))
        try:
            result = optimize_cutting(
                demands,
                bar_lengths_mm=bars,
                kerf_mm=kerf,
                min_scrap_mm=min_scrap,
                initial_scraps_mm=initial_scraps if initial_scraps else None,
            )
        except Exception as e:  # noqa: BLE001
            logger.exception("Recalc failed")
            QMessageBox.critical(self, "Пересчёт", str(e))
            return

        summary_lines = [summarize(result)]
        summary_lines.extend(self._mass_summary_lines(result.cuts))
        note = (
            "\n\n(Пересчёт по отредактированной таблице; "
            "колонки «Серия», «Источник», «Заготовка», «Остаток» обновлены.)"
        )
        full = "\n".join(summary_lines) + note
        chart_m = self._compute_chart_metrics(result)
        if chart_m is not None:
            full += "\n\n" + "\n".join(self._chart_summary_lines(chart_m))
        self.advisor_text.setPlainText(full)
        self._last_summary_text = full
        self._apply_optimization_result(
            result=result,
            kerf_mm=kerf,
            chart_metrics=chart_m,
        )
        logger.info(
            "Recalc completed: cuts=%d new_bars=%d",
            len(result.cuts),
            sum(result.bars_used.values()),
        )

    def _mass_summary_lines(self, cuts: list[CutEvent]) -> list[str]:
        """
        Возвращает строки сводки с общей массой и массой по каждому профилю/серии.
        """
        total_kg = 0.0
        by_profile_kg: dict[str, float] = defaultdict(float)
        unknown = 0

        for cut in cuts:
            d = cut.demand
            kg, _txt = row_mass_kg_display(d.profile_code, float(d.length_mm), 1.0)
            if kg is None:
                unknown += 1
                continue
            total_kg += kg
            label = profile_label_for_code(d.profile_code)
            if label == "—":
                label = d.profile_code
            by_profile_kg[label] += kg

        lines: list[str] = []
        if by_profile_kg:
            lines.append(f"Итоговая масса профиля, кг: {total_kg:.3f}".rstrip("0").rstrip("."))
            lines.append("Масса по профилям, кг:")
            # Сначала стандартные серии Н20..Н23, затем остальные.
            std_names = [PROFILE_DIGIT_TO_NAME[d] for d in sorted(PROFILE_DIGIT_TO_NAME)]
            for name in std_names:
                if name in by_profile_kg:
                    txt = f"{by_profile_kg[name]:.3f}".rstrip("0").rstrip(".")
                    lines.append(f"  {name}: {txt}")
            for name in sorted(by_profile_kg):
                if name in std_names:
                    continue
                txt = f"{by_profile_kg[name]:.3f}".rstrip("0").rstrip(".")
                lines.append(f"  {name}: {txt}")
        else:
            lines.append("Итоговая масса профиля, кг: —")
        if unknown > 0:
            lines.append(f"Масса не определена для {unknown} строк(и) (нет справочного кг/м).")
        return lines

    def _profile_color_map(self, names: list[str]) -> dict[str, QColor]:
        """
        Стабильные цвета профилей для обеих диаграмм.
        """
        base: dict[str, QColor] = {
            "Н20": QColor(59, 130, 246),
            "Н21": QColor(16, 185, 129),
            "Н22": QColor(245, 158, 11),
            "Н23": QColor(168, 85, 247),
        }
        palette = [
            QColor(236, 72, 153),
            QColor(14, 165, 233),
            QColor(34, 197, 94),
            QColor(251, 146, 60),
            QColor(244, 63, 94),
        ]
        for i, n in enumerate(sorted(names)):
            if n not in base:
                base[n] = palette[i % len(palette)]
        return base

    def _module_order_key(self, name: str) -> tuple[int, str]:
        m = re.search(r"[MМ]\s*(\d+)", name, flags=re.IGNORECASE)
        if m:
            return (int(m.group(1)), name)
        return (10**9, name)

    def _module_short_name(self, name: str) -> str:
        m = re.search(r"[MМ]\s*(\d+)", name, flags=re.IGNORECASE)
        if m:
            return f"М{int(m.group(1))}"
        return name

    def _update_layout_plan(self, cuts: list[CutEvent], kerf_mm: int) -> None:
        """Готовит данные для вкладки схемы раскроя (прямоугольники-прутки)."""
        by_opening: dict[int, list[CutEvent]] = defaultdict(list)
        for c in cuts:
            if c.stock_opening_id <= 0:
                continue
            by_opening[c.stock_opening_id].append(c)
        if not by_opening:
            self._layout_rows = []
            self.layout_widget.clear_plan()
            return

        rows: list[dict[str, object]] = []
        profile_names: set[str] = set()
        for opening in sorted(by_opening):
            group = by_opening[opening]
            first_new = next((c for c in group if c.stock_source == "new_bar"), None)
            if first_new is None:
                continue
            bar_len = int(first_new.stock_length_mm)
            segs: list[dict[str, object]] = []
            consumed = 0
            for c in group:
                d = c.demand
                profile_name = profile_label_for_code(d.profile_code)
                if profile_name == "—":
                    profile_name = d.profile_code
                profile_names.add(profile_name)
                cut_angle = int(d.cut_angle)
                right_angle = int(d.cut_angle_2) if d.cut_angle_2 is not None else int(d.cut_angle)
                tech = 30 if cut_angle == 90 else 50
                # Перед каждой деталью: левый тех. отступ + левый пропил (суммарно).
                left_kerf = int(kerf_mm)
                segs.append(
                    {
                        "kind": "tech",
                        "length_mm": float(tech),
                        "label": f"тех. {tech} мм",
                        "tech_mm": tech,
                    }
                )
                consumed += tech
                if left_kerf > 0:
                    segs.append(
                        {
                            "kind": "kerf",
                            "length_mm": float(left_kerf),
                            "label": "пропил",
                            "side": "leading",
                            "angle": cut_angle,
                            "kerf_mm": int(kerf_mm),
                        }
                    )
                    consumed += left_kerf
                segs.append(
                    {
                        "kind": "profile",
                        "length_mm": int(d.length_mm),
                        "label": f"{self._module_short_name(d.module_name)} {d.profile_code}",
                        "module_name": d.module_name,
                        "profile_code": d.profile_code,
                        "part_length_mm": int(d.length_mm),
                        "cut_angles": format_cut_angles(d),
                        "left_angle": cut_angle,
                        "right_angle": right_angle,
                        "source_label": "Обрезок" if c.stock_source == "scrap" else "Новая",
                        "opening_color": (
                            opening_row_color(opening).red(),
                            opening_row_color(opening).green(),
                            opening_row_color(opening).blue(),
                        ),
                        "profile_name": profile_name,
                        "is_scrap": c.stock_source == "scrap",
                        "origin": (
                            "из склада"
                            if c.stock_source == "scrap" and c.stock_opening_id == 0
                            else f"из прутка {c.stock_opening_id}"
                            if c.stock_source == "scrap"
                            else f"новый пруток {opening}"
                        ),
                    }
                )
                consumed += int(d.length_mm)
                right_kerf = int(kerf_mm)
                if right_kerf > 0:
                    segs.append(
                        {
                            "kind": "kerf",
                            "length_mm": right_kerf,
                            "label": "пропил",
                            "side": "trailing",
                            "angle": right_angle,
                            "kerf_mm": int(kerf_mm),
                        }
                    )
                    consumed += right_kerf
            remainder = max(bar_len - consumed, 0)
            rows.append(
                {
                    "opening": opening,
                    "bar_len": bar_len,
                    "segments": segs,
                    "remainder": remainder,
                }
            )
        color_map = self._profile_color_map(list(profile_names))
        self._layout_rows = rows
        self.layout_widget.set_plan(rows, color_map)

    def _update_album_plan(self, cuts: list[CutEvent]) -> None:
        by_opening: dict[int, list[CutEvent]] = defaultdict(list)
        for c in cuts:
            if c.stock_opening_id <= 0:
                continue
            by_opening[c.stock_opening_id].append(c)
        rows: list[dict[str, object]] = []
        profile_names: set[str] = set()
        mode = self.album_mode_combo.currentData() if hasattr(self, "album_mode_combo") else "joints"
        for opening in sorted(by_opening):
            g = by_opening[opening]
            if mode == "details":
                for c in g:
                    d = c.demand
                    prof = profile_label_for_code(d.profile_code)
                    if prof == "—":
                        prof = d.profile_code
                    profile_names.add(prof)
                    la = int(d.cut_angle)
                    ra = int(d.cut_angle_2) if d.cut_angle_2 is not None else int(d.cut_angle)
                    rows.append(
                        {
                            "kind": "detail",
                            "opening": opening,
                            "left_title": f"{self._module_short_name(d.module_name)} {d.profile_code} ({d.length_mm})",
                            "right_title": f"{self._module_short_name(d.module_name)} {d.profile_code} ({d.length_mm})",
                            "left_profile_name": prof,
                            "right_profile_name": prof,
                            "left_right_angle": la,
                            "right_left_angle": ra,
                            "left_left_angle": la,
                            "right_right_angle": ra,
                            "left_tech_mm": 30 if la == 90 else 50,
                            "right_tech_mm": 30 if ra == 90 else 50,
                            "kerf_mm": self._last_kerf_mm or 4,
                        }
                    )
            else:
                for i in range(len(g) - 1):
                    left = g[i].demand
                    right = g[i + 1].demand
                    left_prof = profile_label_for_code(left.profile_code)
                    right_prof = profile_label_for_code(right.profile_code)
                    if left_prof == "—":
                        left_prof = left.profile_code
                    if right_prof == "—":
                        right_prof = right.profile_code
                    profile_names.add(left_prof)
                    profile_names.add(right_prof)
                    rows.append(
                        {
                            "kind": "joint",
                            "opening": opening,
                            "left_title": f"{self._module_short_name(left.module_name)} {left.profile_code} ({left.length_mm})",
                            "right_title": f"{self._module_short_name(right.module_name)} {right.profile_code} ({right.length_mm})",
                            "left_profile_name": left_prof,
                            "right_profile_name": right_prof,
                            "left_right_angle": int(left.cut_angle_2) if left.cut_angle_2 is not None else int(left.cut_angle),
                            "right_left_angle": int(right.cut_angle),
                            "left_left_angle": int(left.cut_angle),
                            "right_right_angle": int(right.cut_angle_2) if right.cut_angle_2 is not None else int(right.cut_angle),
                            "left_tech_mm": 30 if int(left.cut_angle) == 90 else 50,
                            "right_tech_mm": 30 if int(right.cut_angle) == 90 else 50,
                            "kerf_mm": self._last_kerf_mm or 4,
                        }
                    )
        self._album_rows = rows
        self.album_widget.set_rows(rows, self._profile_color_map(list(profile_names)))

    def _ordered_profile_names(self, names: set[str]) -> list[str]:
        """Порядок подписей как в сводке массы: Н20…Н23, затем остальные по алфавиту."""
        out: list[str] = []
        std_names = [PROFILE_DIGIT_TO_NAME[d] for d in sorted(PROFILE_DIGIT_TO_NAME)]
        for name in std_names:
            if name in names:
                out.append(name)
        for name in sorted(names):
            if name in std_names:
                continue
            out.append(name)
        return out

    def _compute_chart_metrics(self, result: OptimizationResult) -> dict[str, object] | None:
        """
        Единый расчёт для кольцевых диаграмм и текстового дубля в поле сводки.
        Диаграмма 1: масса по профилям (кг), в центре суммарная масса.
        Диаграмма 2: отходы по профилям (кг), в центре суммарный вес отходов.
        """
        purchased_mm = sum(L * c for L, c in result.bars_used.items())
        if purchased_mm <= 0:
            logger.info("chart metrics: purchased_mm=0")
            return None

        used_total_mm = sum(c.stock_length_mm - c.remainder_mm for c in result.cuts)
        kpd_pct = (used_total_mm / purchased_mm) * 100.0

        profile_mass_kg: dict[str, float] = defaultdict(float)
        profile_purchased_kg: dict[str, float] = defaultdict(float)
        total_mass_kg_all = 0.0
        seen_new_openings: set[int] = set()

        for c in result.cuts:
            d = c.demand
            label = profile_label_for_code(d.profile_code)
            if label == "—":
                label = d.profile_code

            used_kg, _ = row_mass_kg_display(d.profile_code, float(d.length_mm), 1.0)
            if used_kg is not None:
                profile_mass_kg[label] += used_kg
                total_mass_kg_all += used_kg

            if c.stock_source == "new_bar" and c.stock_opening_id not in seen_new_openings:
                seen_new_openings.add(c.stock_opening_id)
                lookup = kg_per_meter_from_profile_code(d.profile_code)
                if lookup.source != "unknown" and lookup.kg_per_m > 0:
                    purchased_kg = total_mass_kg(float(c.stock_length_mm), lookup.kg_per_m, 1.0)
                    profile_purchased_kg[label] += purchased_kg

        profile_waste_kg: dict[str, float] = {}
        for name in set(profile_purchased_kg) | set(profile_mass_kg):
            purchased = max(profile_purchased_kg.get(name, 0.0), 0.0)
            used = max(profile_mass_kg.get(name, 0.0), 0.0)
            # Обрезки не могут превышать массу закупки по профилю.
            waste_kg = min(max(purchased - used, 0.0), purchased)
            if waste_kg > 0:
                profile_waste_kg[name] = waste_kg

        purchased_total_kg = sum(max(v, 0.0) for v in profile_purchased_kg.values())
        waste_total = min(sum(profile_waste_kg.values()), purchased_total_kg)

        return {
            "purchased_mm": purchased_mm,
            "used_total_mm": used_total_mm,
            "kpd_pct": kpd_pct,
            "profile_mass_kg": dict(profile_mass_kg),
            "profile_waste_kg": profile_waste_kg,
            "total_mass_kg_all": total_mass_kg_all,
            "purchased_total_kg_all": purchased_total_kg,
            "waste_total_kg": waste_total,
        }

    def _chart_summary_lines(self, m: dict[str, object]) -> list[str]:
        """Словесное описание тех же данных, что на кольцевых диаграммах слева."""
        kpd_pct = float(m["kpd_pct"])
        total_mass = float(m["total_mass_kg_all"])
        purchased_total = float(m.get("purchased_total_kg_all", 0.0))
        waste_total = float(m["waste_total_kg"])
        profile_mass_kg = m["profile_mass_kg"]
        profile_waste_kg = m["profile_waste_kg"]
        if not isinstance(profile_mass_kg, dict) or not isinstance(profile_waste_kg, dict):
            return ["— Кольцевые диаграммы: ошибка формата данных —"]

        lines: list[str] = [
            "— Кольцевые диаграммы (текст, как на схеме слева) —",
            "",
            "«Масса профилей» (первая диаграмма)",
            f"  Подзаголовок: КПД раскроя — {kpd_pct:.1f}% (доля использованной длины от суммы закупленных длин новых прутков).",
            f"  В центре кольца: суммарная масса деталей — {_fmt_kg_trim(total_mass)} кг.",
            "  Сегменты — масса деталей по профилям, кг:",
        ]
        mass_positive = {k: v for k, v in profile_mass_kg.items() if v > 0}
        for name in self._ordered_profile_names(set(mass_positive.keys())):
            kg = mass_positive[name]
            pct = (kg / total_mass * 100.0) if total_mass > 0 else 0.0
            lines.append(
                f"    {name}: {_fmt_kg_trim(kg)} кг ({pct:.1f}% от суммарной массы деталей)"
            )
        if not mass_positive:
            lines.append("    (нет положительных масс по профилям)")

        lines.extend(
            [
                "",
                "«Отходы по профилям» (вторая диаграмма)",
                "  Подзаголовок: сегменты по профилям, кг (для новых прутков: масса закупленного прутка минус масса деталей того же профиля).",
                f"  В центре кольца: суммарный вес отходов — {_fmt_kg_trim(waste_total)} кг.",
                "  Сегменты — отходы по профилям, кг:",
            ]
        )
        waste_pos = {k: v for k, v in profile_waste_kg.items() if v > 0}
        if not waste_pos or waste_total <= 0:
            lines.append("    (отходов нет или доли не выделены)")
        else:
            for name in self._ordered_profile_names(set(waste_pos.keys())):
                kg = waste_pos[name]
                pct = (kg / waste_total * 100.0) if waste_total > 0 else 0.0
                lines.append(
                    f"    {name}: {_fmt_kg_trim(kg)} кг ({pct:.1f}% от суммарных отходов)"
                )
        lines.extend(
            [
                "",
                "Контроль баланса масс:",
                f"  Закупка (новые прутки): {_fmt_kg_trim(purchased_total)} кг",
                f"  Детали: {_fmt_kg_trim(total_mass)} кг",
                f"  Обрезки: {_fmt_kg_trim(waste_total)} кг",
                (
                    "  Проверка: детали + обрезки <= закупка — OK"
                    if (total_mass + waste_total) <= (purchased_total + 1e-6)
                    else "  Проверка: детали + обрезки > закупка (проверьте исходные данные)"
                ),
            ]
        )
        return lines

    def _apply_chart_metrics(self, m: dict[str, object] | None) -> None:
        """Обновляет виджеты диаграмм по уже посчитанным метрикам."""
        if m is None:
            self.mass_chart.clear_data()
            self.waste_chart.clear_data()
            logger.info("charts cleared: no chart metrics")
            return

        kpd_pct = float(m["kpd_pct"])
        total_mass_kg_all = float(m["total_mass_kg_all"])
        profile_mass_kg = m["profile_mass_kg"]
        profile_waste_kg = m["profile_waste_kg"]
        waste_total = float(m["waste_total_kg"])
        if not isinstance(profile_mass_kg, dict) or not isinstance(profile_waste_kg, dict):
            self.mass_chart.clear_data()
            self.waste_chart.clear_data()
            return

        color_map = self._profile_color_map(
            list(set(profile_mass_kg.keys()) | set(profile_waste_kg.keys()))
        )
        mass_center = _fmt_kg_trim(total_mass_kg_all)
        self.mass_chart.set_data(
            title="Масса профилей",
            subtitle=f"КПД раскроя: {kpd_pct:.1f}%",
            center_top=(mass_center if mass_center else "0"),
            center_bottom="кг",
            values_kg=profile_mass_kg,
            color_map=color_map,
        )
        waste_center = _fmt_kg_trim(waste_total)
        self.waste_chart.set_data(
            title="Обрезки по профилям",
            subtitle="Сегменты по профилям, кг",
            center_top=(waste_center if waste_center else "0"),
            center_bottom="кг обрезков",
            values_kg=profile_waste_kg,
            color_map=color_map,
        )
        logger.info(
            "charts updated: total_mass_kg=%.3f kpd_pct=%.3f waste_kg=%.3f profiles_mass=%d profiles_waste=%d",
            total_mass_kg_all,
            kpd_pct,
            waste_total,
            len(profile_mass_kg),
            len(profile_waste_kg),
        )


if __name__ == "__main__":
    run_app()
