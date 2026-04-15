"""Десктопное приложение на Qt (PySide6)."""

from __future__ import annotations

from pathlib import Path

try:
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QColor, QFont
    from PySide6.QtWidgets import (
        QAbstractItemView,
        QApplication,
        QCheckBox,
        QComboBox,
        QFileDialog,
        QGridLayout,
        QGroupBox,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QTableWidget,
        QTableWidgetItem,
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
from nordfox_raskroy.excel_io import parse_specification
from nordfox_raskroy.export_results import export_cuts_excel, export_cuts_pdf
from nordfox_raskroy.models import CutEvent, PartDemand
from nordfox_raskroy.module_colors import module_row_rgb
from nordfox_raskroy.optimizer import (
    optimize_cutting,
    spec_rows_to_demands,
    summarize,
)
from nordfox_raskroy.result_sort import SORT_MODES, sort_cuts
from nordfox_raskroy.materials_library import row_mass_kg_display
from nordfox_raskroy.profile_codes import (
    PROFILE_DIGIT_TO_NAME,
    filter_spec_by_profiles,
    profile_label_for_code,
)
from nordfox_raskroy.scrap_stock_io import parse_scrap_inventory
from nordfox_raskroy import __version__
from nordfox_raskroy.table_demand_import import demands_from_cut_table_rows


def run_app() -> None:
    app = QApplication([])
    app.setStyle("Fusion")

    win = MainWindow()
    win.show()
    app.exec()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"NordFox — раскрой v{__version__}")
        self.resize(1180, 720)

        self._last_sorted_cuts: list[CutEvent] | None = None
        self._optimizer_cuts: list[CutEvent] | None = None
        self._last_summary_text: str = ""
        self._recommended_bars: tuple[int, ...] | None = None
        self._footer_row_index: int | None = None

        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        top = QHBoxLayout()
        layout.addLayout(top)
        top.addWidget(QLabel("Файл спецификации (.xlsx):"))
        self.path_edit = QLineEdit()
        self.path_edit.setMinimumWidth(520)
        top.addWidget(self.path_edit, stretch=1)
        browse = QPushButton("Обзор…")
        browse.clicked.connect(self._browse)
        top.addWidget(browse)

        scrap_row = QHBoxLayout()
        layout.addLayout(scrap_row)
        scrap_row.addWidget(QLabel("Склад обрезков (.xlsx, колонки «Длина» и «Количество», лист «Склад» или активный):"))
        self.scrap_path_edit = QLineEdit()
        self.scrap_path_edit.setMinimumWidth(400)
        scrap_row.addWidget(self.scrap_path_edit, stretch=1)
        sb = QPushButton("Обзор…")
        sb.clicked.connect(self._browse_scrap)
        scrap_row.addWidget(sb)

        prof = QGroupBox(
            "Профили в раскрое (СК-/СС-/Р-: 0=Н20, 1=Н21, 2=Н22, 3=Н23)"
        )
        pg = QGridLayout(prof)
        self.profile_checks: dict[int, QCheckBox] = {}
        for i, name in PROFILE_DIGIT_TO_NAME.items():
            cb = QCheckBox(f"{name} (цифра {i})")
            cb.setChecked(True)
            self.profile_checks[i] = cb
            pg.addWidget(cb, i // 2, i % 2)
        layout.addWidget(prof)

        bar_fr = QGroupBox("Длины заготовок (мм)")
        bl = QHBoxLayout(bar_fr)
        self.bar6000 = QCheckBox("6000")
        self.bar7500 = QCheckBox("7500")
        self.bar12000 = QCheckBox("12000")
        for b in (self.bar6000, self.bar7500, self.bar12000):
            b.setChecked(True)
            bl.addWidget(b)
        layout.addWidget(bar_fr)

        opt = QHBoxLayout()
        layout.addLayout(opt)
        opt.addWidget(QLabel("Пропил (kerf), мм:"))
        self.kerf_edit = QLineEdit("0")
        self.kerf_edit.setMaximumWidth(72)
        opt.addWidget(self.kerf_edit)
        opt.addWidget(QLabel("Мин. обрезок, мм:"))
        self.scrap_edit = QLineEdit("50")
        self.scrap_edit.setMaximumWidth(72)
        opt.addWidget(self.scrap_edit)
        opt.addStretch()

        btn_row = QHBoxLayout()
        layout.addLayout(btn_row)
        run_btn = QPushButton("Рассчитать раскрой")
        run_btn.clicked.connect(self._compute)
        btn_row.addWidget(run_btn)
        self.btn_xlsx = QPushButton("Экспорт Excel…")
        self.btn_xlsx.clicked.connect(self._export_excel)
        self.btn_xlsx.setEnabled(False)
        btn_row.addWidget(self.btn_xlsx)
        self.btn_pdf = QPushButton("Экспорт PDF…")
        self.btn_pdf.clicked.connect(self._export_pdf)
        self.btn_pdf.setEnabled(False)
        self.btn_recalc = QPushButton("Пересчитать по таблице")
        self.btn_recalc.setToolTip(
            "Учитываются колонки: модуль, тип профиля, длина, угол. "
            "Колонка «Масса» только для отображения. Остальные после пересчёта обновятся."
        )
        self.btn_recalc.clicked.connect(self._recalc_from_table)
        self.btn_recalc.setEnabled(False)
        btn_row.addWidget(self.btn_recalc)
        btn_adv = QPushButton("Подбор заготовок")
        btn_adv.setToolTip(
            "Сравнивает типовые наборы длин (только 6 м, только 7,5 м, комбинации…); "
            "учитывает склад обрезков, если указан файл."
        )
        btn_adv.clicked.connect(self._run_bar_advisor)
        btn_row.addWidget(btn_adv)
        btn_apply_rec = QPushButton("Отметить рекоменд. длины")
        btn_apply_rec.setToolTip("Ставит галочки по последней рекомендации из блока ниже")
        btn_apply_rec.clicked.connect(self._apply_recommended_bars)
        btn_row.addWidget(btn_apply_rec)
        btn_row.addStretch()

        self.advisor_text = QTextEdit()
        self.advisor_text.setReadOnly(True)
        self.advisor_text.setMaximumHeight(110)
        self.advisor_text.setPlaceholderText(
            "Нажмите «Подбор заготовок» для сравнения вариантов 6 / 7,5 / 12 м по текущей спецификации."
        )
        self.advisor_text.setFont(QFont("Segoe UI", 9))
        layout.addWidget(self.advisor_text)

        sort_row = QHBoxLayout()
        layout.addLayout(sort_row)
        sort_row.addWidget(QLabel("Сортировка таблицы:"))
        self.sort_combo = QComboBox()
        for mode_id, label in SORT_MODES:
            self.sort_combo.addItem(label, mode_id)
        self.sort_combo.currentIndexChanged.connect(self._apply_sort)
        sort_row.addWidget(self.sort_combo, stretch=1)

        self.summary = QTextEdit()
        self.summary.setReadOnly(True)
        self.summary.setMaximumHeight(140)
        self.summary.setFont(QFont("Segoe UI", 10))
        layout.addWidget(self.summary)

        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels(
            [
                "Модуль",
                "Тип профиля",
                "Серия",
                "Длина",
                "Угол",
                "Источник",
                "Заготовка мм",
                "Остаток мм",
                "Масса, кг",
            ]
        )
        self.table.setAlternatingRowColors(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.SelectedClicked
        )
        hdr = self.table.horizontalHeader()
        hdr.setStretchLastSection(True)
        self.table.setFont(QFont("Segoe UI", 10))
        layout.addWidget(self.table)

        proj_root = Path(__file__).resolve().parents[2]
        for name in ("spec_20x5_modules.xlsx", "spec_10x5_modules.xlsx"):
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
        demands, err, fw = self._demands_from_spec()
        if demands is None:
            QMessageBox.warning(self, "Подбор заготовок", err)
            return
        try:
            kerf = int(self.kerf_edit.text().strip() or "0")
            min_scrap = int(self.scrap_edit.text().strip() or "0")
        except ValueError:
            QMessageBox.critical(self, "Ошибка", "Пропил и мин. обрезок — целые числа")
            return
        initial, scrap_warns = self._load_initial_scraps()
        outcomes = compare_bar_scenarios(
            demands,
            kerf_mm=kerf,
            min_scrap_mm=min_scrap,
            initial_scraps_mm=initial if initial else None,
        )
        text = format_scenario_report(outcomes)
        extra = list(fw) + list(scrap_warns)
        if extra:
            text += "\n\n" + "\n".join(extra[:15])
        self.advisor_text.setPlainText(text)
        rec = pick_recommended(outcomes)
        self._recommended_bars = rec.bars_mm if rec else None

    def _apply_recommended_bars(self) -> None:
        if not self._recommended_bars:
            QMessageBox.information(
                self,
                "Рекомендация",
                "Сначала выполните «Подбор заготовок».",
            )
            return
        self.bar6000.setChecked(6000 in self._recommended_bars)
        self.bar7500.setChecked(7500 in self._recommended_bars)
        self.bar12000.setChecked(12000 in self._recommended_bars)
        QMessageBox.information(
            self,
            "Длины заготовок",
            f"Отмечены: {', '.join(str(x) for x in sorted(self._recommended_bars))} мм",
        )

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
            export_cuts_excel(
                self._last_sorted_cuts,
                p,
                summary=self._last_summary_text,
            )
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Экспорт Excel", str(e))
            return
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
            export_cuts_pdf(
                self._last_sorted_cuts,
                p,
                summary=self._last_summary_text,
            )
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Экспорт PDF", str(e))
            return
        QMessageBox.information(self, "Экспорт", f"Сохранено:\n{p}")

    def _compute(self) -> None:
        path = self.path_edit.text().strip()
        if not path or not Path(path).is_file():
            QMessageBox.critical(self, "Ошибка", "Укажите существующий файл .xlsx")
            return

        bars: list[int] = []
        if self.bar6000.isChecked():
            bars.append(6000)
        if self.bar7500.isChecked():
            bars.append(7500)
        if self.bar12000.isChecked():
            bars.append(12000)
        if not bars:
            QMessageBox.critical(self, "Ошибка", "Выберите хотя бы одну длину заготовки")
            return

        allowed: set[int] = {d for d, cb in self.profile_checks.items() if cb.isChecked()}
        if not allowed:
            QMessageBox.critical(self, "Ошибка", "Отметьте хотя бы один профиль Н20–Н23")
            return

        try:
            kerf = int(self.kerf_edit.text().strip() or "0")
            min_scrap = int(self.scrap_edit.text().strip() or "0")
        except ValueError:
            QMessageBox.critical(self, "Ошибка", "Пропил и мин. обрезок — целые числа")
            return

        initial_scraps, scrap_warns = self._load_initial_scraps()
        try:
            rows_all = parse_specification(path)
            rows, warns = filter_spec_by_profiles(rows_all, allowed)
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
                return
            demands = spec_rows_to_demands(rows)
            result = optimize_cutting(
                demands,
                bar_lengths_mm=bars,
                kerf_mm=kerf,
                min_scrap_mm=min_scrap,
                initial_scraps_mm=initial_scraps if initial_scraps else None,
            )
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Расчёт", str(e))
            return

        summary_lines = [summarize(result)]
        if warns:
            summary_lines.append(f"Исключено при фильтре: {len(warns)} строк(и)")
        if scrap_warns:
            summary_lines.append("Склад обрезков: " + "; ".join(scrap_warns[:5]))
        full_summary = "\n".join(summary_lines) + (f"\n\n{warn_text}" if warn_text else "")
        self.summary.setPlainText(full_summary)
        self._last_summary_text = full_summary

        self._optimizer_cuts = list(result.cuts)
        self.btn_xlsx.setEnabled(True)
        self.btn_pdf.setEnabled(True)
        self.btn_recalc.setEnabled(True)
        self._apply_sort()

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
            mode = "module"
        cuts = sort_cuts(self._optimizer_cuts, mode)
        self._last_sorted_cuts = cuts
        self._populate_table(cuts)

    def _populate_table(self, cuts: list[CutEvent]) -> None:
        self._footer_row_index = None
        self.table.setRowCount(0)
        total_kg = 0.0
        any_mass = False
        for cut in cuts:
            d = cut.demand
            row = self.table.rowCount()
            self.table.insertRow(row)
            series = profile_label_for_code(d.profile_code)
            kg, mtxt = row_mass_kg_display(d.profile_code, float(d.length_mm), 1.0)
            if kg is not None:
                total_kg += kg
                any_mass = True
            vals = [
                d.module_name,
                d.profile_code,
                series,
                str(d.length_mm),
                str(d.cut_angle),
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
                if col != 8:
                    flags |= Qt.ItemIsEditable
                it.setFlags(flags)
                it.setBackground(bg)
                self.table.setItem(row, col, it)
        if cuts:
            fr = len(cuts)
            self.table.insertRow(fr)
            self._footer_row_index = fr
            foot_bg = QColor(226, 232, 240)
            lab = QTableWidgetItem("Итого, кг")
            lab.setFlags(Qt.ItemIsEnabled)
            lab.setBackground(foot_bg)
            lab.setTextAlignment(int(Qt.AlignRight | Qt.AlignVCenter))
            self.table.setItem(fr, 0, lab)
            self.table.setSpan(fr, 0, 1, 8)
            tot_txt = (
                f"{total_kg:.3f}".rstrip("0").rstrip(".")
                if any_mass
                else "—"
            )
            tot_it = QTableWidgetItem(tot_txt if tot_txt else "0")
            tot_it.setFlags(Qt.ItemIsEnabled)
            tot_it.setBackground(foot_bg)
            self.table.setItem(fr, 8, tot_it)
        self.table.resizeColumnsToContents()

    def _data_rows_for_recalc(self) -> list[list[str]]:
        """Первые 8 колонок, без строки «Итого»."""
        n = self.table.rowCount()
        end = self._footer_row_index if self._footer_row_index is not None else n
        rows: list[list[str]] = []
        for r in range(end):
            row: list[str] = []
            for c in range(8):
                it = self.table.item(r, c)
                row.append(it.text().strip() if it is not None else "")
            rows.append(row)
        return rows

    def _recalc_from_table(self) -> None:
        if self.table.rowCount() == 0:
            QMessageBox.information(self, "Пересчёт", "Таблица пуста.")
            return

        bars: list[int] = []
        if self.bar6000.isChecked():
            bars.append(6000)
        if self.bar7500.isChecked():
            bars.append(7500)
        if self.bar12000.isChecked():
            bars.append(12000)
        if not bars:
            QMessageBox.critical(self, "Ошибка", "Выберите хотя бы одну длину заготовки")
            return

        allowed: set[int] = {d for d, cb in self.profile_checks.items() if cb.isChecked()}
        if not allowed:
            QMessageBox.critical(self, "Ошибка", "Отметьте хотя бы один профиль Н20–Н23")
            return

        try:
            kerf = int(self.kerf_edit.text().strip() or "0")
            min_scrap = int(self.scrap_edit.text().strip() or "0")
        except ValueError:
            QMessageBox.critical(self, "Ошибка", "Пропил и мин. обрезок — целые числа")
            return

        matrix = self._data_rows_for_recalc()
        demands, err = demands_from_cut_table_rows(matrix, allowed)
        if err or demands is None:
            QMessageBox.critical(self, "Таблица", err or "Ошибка разбора таблицы")
            return

        initial_scraps, _sw = self._load_initial_scraps()
        try:
            result = optimize_cutting(
                demands,
                bar_lengths_mm=bars,
                kerf_mm=kerf,
                min_scrap_mm=min_scrap,
                initial_scraps_mm=initial_scraps if initial_scraps else None,
            )
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Пересчёт", str(e))
            return

        note = "\n\n(Пересчёт по отредактированной таблице; колонки «Серия», «Источник», «Заготовка», «Остаток» обновлены.)"
        full = summarize(result) + note
        self.summary.setPlainText(full)
        self._last_summary_text = full
        self._optimizer_cuts = list(result.cuts)
        self._apply_sort()


if __name__ == "__main__":
    run_app()
