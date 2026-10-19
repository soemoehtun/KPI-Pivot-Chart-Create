from __future__ import annotations

import tempfile
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from PySide6.QtCore import QPointF, QObject, QRunnable, QRectF, Qt, QTimer, Signal, Slot, QThreadPool
from PySide6.QtGui import QColor, QImage, QPainter, QPen, QPolygonF
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QSizePolicy,
    QTableView,
    QVBoxLayout,
    QWidget,
    QProgressBar,
)

from kpi_creator.pivot_pipeline import (
    AGG_OPTIONS,
    CHART_OPTIONS,
    LAYOUT_OPTIONS,
    MONTH_ABBR,
    available_dates,
    run_pivot_pipeline,
)
from kpi_creator.services import (
    numeric_columns,
    read_dataset,
    suggest_date_column,
    workbook_sheets,
)
from kpi_creator.ui.widgets import DataFrameModel, MultiSelectPanel, style_table


TEAL = "#0F766E"
INK = "#172033"


def _checkmark_icon_path() -> str:
    """Render a white checkmark PNG to disk and return its path for use in the stylesheet."""
    path = Path(tempfile.gettempdir()) / "kpi_pivot_checkmark.png"
    image = QImage(16, 16, QImage.Format.Format_ARGB32)
    image.fill(Qt.GlobalColor.transparent)
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    pen = QPen(QColor("#FFFFFF"))
    pen.setWidthF(2.4)
    pen.setCapStyle(Qt.PenCapStyle.RoundCap)
    pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    painter.setPen(pen)
    painter.drawPolyline(QPolygonF([QPointF(3.5, 8.5), QPointF(6.5, 12), QPointF(13, 4)]))
    painter.end()
    image.save(str(path), "PNG")
    return str(path).replace("\\", "/")


_GITHUB_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 640"><path d="M280.5 426.5C214.5 418.5 168 371 168 309.5C168 284.5 177 257.5 192 239.5C185.5 223 186.5 188 194 173.5C214 171 241 181.5 257 196C276 190 296 187 320.5 187C345 187 365 190 383 195.5C398.5 181.5 426 171 446 173.5C453 187 454 222 447.5 239C463.5 258 472 283.5 472 309.5C472 371 425.5 417.5 358.5 426C375.5 437 387 461 387 488.5L387 540.5C387 555.5 399.5 564 414.5 558C505 523.5 576 433 576 321C576 179.5 461 64 319.5 64C178 64 64 179.5 64 321C64 432 134.5 524 229.5 558.5C243 563.5 256 554.5 256 541L256 501C249 504 240 506 232 506C199 506 179.5 488 165.5 454.5C160 441 154 433 142.5 431.5C136.5 431 134.5 428.5 134.5 425.5C134.5 419.5 144.5 415 154.5 415C169 415 181.5 424 194.5 442.5C204.5 457 215 463.5 227.5 463.5C240 463.5 248 459 259.5 447.5C268 439 274.5 431.5 280.5 426.5z"/></svg>"""


def _github_icon_path() -> str:
    """Render the GitHub octocat SVG to a PNG file and return its path."""
    path = Path(tempfile.gettempdir()) / "kpi_pivot_github.png"
    renderer = QSvgRenderer(_GITHUB_SVG.encode())
    image = QImage(32, 32, QImage.Format.Format_ARGB32)
    image.fill(Qt.GlobalColor.transparent)
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    renderer.render(painter, QRectF(0, 0, 32, 32))
    painter.end()
    image.save(str(path), "PNG")
    return str(path).replace("\\", "/")


class TaskSignals(QObject):
    succeeded = Signal(object)
    failed = Signal(str)
    progress = Signal(tuple)


class BackgroundTask(QRunnable):
    def __init__(self, function, emits_progress: bool = False):
        super().__init__()
        self.function = function
        self.emits_progress = emits_progress
        self.signals = TaskSignals()

    @Slot()
    def run(self):
        try:
            if self.emits_progress:
                result = self.function(self.signals.progress.emit)
            else:
                result = self.function()
            self.signals.succeeded.emit(result)
        except Exception as exc:
            self.signals.failed.emit(str(exc))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("KPI Pivot Chart Creator")
        self.setMinimumSize(1060, 720)
        self.resize(1320, 860)
        self.pool = QThreadPool.globalInstance()
        self._tasks: list[BackgroundTask] = []
        self.source_path: str | None = None
        self.sheet_names: list[str] = []
        self.custom_sheet_names: list[str] | None = None
        self.import_warnings: list[str] = []
        self.text_import_options = {
            "delimiter": "Auto-detect", "encoding": "utf-8-sig", "header": True,
            "decimal": ".", "null_values": "",
        }
        self.frame: pd.DataFrame | None = None
        self.busy = False
        self.available_days: list = []
        self._build_ui()
        self._set_style()
        self._update_summary_controls()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        QTimer.singleShot(0, self._sync_left_frame_width)

    def _sync_left_frame_width(self):
        changed = False
        if hasattr(self, "report_left_box") and hasattr(self, "source_card"):
            target = self.source_card.width()
            if self.report_left_box.width() != target:
                self.report_left_box.setFixedWidth(target)
                changed = True
        browse = getattr(self, "browse_button", None)
        choose = getattr(self, "choose_folder_button", None)
        if browse is not None and choose is not None and choose.width() != browse.width():
            choose.setFixedWidth(browse.width())
            changed = True
        file_w = getattr(self, "file_path_edit", None)
        out_w = getattr(self, "output_edit", None)
        if file_w is not None and out_w is not None and out_w.width() != file_w.width():
            out_w.setFixedWidth(file_w.width())
            changed = True
        if changed:
            QTimer.singleShot(0, self._sync_left_frame_width)

    def _build_ui(self):
        root = QWidget()
        root.setObjectName("appRoot")
        root.setAutoFillBackground(True)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(16, 10, 16, 12)
        layout.setSpacing(8)

        top = QHBoxLayout()
        title = QLabel("KPI Pivot Chart Creator")
        title.setObjectName("pageTitle")
        top.addWidget(title)
        top.addStretch(1)
        self.links_label = QLabel(
            "<a href=\"https://soemoetun.vercel.app/\">My Profile</a>"
        )
        self.links_label.setObjectName("topLinks")
        self.links_label.setContentsMargins(0, 0, 12, 0)
        self.links_label.setOpenExternalLinks(True)
        top.addWidget(self.links_label)
        layout.addLayout(top)

        body = QWidget()
        body.setObjectName("generatorBody")
        body.setAutoFillBackground(True)
        form = QVBoxLayout(body)
        form.setContentsMargins(0, 0, 0, 0)
        form.setSpacing(8)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)
        source = self._card()
        source.setMinimumWidth(320)
        self.source_card = source
        source_layout = QVBoxLayout(source)
        source_layout.setContentsMargins(14, 10, 14, 10)
        source_layout.setSpacing(6)
        source_layout.addLayout(self._numbered_heading("01", "Source file"))
        source_layout.addWidget(self._label("Excel or CSV workbook"))
        file_row = QHBoxLayout()
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        self.file_path_edit.setPlaceholderText("Select a source file")
        file_row.addWidget(self.file_path_edit, 1)
        browse = QPushButton("Browse Files")
        browse.setObjectName("primarySmall")
        browse.clicked.connect(self.choose_file)
        self.browse_button = browse
        file_row.addWidget(browse)
        self.csv_options_button = QPushButton("Text options…")
        self.csv_options_button.setObjectName("secondaryButton")
        self.csv_options_button.clicked.connect(self.configure_text_import)
        self.csv_options_button.hide()
        file_row.addWidget(self.csv_options_button)
        source_layout.addLayout(file_row)
        meta_row = QHBoxLayout()
        self.file_meta_label = QLabel("No file selected  ·  .xlsx, .xlsm, .xlsb, .csv, .txt")
        self.file_meta_label.setObjectName("muted")
        meta_row.addWidget(self.file_meta_label, 1)
        self.preview_button = QPushButton("Preview rows")
        self.preview_button.setObjectName("textButton")
        self.preview_button.setEnabled(False)
        self.preview_button.clicked.connect(self.show_preview)
        meta_row.addWidget(self.preview_button)
        source_layout.addLayout(meta_row)
        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setObjectName("divider")
        source_layout.addWidget(divider)
        source_layout.addWidget(self._label("Columns"))
        self.sheet_combo = self._combo("Select a worksheet")
        self.sheet_combo.currentTextChanged.connect(self._sheet_changed)
        sheet_row = QHBoxLayout()
        sheet_row.addWidget(self.sheet_combo, 1)
        self.custom_sheet_button = QPushButton("Custom…")
        self.custom_sheet_button.setObjectName("secondaryButton")
        self.custom_sheet_button.clicked.connect(self.choose_custom_sheets)
        self.custom_sheet_button.hide()
        sheet_row.addWidget(self.custom_sheet_button)
        source_layout.addWidget(self._field("Sheet", sheet_row))
        date_time_row = QHBoxLayout()
        self.date_combo = self._combo("Select date / category")
        self.time_combo = self._combo("None")
        self.date_combo.currentTextChanged.connect(self._refresh_available_periods)
        date_time_row.addWidget(self._field("Date column", self.date_combo), 1)
        date_time_row.addWidget(self._field("Time column (optional)", self.time_combo), 1)
        source_layout.addLayout(date_time_row)
        self.group_combo = self._combo("None")
        source_layout.addWidget(self._field("Group / series by", self.group_combo))
        helper = QLabel("Column selections apply to every KPI sheet in the workbook.")
        helper.setObjectName("hintBox")
        helper.setWordWrap(True)
        source_layout.addWidget(helper)
        source_layout.addStretch(1)
        self.source_ready = QLabel("Ready. Select a workbook to get started.")
        self.source_ready.setObjectName("statusText")
        source_layout.addWidget(self.source_ready)
        top_row.addWidget(source, 12)

        kpi_card = self._card()
        kpi_layout = QVBoxLayout(kpi_card)
        kpi_layout.setContentsMargins(14, 10, 14, 10)
        self.kpi_panel = MultiSelectPanel("02   KPI value columns", "Select one or more metrics to visualize.")
        kpi_layout.addWidget(self.kpi_panel)
        top_row.addWidget(kpi_card, 10)

        slicer_card = self._card()
        slicer_layout = QVBoxLayout(slicer_card)
        slicer_layout.setContentsMargins(14, 10, 14, 10)
        self.slicer_panel = MultiSelectPanel("03   Slicer fields", "Add interactive filters to the workbook.")
        slicer_layout.addWidget(self.slicer_panel)
        top_row.addWidget(slicer_card, 10)
        form.addLayout(top_row, 1)

        report = self._card()
        report_layout = QVBoxLayout(report)
        report_layout.setContentsMargins(14, 10, 14, 10)
        report_layout.setSpacing(8)
        header = QHBoxLayout()
        header.addLayout(self._numbered_heading("04", "Report options"))
        header.addStretch(1)
        note = QLabel("Creates native Excel PivotTables, PivotCharts, and slicers.")
        note.setObjectName("muted")
        header.addWidget(note)
        report_layout.addLayout(header)

        lower = QHBoxLayout()
        lower.setSpacing(12)
        lower.setContentsMargins(0, 0, 0, 0)
        left_box = self._sub_box()
        self.report_left_box = left_box
        left_col = left_box.layout()
        self.aggregation_combo = self._combo("Average")
        self.aggregation_combo.addItems(list(AGG_OPTIONS))
        self.aggregation_combo.setCurrentIndex(0)
        left_col.addWidget(self._field("Aggregation", self.aggregation_combo))
        self.chart_combo = self._combo("Column")
        self.chart_combo.addItems(list(CHART_OPTIONS))
        self.chart_combo.setCurrentIndex(0)
        left_col.addWidget(self._field("Chart type", self.chart_combo))
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("Choose a folder for the exported workbook")
        choose_folder = QPushButton("Choose…")
        choose_folder.setObjectName("primarySmall")
        choose_folder.clicked.connect(self.choose_output_folder)
        self.choose_folder_button = choose_folder
        output_row = QHBoxLayout()
        output_row.setContentsMargins(0, 0, 0, 0)
        output_row.addWidget(self.output_edit, 1)
        output_row.addWidget(choose_folder)
        left_col.addWidget(self._field("Output folder", output_row))
        left_col.addWidget(self._label("Pivot layout"))
        self.layout_buttons = QButtonGroup(self)
        self.layout_radios: dict[str, QRadioButton] = {}
        for label, value in LAYOUT_OPTIONS.items():
            radio = QRadioButton(label)
            radio.setProperty("layout_value", value)
            self.layout_buttons.addButton(radio)
            radio.toggled.connect(self._update_summary_controls)
            left_col.addWidget(radio)
            self.layout_radios[value] = radio
        self.layout_radios["all_pivot"].setChecked(True)
        left_col.addStretch(1)
        lower.addWidget(left_box, 4)

        right_box = self._sub_box()
        right_col = right_box.layout()
        self.summary_check = QCheckBox("Include a summary table")
        self.summary_check.setChecked(True)
        self.summary_check.toggled.connect(self._update_summary_controls)
        right_col.addWidget(self.summary_check)
        right_col.setAlignment(self.summary_check, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        comparison_group = QVBoxLayout()
        comparison_group.setContentsMargins(0, 0, 0, 0)
        comparison_group.setSpacing(6)
        comparison_group.addWidget(self._label("Compare by"))
        period_row = QHBoxLayout()
        period_row.setSpacing(6)
        period_row.setContentsMargins(0, 0, 0, 0)
        self.period_buttons: dict[str, QPushButton] = {}
        for period in ("Day", "Month", "Year"):
            button = QPushButton(period)
            button.setObjectName("periodButton")
            button.setCheckable(True)
            button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            button.clicked.connect(lambda checked=False, p=period: self.set_period(p))
            self.period_buttons[period] = button
            period_row.addWidget(button)
        self.period_buttons["Month"].setChecked(True)
        comparison_group.addLayout(period_row)
        self.comparison_combo = self._combo("Previous Month vs Last Month")
        self.comparison_combo.addItems(["Previous month vs last month", "Custom PRE vs POST"])
        self.comparison_combo.setCurrentIndex(0)
        comparison_group.addWidget(self._field("Compare", self.comparison_combo))
        self.custom_period_widget = QWidget()
        self.custom_outer = QHBoxLayout(self.custom_period_widget)
        self.custom_outer.setContentsMargins(0, 0, 0, 0)
        self.custom_outer.setSpacing(14)
        self.custom_outer.addWidget(self._period_side("pre"), 1)
        self.custom_outer.addWidget(self._period_side("post"), 1)
        self.custom_period_widget.hide()
        comparison_group.addWidget(self.custom_period_widget)
        self.comparison_combo.currentTextChanged.connect(lambda text: self.custom_period_widget.setVisible(text.startswith("Custom")))
        self.period_hint = QLabel("")
        self.period_hint.setObjectName("muted")
        self.period_hint.setWordWrap(True)
        comparison_group.addWidget(self.period_hint)
        comparison_group.addStretch(1)
        self.compare_widget = QWidget()
        self.compare_widget.setLayout(comparison_group)
        self.compare_widget.hide()
        right_col.addWidget(self.compare_widget, 1)
        lower.addWidget(right_box, 6)
        report_layout.addLayout(lower)
        self._update_summary_controls()

        footer = QHBoxLayout()
        footer.setSpacing(10)
        self.status_title = QLabel("Ready.")
        self.status_title.setObjectName("statusText")
        self.status_message = QLabel("Select a workbook to get started")
        self.status_message.setObjectName("muted")
        status_box = QWidget()
        status_layout = QVBoxLayout(status_box)
        status_layout.setContentsMargins(0, 0, 0, 0)
        status_layout.setSpacing(2)
        status_layout.addWidget(self.status_title)
        status_layout.addWidget(self.status_message)
        footer.addWidget(status_box, 1)
        self.progress = QProgressBar()
        self.progress.setObjectName("progress")
        self.progress.setTextVisible(False)
        self.progress.setFixedWidth(150)
        self.progress.setMaximumHeight(8)
        self.progress.hide()
        footer.addWidget(self.progress)
        self.progress_percent = QLabel("")
        self.progress_percent.setObjectName("progressPercent")
        self.progress_percent.hide()
        footer.addWidget(self.progress_percent)
        self.generate_button = QPushButton("Generate Excel Report")
        self.generate_button.setObjectName("primaryLarge")
        self.generate_button.clicked.connect(self.generate_report)
        footer.addWidget(self.generate_button)
        report_layout.addLayout(footer)
        form.addWidget(report)
        layout.addWidget(body, 1)
        self.setCentralWidget(root)

    def _card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        return card

    def _sub_box(self) -> QFrame:
        box = QFrame()
        box.setObjectName("subBox")
        layout = QVBoxLayout(box)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(6)
        return box

    def _label(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("fieldLabel")
        return label

    def _combo(self, placeholder: str) -> QComboBox:
        combo = QComboBox()
        combo.setMinimumHeight(26)
        combo.setPlaceholderText(placeholder)
        combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        return combo

    def _field(self, label: str, widget_or_layout) -> QWidget:
        box = QWidget()
        layout = QVBoxLayout(box)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(3)
        layout.addWidget(self._label(label))
        if isinstance(widget_or_layout, QHBoxLayout):
            layout.addLayout(widget_or_layout)
        else:
            layout.addWidget(widget_or_layout)
        return box

    def _numbered_heading(self, number: str, title: str) -> QHBoxLayout:
        row = QHBoxLayout()
        badge = QLabel(number)
        badge.setObjectName("stepBadge")
        badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        row.addWidget(badge)
        label = QLabel(title)
        label.setObjectName("sectionTitle")
        row.addWidget(label)
        row.addStretch(1)
        return row

    def _vertical_divider(self) -> QFrame:
        frame = QFrame()
        frame.setFrameShape(QFrame.Shape.VLine)
        frame.setObjectName("divider")
        return frame

    def _set_style(self):
        check_icon = _checkmark_icon_path()
        self.setStyleSheet(f"""
            QWidget {{ color: {INK}; font-family: 'Segoe UI', Arial; font-size: 12px; }}
            QMainWindow, #appRoot, #generatorBody, QScrollArea::viewport {{ background-color: #F4F6F8; }}
            #pageTitle {{ color: {TEAL}; font-size: 18px; font-weight: 700; }}
            #localBadge, #countBadge {{ color: {TEAL}; background: #CCFBF1; border-radius: 3px; padding: 3px 7px; font-weight: 600; font-size: 10px; }}
            #topLinks {{ color: #475467; font-weight: 600; font-size: 11px; }}
            #topLinks a {{ color: {TEAL}; text-decoration: none; }}
            #topLinks a:hover {{ text-decoration: underline; }}
            #card {{ background: #FFFFFF; border: 1px solid #DFE4EA; border-radius: 3px; }}
            #sectionTitle {{ font-size: 13px; font-weight: 700; }}
            #stepBadge {{ min-width: 26px; max-width: 26px; height: 20px; color: {TEAL}; background: #CCFBF1; font-size: 11px; font-weight: 700; }}
            #fieldLabel {{ color: #475467; font-weight: 600; }}
            #muted {{ color: #667085; }}
            #hintBox {{ color: #667085; background: #F8FAFC; padding: 6px; }}
            #statusText {{ color: {TEAL}; font-weight: 700; }}
            #progressPercent {{ color: {TEAL}; font-weight: 800; font-size: 15px; }}
            #divider {{ color: #E4E7EC; }}
            QLineEdit, QComboBox, QDateEdit {{ background: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 3px; padding: 3px 6px; min-height: 16px; selection-background-color: #99F6E4; }}
            QLineEdit:focus, QComboBox:focus, QDateEdit:focus {{ border: 1px solid {TEAL}; }}
            QPushButton {{ cursor: pointer; }}
            #primarySmall, #primaryLarge {{ color: #FFFFFF; background: {TEAL}; border: 1px solid {TEAL}; border-radius: 2px; font-weight: 650; }}
            #primarySmall {{ padding: 4px 10px; }}
            #primaryLarge {{ padding: 7px 14px; font-size: 13px; }}
            #primarySmall:hover, #primaryLarge:hover {{ background: #115E59; }}
            #secondaryButton {{ background: #FFFFFF; color: {TEAL}; border: 1px solid #B2DAD4; border-radius: 4px; padding: 4px 8px; font-weight: 600; }}
            #periodButton {{ background: #F2F4F7; border: 1px solid #E4E7EC; border-radius: 4px; padding: 7px 12px; color: #475467; font-weight: 600; }}
            #periodButton:checked {{ background: #CCFBF1; color: {TEAL}; border-color: #99F6E4; font-weight: 700; }}
            #subBox {{ background: #FFFFFF; border: 1px solid #DFE4EA; border-radius: 3px; }}
            #listBox {{ border: 1px solid #D0D5DD; background: #FFFFFF; }}
            #emptyMessage {{ color: #8A94A6; padding: 10px; }}
            QListWidget {{ color: {INK}; background: #FFFFFF; border: 0; alternate-background-color: #FFFFFF; }}
            QListWidget::item {{ color: {INK}; background: #FFFFFF; padding: 3px 4px; }}
            QListWidget::item:selected {{ color: {TEAL}; background: #CCFBF1; }}
            QCheckBox {{ spacing: 6px; }}
            QCheckBox::indicator {{ width: 11px; height: 11px; }}
            QCheckBox::indicator:unchecked {{ border: 1px solid #98A2B3; background: #FFFFFF; border-radius: 2px; }}
            QCheckBox::indicator:unchecked:hover {{ border-color: {TEAL}; }}
            QCheckBox::indicator:checked {{ border: 1px solid {TEAL}; background: {TEAL}; border-radius: 2px; image: url("{check_icon}"); }}
            QCheckBox::indicator:disabled {{ border: 1px solid #D0D5DD; background: #F8FAFC; }}
            QTableView {{ background: #FFFFFF; border: 0; alternate-background-color: #F8FAFC; selection-background-color: #CCFBF1; }}
            QHeaderView::section {{ color: #475467; background: #F8FAFC; border: 0; border-bottom: 1px solid #E4E7EC; padding: 5px 6px; font-weight: 650; }}
            #textButton {{ border: 0; background: transparent; color: {TEAL}; font-weight: 600; padding: 3px 5px; }}
            #progress {{ border: 0; background: #D1FAE5; border-radius: 4px; }}
            #progress::chunk {{ background: {TEAL}; border-radius: 4px; }}
        """)

    def set_period(self, period: str):
        for name, button in self.period_buttons.items():
            button.setChecked(name == period)
        if period == "Day":
            options = ["Previous day vs last day", "Last day vs day - 7", "Custom PRE vs POST"]
        elif period == "Month":
            options = ["Previous month vs last month", "Custom PRE vs POST"]
        else:
            options = ["Previous year vs last year", "Custom PRE vs POST"]
        old = self.comparison_combo.currentText()
        self.comparison_combo.blockSignals(True)
        self.comparison_combo.clear()
        self.comparison_combo.addItems(options)
        if old in options:
            self.comparison_combo.setCurrentText(old)
        else:
            self.comparison_combo.setCurrentIndex(0)
        self.comparison_combo.blockSignals(False)
        self._refresh_custom_periods()
        self.custom_period_widget.setVisible(self.comparison_combo.currentText().startswith("Custom"))

    def _period_side(self, name: str) -> QWidget:
        box = QWidget()
        layout = QVBoxLayout(box)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        badge = QLabel(name)
        badge.setStyleSheet(
            "background:#CCFBF1;color:#0F766E;font-weight:700;font-size:11px;"
            "padding:2px 8px;border-radius:3px;"
        )
        layout.addWidget(badge, 0, Qt.AlignmentFlag.AlignLeft)
        combos_layout = QHBoxLayout()
        combos_layout.setSpacing(8)
        for label, attr in (("Year", f"{name}_year"), ("Month", f"{name}_month"), ("Day", f"{name}_day")):
            combo = self._combo("")
            combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            setattr(self, attr, combo)
            field = self._field(label, combo)
            setattr(self, f"{attr}_field", field)
            combos_layout.addWidget(field, 1)
        layout.addLayout(combos_layout)
        layout.addStretch(1)
        return box

    def _build_custom_period(self, side: str) -> str:
        year = getattr(self, f"{side}_year").currentData()
        if year in (None, ""):
            return ""
        level = self._current_level()
        if level == "year":
            return str(year)
        month = getattr(self, f"{side}_month").currentData()
        if month in (None, ""):
            return ""
        if level == "month":
            return f"{month[0]:04d}-{month[1]:02d}"
        day = getattr(self, f"{side}_day").currentData()
        if day in (None, ""):
            return ""
        return f"{day[0]:04d}-{day[1]:02d}-{day[2]:02d}"

    def _refresh_custom_periods(self):
        """Fill the PRE/POST Year/Month/Day dropdowns with only values present in the loaded data."""
        if not hasattr(self, "pre_year"):
            return
        level = self._current_level()
        days = self.available_days
        year_opts = sorted({day.year for day in days})
        month_opts = sorted({(day.year, day.month) for day in days})
        day_opts = sorted({(day.year, day.month, day.day) for day in days})

        def populate(combo, options, formatter, fallback):
            previous = combo.currentData()
            combo.blockSignals(True)
            combo.clear()
            for option in options:
                combo.addItem(formatter(option), option)
            if previous in options:
                combo.setCurrentIndex(options.index(previous))
            elif options:
                index = len(options) + fallback
                combo.setCurrentIndex(index if 0 <= index < len(options) else 0)
            combo.blockSignals(False)

        def year_fmt(option):
            return str(option)

        def month_fmt(option):
            return "{}".format(MONTH_ABBR[option[1] - 1])

        def day_fmt(option):
            return date(*option).strftime("%d")

        for side, fallback in (("pre", -2), ("post", -1)):
            populate(getattr(self, f"{side}_year"), year_opts, year_fmt, fallback)
            populate(getattr(self, f"{side}_month"), month_opts, month_fmt, fallback)
            populate(getattr(self, f"{side}_day"), day_opts, day_fmt, fallback)
            getattr(self, f"{side}_month_field").setVisible(level in ("month", "day") and bool(month_opts))
            getattr(self, f"{side}_day_field").setVisible(level == "day" and bool(day_opts))

    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Choose a KPI workbook", "",
            "Supported data files (*.csv *.txt *.xlsx *.xlsm *.xlsb);;All files (*.*)"
        )
        if not path:
            return
        self.source_path = path
        self.custom_sheet_names = None
        self.file_path_edit.setText(path)
        self.file_meta_label.setText(f"{Path(path).name}  ·  {Path(path).suffix.upper()[1:]} · processed locally")
        is_text = Path(path).suffix.lower() in {".csv", ".txt"}
        self.csv_options_button.setVisible(is_text)
        try:
            self.sheet_names = workbook_sheets(path)
        except Exception as exc:
            self._show_error("Could not read workbook", str(exc))
            return
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        is_excel = Path(path).suffix.lower() in {".xlsx", ".xlsm", ".xlsb"}
        if is_excel:
            self.sheet_combo.addItem("All Sheets")
            self.sheet_combo.addItems(self.sheet_names)
            self.sheet_combo.setCurrentIndex(1 if self.sheet_names else 0)
        else:
            self.sheet_combo.addItem("Data")
            self.sheet_combo.setCurrentIndex(0)
        self.sheet_combo.setEnabled(is_excel)
        self.custom_sheet_button.setVisible(is_excel)
        self.sheet_combo.blockSignals(False)
        self._load_selected_sheet()

    def configure_text_import(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("CSV / TXT import options")
        dialog.setMinimumWidth(420)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        delimiter = QComboBox()
        delimiter.addItems(["Auto-detect", "Comma ,", "Tab", "Semicolon ;", "Pipe |", "Custom"])
        custom_delimiter = QLineEdit()
        custom_delimiter.setMaxLength(4)
        custom_delimiter.setPlaceholderText("Enter a delimiter character")
        encoding = QComboBox()
        encoding.addItems(["utf-8-sig", "utf-8", "cp1252", "latin-1"])
        header = QCheckBox("First row contains column names")
        header.setChecked(True)
        decimal = QComboBox()
        decimal.addItem("Period (.)", ".")
        decimal.addItem("Comma (,) locale", ",")
        null_values = QLineEdit()
        null_values.setPlaceholderText("Optional, comma-separated values")
        current_delimiter = self.text_import_options["delimiter"]
        delimiter_labels = {"Auto-detect": "Auto-detect", ",": "Comma ,", "\t": "Tab", ";": "Semicolon ;", "|": "Pipe |"}
        delimiter.setCurrentText(delimiter_labels.get(current_delimiter, "Custom"))
        if delimiter.currentText() == "Custom":
            custom_delimiter.setText(current_delimiter)
        encoding.setCurrentText(self.text_import_options["encoding"])
        header.setChecked(self.text_import_options["header"])
        decimal.setCurrentIndex(0 if self.text_import_options["decimal"] == "." else 1)
        null_values.setText(self.text_import_options["null_values"])
        form.addRow("Delimiter", delimiter)
        form.addRow("Custom delimiter", custom_delimiter)
        form.addRow("Encoding", encoding)
        form.addRow("Header row", header)
        form.addRow("Decimal separator", decimal)
        form.addRow("Null values", null_values)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        delimiter_map = {
            "Auto-detect": "Auto-detect", "Comma ,": ",", "Tab": "\t",
            "Semicolon ;": ";", "Pipe |": "|",
        }
        selected_delimiter = custom_delimiter.text() if delimiter.currentText() == "Custom" else delimiter_map[delimiter.currentText()]
        if delimiter.currentText() == "Custom" and not selected_delimiter:
            self._show_error("Delimiter required", "Enter a delimiter character or choose one of the presets.")
            return
        self.text_import_options = {
            "delimiter": selected_delimiter,
            "encoding": encoding.currentText(),
            "header": header.isChecked(),
            "decimal": decimal.currentData(),
            "null_values": null_values.text(),
        }
        if self.source_path:
            self._load_selected_sheet()

    def choose_custom_sheets(self):
        if not self.sheet_names:
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Choose worksheets")
        dialog.resize(360, 480)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel("Select worksheets to combine. Matching headers are combined."))
        items = QListWidget()
        for name in self.sheet_names:
            item = QListWidgetItem(name)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            initial = self.custom_sheet_names or [self.sheet_names[0]]
            item.setCheckState(Qt.CheckState.Checked if name in initial else Qt.CheckState.Unchecked)
            items.addItem(item)
        layout.addWidget(items, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        selected = [items.item(i).text() for i in range(items.count()) if items.item(i).checkState() == Qt.CheckState.Checked]
        if not selected:
            self._show_error("No worksheets selected", "Choose at least one worksheet.")
            return
        self.custom_sheet_names = selected
        label = f"Custom Sheets ({len(selected)})"
        self.sheet_combo.blockSignals(True)
        for index in reversed(range(self.sheet_combo.count())):
            if self.sheet_combo.itemText(index).startswith("Custom Sheets ("):
                self.sheet_combo.removeItem(index)
        self.sheet_combo.addItem(label)
        index = self.sheet_combo.count() - 1
        self.sheet_combo.setCurrentIndex(index)
        self.sheet_combo.blockSignals(False)
        self._load_selected_sheet()

    def show_preview(self):
        if self.frame is None:
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Import preview · first 100 rows")
        dialog.resize(1000, 620)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(f"{Path(self.source_path).name}  ·  {len(self.frame):,} total rows × {len(self.frame.columns)} columns"))
        table = QTableView()
        model = DataFrameModel(self.frame.head(100), dialog)
        style_table(table, model)
        table.resizeColumnsToContents()
        layout.addWidget(table, 1)
        close = QPushButton("Close")
        close.clicked.connect(dialog.accept)
        layout.addWidget(close, 0, Qt.AlignmentFlag.AlignRight)
        dialog.exec()

    def choose_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Choose output folder", self.output_edit.text() or "")
        if folder:
            self.output_edit.setText(folder)

    def _sheet_changed(self, _text):
        text = self.sheet_combo.currentText()
        if text.startswith("Custom Sheets ("):
            if self.custom_sheet_names:
                self._load_selected_sheet()
            else:
                self.choose_custom_sheets()
            return
        if text != "All Sheets":
            self.custom_sheet_names = None
        if self.source_path and not self.busy:
            self._load_selected_sheet()

    def _load_selected_sheet(self):
        if not self.source_path:
            return
        self._set_busy(True, "Loading source file…")
        path = self.source_path
        selection = self.sheet_combo.currentText() or None
        if selection and selection.startswith("Custom Sheets ("):
            sheet = list(self.custom_sheet_names or [])
        elif selection == "All Sheets":
            sheet = "All Sheets"
        else:
            sheet = selection
        options = dict(self.text_import_options)
        task = BackgroundTask(lambda: read_dataset(path, sheet, options))
        task.signals.succeeded.connect(self._dataset_loaded)
        task.signals.failed.connect(self._dataset_failed)
        self._tasks.append(task)
        self.pool.start(task)

    @Slot(object)
    def _dataset_loaded(self, frame: pd.DataFrame):
        self._tasks.clear()
        self.frame = frame
        self.import_warnings = list(frame.attrs.get("import_warnings", []))
        self._set_busy(False)
        self._configure_columns()
        self.preview_button.setEnabled(True)
        self.file_meta_label.setText(f"{Path(self.source_path).name}  ·  {len(frame):,} rows × {len(frame.columns):,} columns")
        self.source_ready.setText("Workbook loaded.")
        self.status_title.setText("Ready.")
        self.status_message.setText(f"{len(frame):,} records · {len(frame.columns)} columns available")

    @Slot(str)
    def _dataset_failed(self, message: str):
        self._tasks.clear()
        self._set_busy(False)
        self._show_error("Could not load source file", message)
        self.source_ready.setText("Could not load the selected file.")

    def _configure_columns(self):
        if self.frame is None:
            return
        columns = [str(col) for col in self.frame.columns]
        self.date_combo.blockSignals(True)
        self.date_combo.clear()
        self.date_combo.addItems(columns)
        suggestion = suggest_date_column(self.frame)
        if suggestion:
            self.date_combo.setCurrentText(suggestion)
        elif columns:
            self.date_combo.setCurrentIndex(0)
        self.date_combo.blockSignals(False)
        self.time_combo.clear()
        self.time_combo.addItem("None")
        self.time_combo.addItems(columns)
        time_suggestion = next((c for c in columns if "time" in c.lower()), None)
        if time_suggestion:
            self.time_combo.setCurrentText(time_suggestion)
        else:
            self.time_combo.setCurrentIndex(0)
        self.group_combo.clear()
        self.group_combo.addItem("None")
        self.group_combo.addItems(columns)
        self.group_combo.setCurrentIndex(0)
        ncols = numeric_columns(self.frame)
        self.kpi_panel.set_items(ncols, ncols[:min(3, len(ncols))])
        self.slicer_panel.set_items(columns, [])
        self._refresh_available_periods()

    def _refresh_available_periods(self, _column: str = ""):
        if self.frame is None or not self.date_combo.currentText():
            self.available_days = []
            self.period_hint.setText("")
            return
        column = self.date_combo.currentText()
        if column not in self.frame.columns:
            return
        self.available_days = available_dates(self.frame, column)
        self._refresh_custom_periods()
        if self.available_days:
            first = self.available_days[0].strftime("%d-%b-%Y")
            last = self.available_days[-1].strftime("%d-%b-%Y")
            self.period_hint.setText(f"Source date range: {first} to {last}.")
        else:
            self.period_hint.setText("No usable dates in the selected column; period comparison is unavailable.")

    def _selected_layout(self) -> str:
        button = self.layout_buttons.checkedButton()
        return str(button.property("layout_value")) if button else "all_pivot"

    def _update_summary_controls(self, *_args):
        if not hasattr(self, "summary_check"):
            return
        allowed = self._selected_layout() != "per_sheet"
        self.summary_check.setEnabled(allowed)
        if not allowed and self.summary_check.isChecked():
            self.summary_check.blockSignals(True)
            self.summary_check.setChecked(True)
            self.summary_check.blockSignals(False)
        active = self.summary_check.isChecked() and allowed
        self.compare_widget.setVisible(active)
        box = self.compare_widget.parentWidget()
        layout = box.layout() if box else None
        in_layout = False
        if layout:
            in_layout = any(layout.itemAt(i).widget() is self.compare_widget for i in range(layout.count()))
        if active and not in_layout:
            layout.addWidget(self.compare_widget, 1)
        elif not active and in_layout:
            layout.removeWidget(self.compare_widget)
        if not active:
            self.custom_period_widget.hide()
        else:
            self.custom_period_widget.setVisible(self.comparison_combo.currentText().startswith("Custom"))

    def _current_level(self) -> str:
        period = next((name for name, button in self.period_buttons.items() if button.isChecked()), "Month")
        return period.lower()

    def _auto_periods(self) -> tuple[str, str]:
        level = self._current_level()
        days = self.available_days
        if not days:
            return "", ""
        mode = self.comparison_combo.currentText()
        if level == "day":
            last = days[-1]
            if "day - 7" in mode:
                cutoff = last - timedelta(days=7)
                earlier = [day for day in days if day <= cutoff]
                previous = earlier[-1] if earlier else days[0]
            else:
                previous = days[-2] if len(days) > 1 else last
            return previous.isoformat(), last.isoformat()
        if level == "month":
            periods = sorted({(day.year, day.month) for day in days})
            previous = periods[-2] if len(periods) > 1 else periods[-1]
            latest = periods[-1]
            return f"{previous[0]}-{previous[1]:02d}", f"{latest[0]}-{latest[1]:02d}"
        years = sorted({day.year for day in days})
        previous = years[-2] if len(years) > 1 else years[-1]
        return str(previous), str(years[-1])

    def _read_pivot_params(self) -> dict:
        if self.frame is None or not self.source_path:
            raise ValueError("Choose a source file first.")
        kpis = self.kpi_panel.selected_items()
        if not kpis:
            raise ValueError("Select at least one KPI value column.")
        date_column = self.date_combo.currentText()
        if not date_column:
            raise ValueError("Select a date column for the pivot X-axis.")
        layout = self._selected_layout()
        summary = self.summary_check.isChecked() and layout != "per_sheet"
        level = self._current_level()
        if summary:
            if self.comparison_combo.currentText().startswith("Custom"):
                period_a = self._build_custom_period("pre")
                period_b = self._build_custom_period("post")
            else:
                period_a, period_b = self._auto_periods()
            if not period_a or not period_b:
                raise ValueError("No usable dates are available for the period comparison.")
        else:
            period_a = period_b = ""
        group = self.group_combo.currentText()
        time_col = self.time_combo.currentText()
        aggregation_label = self.aggregation_combo.currentText() or "Average"
        return {
            "frame": self.frame,
            "path": self.source_path,
            "date_col": date_column,
            "time_col": None if time_col == "None" else time_col,
            "group_col": None if group == "None" else group,
            "slicers": self.slicer_panel.selected_items(),
            "kpis": kpis,
            "agg": AGG_OPTIONS[aggregation_label],
            "agg_label": aggregation_label,
            "chart": CHART_OPTIONS[self.chart_combo.currentText() or "Column"],
            "out_dir": self.output_edit.text().strip() or str(Path(self.source_path).parent),
            "layout": layout,
            "summary": summary,
            "cmp_level": level,
            "period_a": period_a,
            "period_b": period_b,
        }

    def generate_report(self):
        try:
            params = self._read_pivot_params()
        except ValueError as exc:
            self._show_error("Report setup incomplete", str(exc))
            return
        self._set_busy(True, "Preparing data…")
        task = BackgroundTask(lambda progress: run_pivot_pipeline(params, progress), emits_progress=True)
        task.signals.progress.connect(self._pipeline_progress)
        task.signals.succeeded.connect(self._report_saved)
        task.signals.failed.connect(self._report_failed)
        self._tasks.append(task)
        self.pool.start(task)

    @Slot(tuple)
    def _pipeline_progress(self, event: tuple):
        message, percent = (event + (None,))[:2] if isinstance(event, tuple) else (str(event), None)
        self.status_title.setText("Generating report…")
        self.status_message.setText(f"{message} ({percent}%)" if percent is not None else message)
        if percent is not None:
            if self.progress.maximum() == 1:
                self.progress.setRange(0, 100)
            self.progress.setValue(int(percent))
            self.progress_percent.setText(f"{int(percent)}%")
            self.progress_percent.show()

    @Slot(object)
    def _report_saved(self, result):
        self._tasks.clear()
        self._set_busy(False)
        path = result["xlsx"]
        self.status_title.setText("Report ready.")
        self.status_message.setText(Path(path).name)
        QMessageBox.information(self, "Export complete", f"Pivot report created:\n\n{path}")

    @Slot(str)
    def _report_failed(self, message: str):
        self._tasks.clear()
        self._set_busy(False)
        self._show_error("Could not generate the report", message)

    def _set_busy(self, busy: bool, text: str = ""):
        self.busy = busy
        self.generate_button.setEnabled(not busy)
        self.generate_button.setText("Generating…" if busy else "Generate Excel Report")
        if busy and text:
            self.status_message.setText(text)
        if busy:
            self.progress.setRange(0, 100)
            self.progress.setValue(0)
            self.progress_percent.setText("0%")
            self.progress.show()
            self.progress_percent.show()
        else:
            self.progress.setRange(0, 1)
            self.progress.hide()
            self.progress_percent.hide()

    def _show_error(self, title: str, message: str):
        QMessageBox.warning(self, title, message)
