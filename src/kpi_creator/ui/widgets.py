from __future__ import annotations

import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, Signal
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QSizePolicy,
    QStyle,
    QTableView,
    QVBoxLayout,
    QWidget,
)


class DataFrameModel(QAbstractTableModel):
    def __init__(self, frame: pd.DataFrame | None = None, parent=None):
        super().__init__(parent)
        self.frame = frame if frame is not None else pd.DataFrame()

    def set_frame(self, frame: pd.DataFrame):
        self.beginResetModel()
        self.frame = frame if frame is not None else pd.DataFrame()
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.frame.index)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.frame.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        value = self.frame.iat[index.row(), index.column()]
        if role == Qt.ItemDataRole.DisplayRole:
            if pd.isna(value):
                return "—"
            if isinstance(value, float):
                return f"{value:,.3f}".rstrip("0").rstrip(".")
            return str(value)
        if role == Qt.ItemDataRole.BackgroundRole and index.row() % 2:
            return QBrush(QColor("#F8FAFC"))
        if role == Qt.ItemDataRole.TextAlignmentRole and isinstance(value, (int, float)):
            return int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal and section < len(self.frame.columns):
            return str(self.frame.columns[section])
        if orientation == Qt.Orientation.Vertical:
            return str(section + 1)
        return None


class CheckableListWidget(QListWidget):
    """List where you click or drag with the left button to toggle rows' checkboxes."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._drag_state = None
        self._last_item = None

    def mousePressEvent(self, event):
        if event.button() != Qt.MouseButton.LeftButton:
            super().mousePressEvent(event)
            return
        pos = event.position().toPoint()
        item = self.itemAt(pos)
        if item is None or not (item.flags() & Qt.ItemFlag.ItemIsUserCheckable):
            self._drag_state = None
            super().mousePressEvent(event)
            return
        rect = self.visualItemRect(item)
        indicator = self.style().pixelMetric(QStyle.PixelMetric.PM_IndicatorWidth, None, self)
        if pos.x() <= rect.left() + indicator + 6:
            super().mousePressEvent(event)
            self._drag_state = item.checkState()
        else:
            self._drag_state = (
                Qt.CheckState.Unchecked if item.checkState() == Qt.CheckState.Checked else Qt.CheckState.Checked
            )
            item.setCheckState(self._drag_state)
        self.setCurrentItem(item)
        self._last_item = item
        event.accept()

    def mouseMoveEvent(self, event):
        if self._drag_state is None or not (event.buttons() & Qt.MouseButton.LeftButton):
            super().mouseMoveEvent(event)
            return
        item = self.itemAt(event.position().toPoint())
        if item is not None and item is not self._last_item and (item.flags() & Qt.ItemFlag.ItemIsUserCheckable):
            item.setCheckState(self._drag_state)
            self._last_item = item
        event.accept()

    def mouseReleaseEvent(self, event):
        self._drag_state = None
        self._last_item = None
        super().mouseReleaseEvent(event)


class MultiSelectPanel(QWidget):
    selectionChanged = Signal()

    def __init__(self, title: str, description: str = "", parent=None):
        super().__init__(parent)
        self._all_items: list[str] = []
        self.title = title
        self.count_label = QLabel("0 selected")
        self.count_label.setObjectName("countBadge")
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search columns")
        self.search.textChanged.connect(self._filter_items)
        self.list = CheckableListWidget()
        self.list.setAlternatingRowColors(False)
        self.list.itemChanged.connect(self._on_item_changed)
        self.empty_label = QLabel("Choose a workbook to see available columns.")
        self.empty_label.setObjectName("emptyMessage")
        self.empty_label.setWordWrap(True)

        header = QHBoxLayout()
        title_label = QLabel(title)
        title_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        header.addWidget(title_label, 1)
        header.addWidget(self.count_label)
        toolbar = QHBoxLayout()
        toolbar.addWidget(QLabel("COLUMN NAME"), 1)
        select_all = QPushButton("Select All")
        clear = QPushButton("Clear")
        select_all.setObjectName("textButton")
        clear.setObjectName("textButton")
        select_all.clicked.connect(self.select_all)
        clear.clicked.connect(self.clear)
        toolbar.addWidget(select_all)
        toolbar.addWidget(clear)
        self.list.setMinimumHeight(120)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)
        layout.addLayout(header)
        line_height = self.fontMetrics().lineSpacing()
        if description:
            label = QLabel(description)
            label.setObjectName("muted")
            label.setWordWrap(True)
            label.setMinimumHeight(line_height)
            layout.addWidget(label)
        layout.addWidget(self.search)
        box = QFrame()
        box.setObjectName("listBox")
        box_layout = QVBoxLayout(box)
        box_layout.setContentsMargins(4, 4, 4, 4)
        box_layout.setSpacing(4)
        box_layout.addLayout(toolbar)
        box_layout.addWidget(self.list, 1)
        box_layout.addWidget(self.empty_label, 0, Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(box, 1)
        self._refresh_empty()

    def set_items(self, items: list[str], selected: list[str] | None = None):
        selected = set(selected or [])
        self._all_items = list(items)
        self.list.blockSignals(True)
        self.list.clear()
        for text in items:
            item = QListWidgetItem(text)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if text in selected else Qt.CheckState.Unchecked)
            self.list.addItem(item)
        self.list.blockSignals(False)
        self._filter_items(self.search.text())
        self._update_count()
        self._refresh_empty()

    def selected_items(self) -> list[str]:
        return [self.list.item(i).text() for i in range(self.list.count()) if self.list.item(i).checkState() == Qt.CheckState.Checked]

    def select_all(self):
        self.list.blockSignals(True)
        for i in range(self.list.count()):
            if not self.list.item(i).isHidden():
                self.list.item(i).setCheckState(Qt.CheckState.Checked)
        self.list.blockSignals(False)
        self._update_count()
        self.selectionChanged.emit()

    def clear(self):
        self.list.blockSignals(True)
        for i in range(self.list.count()):
            self.list.item(i).setCheckState(Qt.CheckState.Unchecked)
        self.list.blockSignals(False)
        self._update_count()
        self.selectionChanged.emit()

    def _filter_items(self, value: str):
        query = value.casefold().strip()
        for i in range(self.list.count()):
            self.list.item(i).setHidden(query not in self.list.item(i).text().casefold())
        self._refresh_empty()

    def _on_item_changed(self, _item):
        self._update_count()
        self.selectionChanged.emit()

    def _update_count(self):
        self.count_label.setText(f"{len(self.selected_items())} selected")

    def _refresh_empty(self):
        visible = any(not self.list.item(i).isHidden() for i in range(self.list.count()))
        self.empty_label.setVisible(bool(self._all_items) and not visible)
        self.empty_label.setText("No matching columns.")


def style_table(table: QTableView, model: DataFrameModel):
    table.setModel(model)
    table.setAlternatingRowColors(True)
    table.setSortingEnabled(False)
    table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
    table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
    table.horizontalHeader().setStretchLastSection(True)
    table.horizontalHeader().setMinimumSectionSize(90)
    table.verticalHeader().setVisible(False)
    table.setShowGrid(False)
