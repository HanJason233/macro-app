from __future__ import annotations

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QGridLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..constants import ACTION_GROUPS

ACTION_EMOJIS: dict[str, str] = {
    "键盘按键": "⌨️",
    "输入文本": "📝",
    "绝对坐标点击": "🖱️",
    "窗口相对点击": "🎯",
    "识图点击": "🧩",
    "OCR文字点击": "🔤",
    "检测图片": "🔎",
    "检测文字": "🔠",
    "检测窗口分辨率": "📐",
    "获取窗口分辨率": "📏",
    "启动程序": "🚀",
    "关闭程序": "🛑",
    "窗口大小": "🪟",
    "最小化窗口": "📉",
}


class ActionPalette(QWidget):
    action_clicked = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        for group_name, action_names in ACTION_GROUPS.items():
            group_label = QLabel(group_name)
            group_label.setStyleSheet("font-weight: 700;")
            layout.addWidget(group_label)

            grid = QGridLayout()
            grid.setContentsMargins(0, 0, 0, 0)
            grid.setHorizontalSpacing(8)
            grid.setVerticalSpacing(8)

            for index, action_name in enumerate(action_names):
                row = index // 2
                col = index % 2
                button = QPushButton(f"{ACTION_EMOJIS.get(action_name, '⚙️')} {action_name}")
                button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                button.clicked.connect(lambda _=False, name=action_name: self.action_clicked.emit(name))
                grid.addWidget(button, row, col)

            layout.addLayout(grid)

        layout.addStretch(1)


class StepListWidget(QTreeWidget):
    action_dropped = Signal(str, int)
    row_changed = Signal(int)
    step_double_clicked = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderLabels(["操作类型", "备注", "延迟(s)", "参数详情"])
        self.setRootIsDecorated(False)
        self.setAlternatingRowColors(True)
        self.setAcceptDrops(True)
        self.setDragDropMode(QTreeWidget.NoDragDrop)
        self.itemSelectionChanged.connect(self._emit_row_changed)
        self.itemDoubleClicked.connect(self._emit_double_clicked)

    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasText():
            index = self.indexOfTopLevelItem(self.itemAt(event.position().toPoint()))
            if index < 0:
                index = self.topLevelItemCount()
            self.action_dropped.emit(event.mimeData().text(), index)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)

    def _emit_row_changed(self):
        item = self.currentItem()
        if item is None:
            return
        self.row_changed.emit(self.indexOfTopLevelItem(item))

    def _emit_double_clicked(self, item, _column):
        self.step_double_clicked.emit(self.indexOfTopLevelItem(item))


class BigNodeListWidget(QTreeWidget):
    row_changed = Signal(int)
    node_double_clicked = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderLabels(["别名", "作用域正则", "多窗口模式"])
        self.setRootIsDecorated(False)
        self.setAlternatingRowColors(True)
        self.itemSelectionChanged.connect(self._emit_row_changed)
        self.itemDoubleClicked.connect(self._emit_double_clicked)

    def _emit_row_changed(self):
        item = self.currentItem()
        if item is None:
            return
        self.row_changed.emit(self.indexOfTopLevelItem(item))

    def _emit_double_clicked(self, item, _column):
        self.node_double_clicked.emit(self.indexOfTopLevelItem(item))
