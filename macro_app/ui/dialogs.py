from __future__ import annotations

from typing import Any

from PySide6.QtWidgets import QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFormLayout, QLineEdit, QVBoxLayout, QWidget

from ..constants import ACTION_DEFINITIONS
from ..models import get_default_delay

WINDOW_MULTI_MODE_OPTIONS: list[tuple[str, str]] = [
    ("第一个窗口", "first"),
    ("全部并行", "sync"),
    ("逐个执行", "serial"),
]


def edit_step_dialog(parent: QWidget, step: dict[str, Any]) -> dict[str, Any] | None:
    dialog = QDialog(parent)
    dialog.setWindowTitle("编辑小节点")
    layout = QVBoxLayout(dialog)
    form = QFormLayout()

    alias_edit = QLineEdit(step.get("alias", ""))
    delay_edit = QLineEdit(step.get("delay", get_default_delay()))
    type_combo = QComboBox()
    type_combo.addItems(list(ACTION_DEFINITIONS.keys()))
    type_combo.setCurrentText(step["action"])

    form.addRow("别名", alias_edit)
    form.addRow("延迟(s)", delay_edit)
    form.addRow("类型", type_combo)

    field_edits: dict[str, QLineEdit] = {}

    def rebuild_fields(action_name: str):
        while form.rowCount() > 3:
            form.removeRow(3)
        field_edits.clear()
        params = step["params"] if action_name == step["action"] else {}
        for field, label, default in ACTION_DEFINITIONS[action_name]:
            edit = QLineEdit(params.get(field, default))
            field_edits[field] = edit
            form.addRow(label, edit)

    rebuild_fields(step["action"])
    type_combo.currentTextChanged.connect(rebuild_fields)

    layout.addLayout(form)
    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)

    if dialog.exec() != QDialog.Accepted:
        return None

    return {
        "alias": alias_edit.text().strip(),
        "delay": delay_edit.text().strip() or get_default_delay(),
        "action": type_combo.currentText(),
        "params": {field: edit.text() for field, edit in field_edits.items()},
    }


def edit_big_node_dialog(parent: QWidget, node: dict[str, Any]) -> dict[str, Any] | None:
    scope = node.get("scope", {})

    dialog = QDialog(parent)
    dialog.setWindowTitle("编辑大节点")
    layout = QVBoxLayout(dialog)
    form = QFormLayout()

    alias_edit = QLineEdit(node.get("alias") or "")
    regex_edit = QLineEdit(scope.get("regex", ""))
    bring_front_check = QCheckBox("执行前置顶")
    bring_front_check.setChecked(bool(scope.get("bring_front", True)))
    mode_combo = QComboBox()
    for label, value in WINDOW_MULTI_MODE_OPTIONS:
        mode_combo.addItem(label, value)

    current_mode = str(scope.get("multi_window_mode", "first") or "first").strip().lower()
    mode_index = next((idx for idx, (_label, value) in enumerate(WINDOW_MULTI_MODE_OPTIONS) if value == current_mode), 0)
    mode_combo.setCurrentIndex(mode_index)

    form.addRow("别名", alias_edit)
    form.addRow("作用域正则", regex_edit)
    form.addRow("多窗口模式", mode_combo)
    form.addRow(bring_front_check)
    layout.addLayout(form)

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)

    if dialog.exec() != QDialog.Accepted:
        return None

    return {
        "alias": alias_edit.text().strip(),
        "scope": {
            "regex": regex_edit.text().strip(),
            "bring_front": bring_front_check.isChecked(),
            "multi_window_mode": mode_combo.currentData(),
        },
    }
