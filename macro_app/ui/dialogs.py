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

NODE_FLOW_MODE_OPTIONS: list[tuple[str, str]] = [
    ("顺序到下一个", "next"),
    ("循环当前大节点", "loop"),
    ("无条件跳转", "jump"),
    ("条件跳转", "conditional_jump"),
    ("运行完此大节点停止", "stop"),
]

NODE_FLOW_CONDITION_OPTIONS: list[tuple[str, str]] = [
    ("检测到时跳转", "last_detected"),
    ("未检测到时跳转", "last_not_detected"),
    ("窗口出现文字时跳转", "window_text_detected"),
    ("窗口匹配图片时跳转", "window_image_detected"),
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
    flow = node.get("flow", {})

    dialog = QDialog(parent)
    dialog.setWindowTitle("编辑大节点")
    layout = QVBoxLayout(dialog)
    form = QFormLayout()

    alias_edit = QLineEdit(node.get("alias") or "")
    node_delay_edit = QLineEdit(str(node.get("node_delay", "0") or "0"))
    node_delay_edit.setPlaceholderText("切换到下一个大节点前等待秒数，默认 0")
    regex_edit = QLineEdit(scope.get("regex", ""))
    bring_front_check = QCheckBox("执行前置顶")
    bring_front_check.setChecked(bool(scope.get("bring_front", True)))
    mode_combo = QComboBox()
    for label, value in WINDOW_MULTI_MODE_OPTIONS:
        mode_combo.addItem(label, value)

    current_mode = str(scope.get("multi_window_mode", "first") or "first").strip().lower()
    mode_index = next((idx for idx, (_label, value) in enumerate(WINDOW_MULTI_MODE_OPTIONS) if value == current_mode), 0)
    mode_combo.setCurrentIndex(mode_index)

    flow_mode_combo = QComboBox()
    for label, value in NODE_FLOW_MODE_OPTIONS:
        flow_mode_combo.addItem(label, value)
    current_flow_mode = str(flow.get("mode", "next") or "next").strip().lower()
    flow_mode_index = next((idx for idx, (_label, value) in enumerate(NODE_FLOW_MODE_OPTIONS) if value == current_flow_mode), 0)
    flow_mode_combo.setCurrentIndex(flow_mode_index)

    flow_target_edit = QLineEdit(str(flow.get("target", "") or ""))
    flow_target_edit.setPlaceholderText("跳转到大节点编号，例如 2")
    flow_condition_combo = QComboBox()
    for label, value in NODE_FLOW_CONDITION_OPTIONS:
        flow_condition_combo.addItem(label, value)
    current_condition = str(flow.get("condition", "last_detected") or "last_detected").strip().lower()
    condition_index = next(
        (idx for idx, (_label, value) in enumerate(NODE_FLOW_CONDITION_OPTIONS) if value == current_condition),
        0,
    )
    flow_condition_combo.setCurrentIndex(condition_index)
    flow_max_loops_edit = QLineEdit(str(flow.get("max_loops", "1") or "1"))
    flow_max_loops_edit.setPlaceholderText("循环次数，默认 1")
    flow_text_edit = QLineEdit(str(flow.get("condition_target_text", "") or ""))
    flow_text_edit.setPlaceholderText("例如：登录成功")
    flow_text_match_mode_combo = QComboBox()
    for label, value in [("包含", "contains"), ("完全匹配", "exact"), ("正则", "regex")]:
        flow_text_match_mode_combo.addItem(label, value)
    current_text_match_mode = str(flow.get("condition_text_match_mode", "contains") or "contains").strip().lower()
    text_match_mode_index = next(
        (idx for idx in range(flow_text_match_mode_combo.count()) if flow_text_match_mode_combo.itemData(idx) == current_text_match_mode),
        0,
    )
    flow_text_match_mode_combo.setCurrentIndex(text_match_mode_index)
    flow_min_score_edit = QLineEdit(str(flow.get("condition_min_score", "0.5") or "0.5"))
    flow_min_score_edit.setPlaceholderText("0~1")
    flow_image_path_edit = QLineEdit(str(flow.get("condition_image_path", "") or ""))
    flow_image_path_edit.setPlaceholderText("模板图片路径")
    flow_image_match_mode_combo = QComboBox()
    for label, value in [("默认", "default"), ("二值化", "binary"), ("轮廓", "contour")]:
        flow_image_match_mode_combo.addItem(label, value)
    current_image_match_mode = str(flow.get("condition_match_mode", "default") or "default").strip().lower()
    image_match_mode_index = next(
        (idx for idx in range(flow_image_match_mode_combo.count()) if flow_image_match_mode_combo.itemData(idx) == current_image_match_mode),
        0,
    )
    flow_image_match_mode_combo.setCurrentIndex(image_match_mode_index)
    flow_threshold_edit = QLineEdit(str(flow.get("condition_threshold", "0.8") or "0.8"))
    flow_threshold_edit.setPlaceholderText("0~1")

    form.addRow("别名", alias_edit)
    form.addRow("大节点间延迟(s)", node_delay_edit)
    form.addRow("作用域正则", regex_edit)
    form.addRow("多窗口模式", mode_combo)
    form.addRow("流转模式", flow_mode_combo)
    form.addRow("跳转目标", flow_target_edit)
    form.addRow("条件类型", flow_condition_combo)
    form.addRow("循环次数", flow_max_loops_edit)
    form.addRow("条件文字", flow_text_edit)
    form.addRow("文字匹配", flow_text_match_mode_combo)
    form.addRow("最低置信度", flow_min_score_edit)
    form.addRow("条件图片", flow_image_path_edit)
    form.addRow("图片匹配", flow_image_match_mode_combo)
    form.addRow("图片阈值", flow_threshold_edit)
    form.addRow(bring_front_check)
    layout.addLayout(form)

    def update_flow_fields() -> None:
        mode = str(flow_mode_combo.currentData() or "next")
        condition = str(flow_condition_combo.currentData() or "last_detected")
        flow_target_edit.setEnabled(mode in {"jump", "conditional_jump"})
        flow_condition_combo.setEnabled(mode == "conditional_jump")
        flow_max_loops_edit.setEnabled(mode == "loop")
        text_enabled = mode == "conditional_jump" and condition == "window_text_detected"
        image_enabled = mode == "conditional_jump" and condition == "window_image_detected"
        flow_text_edit.setEnabled(text_enabled)
        flow_text_match_mode_combo.setEnabled(text_enabled)
        flow_min_score_edit.setEnabled(text_enabled)
        flow_image_path_edit.setEnabled(image_enabled)
        flow_image_match_mode_combo.setEnabled(image_enabled)
        flow_threshold_edit.setEnabled(image_enabled)

    flow_mode_combo.currentIndexChanged.connect(lambda _index: update_flow_fields())
    flow_condition_combo.currentIndexChanged.connect(lambda _index: update_flow_fields())
    update_flow_fields()

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)

    if dialog.exec() != QDialog.Accepted:
        return None

    return {
        "alias": alias_edit.text().strip(),
        "node_delay": node_delay_edit.text().strip() or "0",
        "scope": {
            "regex": regex_edit.text().strip(),
            "bring_front": bring_front_check.isChecked(),
            "multi_window_mode": mode_combo.currentData(),
        },
        "flow": {
            "mode": flow_mode_combo.currentData(),
            "target": flow_target_edit.text().strip(),
            "condition": flow_condition_combo.currentData(),
            "max_loops": flow_max_loops_edit.text().strip() or "1",
            "condition_target_text": flow_text_edit.text().strip(),
            "condition_text_match_mode": flow_text_match_mode_combo.currentData(),
            "condition_min_score": flow_min_score_edit.text().strip() or "0.5",
            "condition_image_path": flow_image_path_edit.text().strip(),
            "condition_match_mode": flow_image_match_mode_combo.currentData(),
            "condition_threshold": flow_threshold_edit.text().strip() or "0.8",
        },
    }
