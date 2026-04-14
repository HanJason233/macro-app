import base64
import html
import json
import mimetypes
import re
import time
from hashlib import sha256
from pathlib import Path
from uuid import uuid4

from PySide6.QtCore import Qt, QThread, QTimer
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QSplitter,
    QTextEdit,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..constants import (
    ACTION_CLICK,
    ACTION_CLICK_IMAGE,
    ACTION_CLICK_OCR,
    ACTION_CLOSE_PROGRAM,
    ACTION_DETECT_IMAGE,
    ACTION_DETECT_OCR,
    ACTION_DETECT_WINDOW_SIZE,
    ACTION_DEFINITIONS,
    ACTION_GET_WINDOW_SIZE,
    ACTION_MOUSE_DRAG,
    ACTION_MOUSE_SCROLL,
    ACTION_MINIMIZE_WINDOW,
    ACTION_RESIZE_WINDOW,
)
from ..models import (
    clone_payload,
    create_big_node,
    create_small_node,
    create_workflow,
    get_default_delay,
    get_big_nodes,
    get_small_nodes,
    normalize_big_node,
    normalize_workflow,
    normalize_small_node,
    set_default_delay,
)
from ..services.runner import WorkflowRunner
from ..services.windows import (
    enumerate_windows,
    get_window_frame_rect,
    resolve_scope_window,
    resolve_scope_windows,
)
from .dialogs import edit_big_node_dialog, edit_step_dialog
from .overlays import PointPickerOverlay, ScreenshotOverlay
from .panels import ActionPalette, BigNodeListWidget, StepListWidget


def _build_test_actions(detection_actions: list[str]) -> list[str]:
    detection_set = set(detection_actions)
    return [action for action in ACTION_DEFINITIONS.keys() if action not in detection_set]


class MainWindow(QMainWindow):
    SESSION_STATE_FILE = Path.cwd() / ".macro_app_state.json"
    WORKFLOW_DIR = Path.cwd() / "workflows"
    EMBEDDED_IMAGE_DIR = Path.cwd() / ".macro_app_embedded_images"
    EMBEDDED_IMAGE_MAX_BYTES = 3 * 1024 * 1024
    IMAGE_STORAGE_FILE = "file"
    IMAGE_STORAGE_BASE64_AUTO = "base64_auto"
    IMAGE_STORAGE_BASE64_ALWAYS = "base64_always"
    DETECTION_ACTIONS = [
        ACTION_DETECT_IMAGE,
        ACTION_DETECT_OCR,
        ACTION_DETECT_WINDOW_SIZE,
        ACTION_GET_WINDOW_SIZE,
    ]
    TEST_ACTIONS = _build_test_actions(DETECTION_ACTIONS)
    WINDOW_REQUIRED_TEST_ACTIONS = {
        ACTION_CLICK,
        ACTION_CLICK_IMAGE,
        ACTION_CLICK_OCR,
        ACTION_MOUSE_DRAG,
        ACTION_MOUSE_SCROLL,
        ACTION_DETECT_IMAGE,
        ACTION_DETECT_OCR,
        ACTION_DETECT_WINDOW_SIZE,
        ACTION_GET_WINDOW_SIZE,
        ACTION_CLOSE_PROGRAM,
        ACTION_RESIZE_WINDOW,
        ACTION_MINIMIZE_WINDOW,
    }
    LOG_COLORS = {
        "default": "#C8CCD4",
        "running": "#6EA8FE",
        "success": "#7AD97A",
        "warning": "#F5C26B",
        "error": "#FF7B7B",
        "timestamp": "#98A2B3",
    }
    COORD_ACTIONS = {ACTION_CLICK, ACTION_MOUSE_DRAG, ACTION_MOUSE_SCROLL}

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Macro App")
        self.resize(1560, 920)

        self.windows = []
        self.workflows = [create_workflow()]
        self.workflow_files: list[Path | None] = [None]
        self.workflow_index = 0
        self.node_index = 0
        self.step_index = 0
        self.current_file = None
        self.selected_hwnd = None
        self.runner = None
        self.runner_thread = None
        self._runner_node_index_map: list[int] | None = None
        self._matched_window_hwnds: list[int] = []
        self._active_point_overlay = None
        self._active_screenshot_overlay = None
        self.node_view_mode = "list"
        self.default_delay = get_default_delay()
        self.step_view_mode = "list"
        self.image_storage_mode = self.IMAGE_STORAGE_BASE64_AUTO
        self.auto_refresh_windows_each_step = False
        self.WORKFLOW_DIR.mkdir(parents=True, exist_ok=True)
        self.EMBEDDED_IMAGE_DIR.mkdir(parents=True, exist_ok=True)

        self._build_ui()
        self._set_node_view_mode("list")
        self._set_step_view_mode("list")
        self._sync_global_settings_widgets()
        self.refresh_windows()
        self.refresh_all()
        self._auto_load_last_workflow()


    def _is_runner_thread_running(self) -> bool:
        if self.runner_thread is None:
            return False
        try:
            return self.runner_thread.isRunning()
        except RuntimeError:
            # Underlying C++ object was already deleted.
            self.runner_thread = None
            self.runner = None
            return False

    def _on_runner_thread_finished(self):
        self.runner_thread = None
        self.runner = None
        self._runner_node_index_map = None

    def _step_display_alias(self, index: int, step: dict) -> str:
        alias = (step.get("alias") or "").strip()
        return alias or f"小节点 {index + 1}"

    def _big_node_display_alias(self, index: int, node: dict) -> str:
        alias = (node.get("alias") or node.get("name") or "").strip()
        return alias or f"大节点 {index + 1}"
    @staticmethod
    def _multi_window_mode_label(mode: str) -> str:
        mode = (mode or "first").strip().lower()
        if mode == "sync":
            return "同步执行"
        if mode == "serial":
            return "串行执行"
        return "仅首个窗口"

    @staticmethod
    def _flow_mode_label(mode: str) -> str:
        mode = (mode or "next").strip().lower()
        if mode == "loop":
            return "循环"
        if mode == "jump":
            return "无条件跳转"
        if mode == "conditional_jump":
            return "条件跳转"
        if mode == "stop":
            return "执行后停止"
        return "顺序"

    def _flow_summary(self, flow: dict) -> str:
        mode = str(flow.get("mode", "next") or "next").strip().lower()
        target = str(flow.get("target", "") or "").strip()
        condition = str(flow.get("condition", "last_detected") or "last_detected").strip().lower()
        max_loops = str(flow.get("max_loops", "1") or "1").strip()
        if mode == "loop":
            return f"循环 x{max_loops or '1'}"
        if mode == "jump":
            return f"跳转 -> {target or '?'}"
        if mode == "conditional_jump":
            if condition == "last_detected":
                cond_text = "检测到"
            elif condition == "last_not_detected":
                cond_text = "未检测到"
            elif condition == "window_text_detected":
                cond_text = "窗口文字命中"
            elif condition == "window_image_detected":
                cond_text = "窗口图片命中"
            else:
                cond_text = condition or "条件"
            return f"{cond_text} -> {target or '?'}"
        if mode == "stop":
            return "执行后停止"
        return "顺序 -> 下一个"

    def _workflow_flowchart_summary(self) -> str:
        big_nodes = get_big_nodes(self.workflow)
        lines: list[str] = [f"工作流流程图：{self.workflow.get('name', '工作流')}", ""]
        for index, node in enumerate(big_nodes, start=1):
            alias = self._big_node_display_alias(index - 1, node)
            flow = node.get("flow", {})
            mode = str(flow.get("mode", "next") or "next").strip().lower()
            target = str(flow.get("target", "") or "").strip()
            condition = str(flow.get("condition", "last_detected") or "last_detected").strip().lower()
            node_delay = str(node.get("node_delay", "0") or "0").strip()
            lines.append(f"[{index}] {alias}")
            if mode == "loop":
                loops = str(flow.get("max_loops", "1") or "1").strip()
                lines.append(f"  ├─ 循环 x{loops or '1'}，结束后 -> [{index + 1}]")
            elif mode == "jump":
                lines.append(f"  ├─ 无条件跳转 -> [{target or '?'}]")
            elif mode == "conditional_jump":
                if condition == "last_detected":
                    cond_text = "检测到"
                elif condition == "last_not_detected":
                    cond_text = "未检测到"
                elif condition == "window_text_detected":
                    cond_text = "窗口文字命中"
                elif condition == "window_image_detected":
                    cond_text = "窗口图片命中"
                else:
                    cond_text = condition or "条件"
                lines.append(f"  ├─ 条件({cond_text}) 满足 -> [{target or '?'}]")
                lines.append(f"  └─ 条件不满足 -> [{index + 1}]")
            elif mode == "stop":
                lines.append("  └─ 执行完毕后停止工作流")
            else:
                lines.append(f"  └─ 顺序 -> [{index + 1}]")
            lines.append(f"  节点间延迟: {node_delay or '0'}s")
            small_nodes = get_small_nodes(node)
            if small_nodes:
                lines.append("  小节点链路:")
                for step_idx, step in enumerate(small_nodes, start=1):
                    step_alias = self._step_display_alias(step_idx - 1, step)
                    action = str(step.get("action", ""))
                    if step_idx < len(small_nodes):
                        lines.append(f"    [{step_idx}] {step_alias} ({action}) -> [{step_idx + 1}]")
                    else:
                        lines.append(f"    [{step_idx}] {step_alias} ({action}) -> 结束")
            lines.append("")
        return "\n".join(lines).strip()

    @property
    def workflow(self):
        return self.workflows[self.workflow_index]

    @property
    def node(self):
        return get_big_nodes(self.workflow)[self.node_index]

    @property
    def step(self):
        return get_small_nodes(self.node)[self.step_index]

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        self.setMenuWidget(self._build_top_title_bar())

        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter, 1)

        splitter.addWidget(self._build_left_panel())
        splitter.addWidget(self._build_center_panel())
        splitter.addWidget(self._build_right_panel())
        splitter.setSizes([110, 870, 400])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 0)

    def _build_top_title_bar(self) -> QWidget:
        bar = QWidget(self)
        bar.setStyleSheet("background: #f7f8fa; border-bottom: 1px solid #d0d5dd;")
        row = QHBoxLayout(bar)
        row.setContentsMargins(10, 6, 10, 6)
        row.setSpacing(8)

        title = QLabel("工作流控制台")
        title.setStyleSheet("font-size: 14px; font-weight: 700;")
        row.addWidget(title)
        row.addSpacing(12)
        row.addWidget(QLabel("图片存储:"))
        self.image_storage_combo = QComboBox()
        self.image_storage_combo.addItem("图片文件", self.IMAGE_STORAGE_FILE)
        self.image_storage_combo.addItem("自动Base64(<3MB)", self.IMAGE_STORAGE_BASE64_AUTO)
        self.image_storage_combo.addItem("强制Base64", self.IMAGE_STORAGE_BASE64_ALWAYS)
        self.image_storage_combo.setMinimumWidth(170)
        self.image_storage_combo.currentIndexChanged.connect(self._on_image_storage_mode_changed)
        row.addWidget(self.image_storage_combo)
        self.auto_refresh_step_checkbox = QCheckBox("小节点前刷新窗口")
        self.auto_refresh_step_checkbox.setChecked(self.auto_refresh_windows_each_step)
        self.auto_refresh_step_checkbox.toggled.connect(self._on_auto_refresh_step_toggled)
        row.addWidget(self.auto_refresh_step_checkbox)

        row.addWidget(QLabel("动作测试:"))
        self.test_action_combo = QComboBox()
        for action_name in self.TEST_ACTIONS:
            self.test_action_combo.addItem(action_name, action_name)
        self.test_action_combo.setMinimumWidth(150)
        row.addWidget(self.test_action_combo)
        test_button = QPushButton("执行动作")
        test_button.clicked.connect(self.test_action)
        row.addWidget(test_button)
        row.addWidget(QLabel("检测测试:"))
        self.detect_action_combo = QComboBox()
        for action_name in self.DETECTION_ACTIONS:
            self.detect_action_combo.addItem(action_name, action_name)
        self.detect_action_combo.setMinimumWidth(150)
        row.addWidget(self.detect_action_combo)
        detect_test_button = QPushButton("执行检测")
        detect_test_button.clicked.connect(self.test_detection_action)
        row.addWidget(detect_test_button)
        row.addWidget(QLabel("测试窗口正则:"))
        self.test_scope_regex_edit = QLineEdit()
        self.test_scope_regex_edit.setPlaceholderText("例如: Notepad|记事本")
        self.test_scope_regex_edit.setMinimumWidth(180)
        row.addWidget(self.test_scope_regex_edit)
        row.addStretch(1)

        self.status_label = QLabel("准备就绪")
        self.status_label.setStyleSheet("color: #667085;")
        row.addWidget(self.status_label)
        return bar

    def _on_image_storage_mode_changed(self, _index: int) -> None:
        mode = self.image_storage_combo.currentData()
        if not isinstance(mode, str):
            return
        self.image_storage_mode = mode
        self._save_session_state()
        self.log(f"图片存储模式已切换为：{self.image_storage_combo.currentText()}")

    def _on_auto_refresh_step_toggled(self, checked: bool) -> None:
        self.auto_refresh_windows_each_step = bool(checked)
        self._save_session_state()
        if checked:
            self.log("已开启：小节点执行前自动刷新窗口")
        else:
            self.log("已关闭：小节点执行前自动刷新窗口")

    def _sync_global_settings_widgets(self) -> None:
        if hasattr(self, "image_storage_combo"):
            index = self.image_storage_combo.findData(self.image_storage_mode)
            if index >= 0:
                self.image_storage_combo.blockSignals(True)
                self.image_storage_combo.setCurrentIndex(index)
                self.image_storage_combo.blockSignals(False)
        if hasattr(self, "auto_refresh_step_checkbox"):
            self.auto_refresh_step_checkbox.blockSignals(True)
            self.auto_refresh_step_checkbox.setChecked(self.auto_refresh_windows_each_step)
            self.auto_refresh_step_checkbox.blockSignals(False)

    def _build_left_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        workflow_box = self._create_section("工作流")
        workflow_layout = workflow_box.layout()
        self.workflow_list = QListWidget()
        self.workflow_list.currentRowChanged.connect(self.on_workflow_selected)
        workflow_layout.addWidget(self.workflow_list)
        workflow_row1 = QHBoxLayout()
        for text, callback in [
            ("新建", self.new_workflow),
            ("打开", self.open_file),
        ]:
            button = QPushButton(text)
            button.clicked.connect(callback)
            workflow_row1.addWidget(button)
        workflow_row1.addStretch(1)
        workflow_layout.addLayout(workflow_row1)

        workflow_row1b = QHBoxLayout()
        for text, callback in [
            ("保存", self.save_file),
            ("另存为", self.save_file_as),
        ]:
            button = QPushButton(text)
            button.clicked.connect(callback)
            workflow_row1b.addWidget(button)
        workflow_row1b.addStretch(1)
        workflow_layout.addLayout(workflow_row1b)

        workflow_row2 = QHBoxLayout()
        delete_button = QPushButton("删除")
        delete_button.clicked.connect(self.delete_workflow)
        workflow_row2.addWidget(delete_button)
        workflow_row2.addStretch(1)
        workflow_layout.addLayout(workflow_row2)
        layout.addWidget(workflow_box, 1)
        return panel

    def _build_center_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        top_split = QSplitter(Qt.Horizontal)
        layout.addWidget(top_split, 4)

        left_col = QWidget()
        left_col_layout = QVBoxLayout(left_col)
        left_col_layout.setContentsMargins(0, 0, 0, 0)
        left_col_layout.setSpacing(8)

        node_box = self._create_section("大节点")
        node_layout = node_box.layout()
        node_header = QHBoxLayout()
        self.node_view_list_button = QPushButton("切换为列表示图")
        self.node_view_list_button.clicked.connect(self.switch_node_view_to_list)
        node_header.addWidget(self.node_view_list_button)
        self.node_view_code_button = QPushButton("切换为代码视图")
        self.node_view_code_button.clicked.connect(self.switch_node_view_to_code)
        node_header.addWidget(self.node_view_code_button)
        self.node_view_flow_button = QPushButton("切换为流程图查看")
        self.node_view_flow_button.clicked.connect(self.switch_node_view_to_flowchart)
        node_header.addWidget(self.node_view_flow_button)
        node_header.addStretch(1)
        node_layout.addLayout(node_header)
        self.node_list = BigNodeListWidget()
        self.node_list.row_changed.connect(self.on_node_selected)
        self.node_list.node_double_clicked.connect(self.edit_big_node_quickly)
        self.node_code_editor = QTextEdit()
        self.node_code_editor.setPlaceholderText("请输入大节点 JSON 数组")
        self.node_flowchart_view = QTextEdit()
        self.node_flowchart_view.setReadOnly(True)
        self.node_view_stack = QStackedWidget()
        self.node_view_stack.addWidget(self.node_list)
        self.node_view_stack.addWidget(self.node_code_editor)
        self.node_view_stack.addWidget(self.node_flowchart_view)
        node_layout.addWidget(self.node_view_stack)
        node_row = QHBoxLayout()
        self.node_operation_buttons: list[QPushButton] = []
        for text, callback in [("新增", self.add_node), ("删除", self.delete_node), ("上移", self.move_node_up), ("下移", self.move_node_down)]:
            button = QPushButton(text)
            button.clicked.connect(callback)
            node_row.addWidget(button)
            self.node_operation_buttons.append(button)
        node_row.addStretch(1)
        node_layout.addLayout(node_row)
        left_col_layout.addWidget(node_box, 3)

        right_col = QWidget()
        right_col_layout = QVBoxLayout(right_col)
        right_col_layout.setContentsMargins(0, 0, 0, 0)
        right_col_layout.setSpacing(8)

        steps_box = self._create_section("小节点")
        steps_layout = steps_box.layout()
        header = QHBoxLayout()
        header.addWidget(QLabel("双击可直接修改"))
        self.default_delay_button = QPushButton()
        self.default_delay_button.clicked.connect(self.edit_default_delay)
        header.addWidget(self.default_delay_button)
        self.step_view_list_button = QPushButton("切换为列表示图")
        self.step_view_list_button.clicked.connect(self.switch_step_view_to_list)
        header.addWidget(self.step_view_list_button)
        self.step_view_code_button = QPushButton("切换为代码视图")
        self.step_view_code_button.clicked.connect(self.switch_step_view_to_code)
        header.addWidget(self.step_view_code_button)
        self.step_view_flow_button = QPushButton("切换为流程图查看")
        self.step_view_flow_button.clicked.connect(self.switch_step_view_to_flowchart)
        header.addWidget(self.step_view_flow_button)
        header.addStretch(1)
        steps_layout.addLayout(header)
        self.step_list = StepListWidget()
        self.step_list.row_changed.connect(self.on_step_selected)
        self.step_list.step_double_clicked.connect(self.edit_step_quickly)
        self.step_list.action_dropped.connect(self.insert_step_from_drop)
        self.step_list.header().setStretchLastSection(True)
        self.step_code_editor = QTextEdit()
        self.step_code_editor.setPlaceholderText("请输入小节点 JSON 数组")
        self.step_flowchart_view = QTextEdit()
        self.step_flowchart_view.setReadOnly(True)
        self.step_view_stack = QStackedWidget()
        self.step_view_stack.addWidget(self.step_list)
        self.step_view_stack.addWidget(self.step_code_editor)
        self.step_view_stack.addWidget(self.step_flowchart_view)
        steps_layout.addWidget(self.step_view_stack, 1)
        step_row = QHBoxLayout()
        self.step_operation_buttons: list[QPushButton] = []
        for text, callback in [("复制", self.copy_step), ("删除", self.delete_step), ("上移", self.move_step_up), ("下移", self.move_step_down)]:
            button = QPushButton(text)
            button.clicked.connect(callback)
            step_row.addWidget(button)
            self.step_operation_buttons.append(button)
        step_row.addStretch(1)
        steps_layout.addLayout(step_row)
        right_col_layout.addWidget(steps_box, 4)

        palette_box = self._create_section("增加动作")
        palette_layout = palette_box.layout()
        self.action_palette = ActionPalette()
        self.action_palette.action_clicked.connect(self.insert_step)
        palette_layout.addWidget(self.action_palette)
        right_col_layout.addWidget(palette_box, 2)

        top_split.addWidget(left_col)
        top_split.addWidget(right_col)
        top_split.setSizes([360, 500])

        return panel

    def _build_right_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        overview_box = self._create_section("概览")
        overview_layout = overview_box.layout()
        self.overview_text = QTextEdit()
        self.overview_text.setReadOnly(True)
        self.overview_text.setMinimumHeight(220)
        overview_layout.addWidget(self.overview_text)
        layout.addWidget(overview_box, 2)

        matched_box = self._create_section("匹配窗口")
        matched_layout = matched_box.layout()
        self.matched_windows_list = QListWidget()
        matched_layout.addWidget(self.matched_windows_list)
        layout.addWidget(matched_box, 1)

        log_box = self._create_section("日志")
        log_layout = log_box.layout()
        log_toolbar = QHBoxLayout()
        start_button = QPushButton("全部运行")
        start_button.clicked.connect(self.start_workflow)
        start_current_button = QPushButton("运行当前大节点")
        start_current_button.clicked.connect(self.start_current_node)
        stop_button = QPushButton("停止")
        stop_button.clicked.connect(self.stop_workflow)
        refresh_windows_button = QPushButton("刷新窗口")
        refresh_windows_button.clicked.connect(self.refresh_windows)
        clear_button = QPushButton("清空日志")
        clear_button.clicked.connect(lambda: self.log_text.clear())
        log_toolbar.addWidget(start_button)
        log_toolbar.addWidget(start_current_button)
        log_toolbar.addWidget(stop_button)
        log_toolbar.addWidget(refresh_windows_button)
        log_toolbar.addStretch(1)
        log_toolbar.addWidget(clear_button)
        log_layout.addLayout(log_toolbar)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text, 1)
        layout.addWidget(log_box, 2)
        return panel

    def _create_section(self, title: str) -> QFrame:
        frame = QFrame()
        frame.setFrameShape(QFrame.StyledPanel)
        frame.setStyleSheet("QFrame { border: 1px solid #d0d5dd; border-radius: 8px; }")
        wrapper = QVBoxLayout(frame)
        wrapper.setContentsMargins(10, 10, 10, 10)
        wrapper.setSpacing(8)
        label = QLabel(title)
        label.setStyleSheet("font-weight: 700;")
        wrapper.addWidget(label)
        return frame

    def log(self, message: str):
        timestamp = time.strftime("%H:%M:%S")
        level = self._detect_log_level(message)
        msg_color = self.LOG_COLORS.get(level, self.LOG_COLORS["default"])
        safe_message = html.escape(str(message))
        self.log_text.append(
            f"<span style='color:{self.LOG_COLORS['timestamp']};'>[{timestamp}]</span> "
            f"<span style='color:{msg_color};'>{safe_message}</span>"
        )
        self.status_label.setText(message)
        self.status_label.setStyleSheet(f"color: {msg_color};")

    def _detect_log_level(self, message: str) -> str:
        text = str(message or "").lower()
        if any(k in text for k in ["失败", "error", "异常", "invalid", "无效", "不支持", "traceback"]):
            return "error"
        if any(k in text for k in ["取消", "canceled", "stop signal", "已停止", "没检测到", "未匹配"]):
            return "warning"
        if any(k in text for k in ["已保存", "完成", "检测到", "已更新", "已新增", "已删除", "saved", "completed"]):
            return "success"
        if any(k in text for k in ["开始", "started", "执行", "进入大节点", "目标窗口", "刷新窗口", "加载工作流"]):
            return "running"
        return "default"

    def refresh_windows(self):
        self.windows = enumerate_windows()
        self._refresh_matched_windows_for_current_node()
        self.log(f"刷新窗口列表，共 {len(self.windows)} 个")

    def refresh_all(self):
        self.workflow_index = min(self.workflow_index, len(self.workflows) - 1)
        self.workflow_list.blockSignals(True)
        self.workflow_list.clear()
        for index, workflow in enumerate(self.workflows):
            file_path = self.workflow_files[index] if index < len(self.workflow_files) else None
            if file_path is None:
                label = f"(未保存) {workflow['name']}"
            else:
                label = f"{file_path.name} | {workflow['name']}"
            self.workflow_list.addItem(label)
        self.workflow_list.setCurrentRow(self.workflow_index)
        self.workflow_list.blockSignals(False)

        big_nodes = get_big_nodes(self.workflow)
        self.node_index = min(self.node_index, len(big_nodes) - 1)
        self.node_list.blockSignals(True)
        self.node_list.clear()
        for index, node in enumerate(big_nodes, start=1):
            scope = node.get("scope", {})
            flow = node.get("flow", {})
            item = QTreeWidgetItem(
                [
                    self._big_node_display_alias(index - 1, node),
                    scope.get("regex", ""),
                    self._multi_window_mode_label(scope.get("multi_window_mode", "first")),
                    self._flow_summary(flow),
                    str(node.get("node_delay", "0") or "0"),
                ]
            )
            self.node_list.addTopLevelItem(item)
        if self.node_list.topLevelItemCount() > 0:
            self.node_list.setCurrentItem(self.node_list.topLevelItem(self.node_index))
        self.node_list.blockSignals(False)
        if self.node_view_mode == "code":
            self._fill_node_code_editor()
        elif self.node_view_mode == "flow":
            self._fill_node_flowchart_view()
        self.refresh_steps()

    def _step_params_detail(self, step: dict) -> str:
        action = step.get("action", "")
        params = step.get("params", {})
        fields = ACTION_DEFINITIONS.get(action, [])
        details = []
        for field, label, default in fields:
            raw = params.get(field, default)
            value = "" if raw is None else str(raw).strip()
            if value:
                details.append(f"{label}: {value}")
        return " | ".join(details) if details else "-"

    def refresh_steps(self):
        big_nodes = get_big_nodes(self.workflow)
        small_nodes = get_small_nodes(self.node)
        self.default_delay_button.setText(f"默认延迟: {self.default_delay}s")
        self.step_index = min(self.step_index, len(small_nodes) - 1)
        self.step_list.blockSignals(True)
        self.step_list.clear()
        for index, step in enumerate(small_nodes, start=1):
            item = QTreeWidgetItem(
                [
                    step["action"],
                    self._step_display_alias(index - 1, step),
                    str(step.get("delay", self.default_delay)),
                    self._step_params_detail(step),
                ]
            )
            item.setData(0, Qt.UserRole, index - 1)
            self.step_list.addTopLevelItem(item)
        if self.step_list.topLevelItemCount() > 0:
            self.step_list.setCurrentItem(self.step_list.topLevelItem(self.step_index))
        self.step_list.blockSignals(False)
        if self.step_view_mode == "code":
            self._fill_step_code_editor()
        elif self.step_view_mode == "flow":
            self._fill_step_flowchart_view()
        self.overview_text.setPlainText(self._workflow_flowchart_summary())

    def _fill_step_code_editor(self) -> None:
        self.step_code_editor.blockSignals(True)
        self.step_code_editor.setPlainText(
            json.dumps(get_small_nodes(self.node), ensure_ascii=False, indent=2)
        )
        self.step_code_editor.blockSignals(False)

    def _fill_step_flowchart_view(self) -> None:
        small_nodes = get_small_nodes(self.node)
        node_alias = self._big_node_display_alias(self.node_index, self.node)
        lines: list[str] = [f"小节点流程图：{node_alias}", ""]
        for index, step in enumerate(small_nodes, start=1):
            lines.append(f"[{index}] {self._step_display_alias(index - 1, step)}")
            lines.append(f"  ├─ 动作: {step.get('action', '')}")
            lines.append(f"  ├─ 延迟: {step.get('delay', self.default_delay)}s")
            if index < len(small_nodes):
                lines.append(f"  └─ 下一步 -> [{index + 1}]")
            else:
                lines.append("  └─ 结束")
            lines.append("")
        self.step_flowchart_view.setPlainText("\n".join(lines).strip())

    def _fill_node_code_editor(self) -> None:
        self.node_code_editor.blockSignals(True)
        self.node_code_editor.setPlainText(
            json.dumps(get_big_nodes(self.workflow), ensure_ascii=False, indent=2)
        )
        self.node_code_editor.blockSignals(False)

    def _fill_node_flowchart_view(self) -> None:
        nodes = get_big_nodes(self.workflow)
        lines: list[str] = [f"工作流流程图：{self.workflow.get('name', '工作流')}", ""]
        for index, node in enumerate(nodes, start=1):
            alias = self._big_node_display_alias(index - 1, node)
            flow = node.get("flow", {})
            mode = str(flow.get("mode", "next") or "next").strip().lower()
            target = str(flow.get("target", "") or "").strip()
            condition = str(flow.get("condition", "last_detected") or "last_detected").strip().lower()
            node_delay = str(node.get("node_delay", "0") or "0").strip()
            lines.append(f"[{index}] {alias}")
            if mode == "loop":
                loops = str(flow.get("max_loops", "1") or "1").strip()
                lines.append(f"  ├─ 循环 x{loops or '1'}，结束后 -> [{index + 1}]")
            elif mode == "jump":
                lines.append(f"  ├─ 无条件跳转 -> [{target or '?'}]")
            elif mode == "conditional_jump":
                if condition == "last_detected":
                    cond_text = "检测到"
                elif condition == "last_not_detected":
                    cond_text = "未检测到"
                elif condition == "window_text_detected":
                    cond_text = "窗口文字命中"
                elif condition == "window_image_detected":
                    cond_text = "窗口图片命中"
                else:
                    cond_text = condition or "条件"
                lines.append(f"  ├─ 条件({cond_text}) 满足 -> [{target or '?'}]")
                lines.append(f"  └─ 条件不满足 -> [{index + 1}]")
            elif mode == "stop":
                lines.append("  └─ 执行完毕后停止工作流")
            else:
                lines.append(f"  └─ 顺序 -> [{index + 1}]")
            lines.append(f"  节点间延迟: {node_delay or '0'}s")
            lines.append("")
        self.node_flowchart_view.setPlainText("\n".join(lines).strip())

    def _apply_node_code_editor(self) -> bool:
        raw = self.node_code_editor.toPlainText().strip()
        if not raw:
            QMessageBox.warning(self, "代码视图错误", "大节点代码不能为空。")
            return False
        try:
            payload = json.loads(raw)
        except Exception as exc:
            QMessageBox.warning(self, "代码视图错误", f"JSON 解析失败：{exc}")
            return False

        if not isinstance(payload, list):
            QMessageBox.warning(self, "代码视图错误", "大节点代码必须是 JSON 数组。")
            return False

        try:
            normalized_nodes: list[dict] = []
            for index, item in enumerate(payload, start=1):
                if not isinstance(item, dict):
                    raise ValueError(f"第 {index} 项必须是对象")
                normalized_nodes.append(normalize_big_node(item, index))
        except Exception as exc:
            QMessageBox.warning(self, "代码视图错误", f"大节点结构无效：{exc}")
            return False

        if not normalized_nodes:
            normalized_nodes = [create_big_node(1)]
        self.workflow["big_nodes"] = normalized_nodes
        self.node_index = min(self.node_index, len(normalized_nodes) - 1)
        self.step_index = 0
        return True

    def _apply_step_code_editor(self) -> bool:
        raw = self.step_code_editor.toPlainText().strip()
        if not raw:
            QMessageBox.warning(self, "代码视图错误", "小节点代码不能为空。")
            return False
        try:
            payload = json.loads(raw)
        except Exception as exc:
            QMessageBox.warning(self, "代码视图错误", f"JSON 解析失败：{exc}")
            return False

        if not isinstance(payload, list):
            QMessageBox.warning(self, "代码视图错误", "小节点代码必须是 JSON 数组。")
            return False

        try:
            normalized_steps: list[dict] = []
            for index, item in enumerate(payload, start=1):
                if not isinstance(item, dict):
                    raise ValueError(f"第 {index} 项必须是对象")
                normalized_steps.append(normalize_small_node(item))
        except Exception as exc:
            QMessageBox.warning(self, "代码视图错误", f"小节点结构无效：{exc}")
            return False

        if not normalized_steps:
            normalized_steps = [create_small_node()]
        self.node["small_nodes"] = normalized_steps
        self.step_index = min(self.step_index, len(normalized_steps) - 1)
        return True

    def _set_step_view_mode(self, mode: str) -> None:
        if mode not in {"list", "code", "flow"}:
            mode = "list"
        list_mode = mode == "list"
        self.step_view_mode = mode
        self.step_view_stack.setCurrentIndex({"list": 0, "code": 1, "flow": 2}[mode])
        self.step_view_list_button.setEnabled(not list_mode)
        self.step_view_code_button.setEnabled(mode != "code")
        self.step_view_flow_button.setEnabled(mode != "flow")
        self.default_delay_button.setEnabled(list_mode)
        for button in self.step_operation_buttons:
            button.setEnabled(list_mode)
        if hasattr(self, "action_palette"):
            self.action_palette.setEnabled(list_mode)

    def _set_node_view_mode(self, mode: str) -> None:
        if mode not in {"list", "code", "flow"}:
            mode = "list"
        list_mode = mode == "list"
        self.node_view_mode = mode
        self.node_view_stack.setCurrentIndex({"list": 0, "code": 1, "flow": 2}[mode])
        self.node_view_list_button.setEnabled(not list_mode)
        self.node_view_code_button.setEnabled(mode != "code")
        self.node_view_flow_button.setEnabled(mode != "flow")
        for button in self.node_operation_buttons:
            button.setEnabled(list_mode)

    def switch_node_view_to_code(self) -> None:
        if self.node_view_mode == "code":
            return
        self._fill_node_code_editor()
        self._set_node_view_mode("code")
        self.log("大节点已切换为代码视图")

    def switch_node_view_to_flowchart(self) -> None:
        if self.node_view_mode == "flow":
            return
        self._fill_node_flowchart_view()
        self._set_node_view_mode("flow")
        self.log("大节点已切换为流程图查看")

    def switch_node_view_to_list(self) -> None:
        if self.node_view_mode == "list":
            return
        if self.node_view_mode == "code" and not self._apply_node_code_editor():
            return
        self._set_node_view_mode("list")
        self.refresh_all()
        self.log("大节点已切换为列表示图")

    def switch_step_view_to_code(self) -> None:
        if self.step_view_mode == "code":
            return
        self._fill_step_code_editor()
        self._set_step_view_mode("code")
        self.log("已切换为代码视图")

    def switch_step_view_to_flowchart(self) -> None:
        if self.step_view_mode == "flow":
            return
        self._fill_step_flowchart_view()
        self._set_step_view_mode("flow")
        self.log("小节点已切换为流程图查看")

    def switch_step_view_to_list(self) -> None:
        if self.step_view_mode == "list":
            return
        if self.step_view_mode == "code" and not self._apply_step_code_editor():
            return
        self._set_step_view_mode("list")
        self.refresh_steps()
        self.log("已切换为列表示图")

    def on_workflow_selected(self, index: int):
        if index < 0:
            return
        self.workflow_index = index
        selected_file = self.workflow_files[index] if index < len(self.workflow_files) else None
        self.current_file = str(selected_file) if selected_file is not None else None
        self.node_index = 0
        self.step_index = 0
        self.refresh_all()

    def on_node_selected(self, index: int):
        if index < 0:
            return
        self.node_index = index
        self.step_index = 0
        self.refresh_steps()
        self._refresh_matched_windows_for_current_node()

    def on_step_selected(self, index: int):
        if index < 0:
            return
        self.step_index = index

    def _map_runner_node_index(self, node_index: int) -> int:
        mapped = node_index
        if self._runner_node_index_map and 0 <= node_index < len(self._runner_node_index_map):
            mapped = self._runner_node_index_map[node_index]
        return max(0, min(mapped, len(get_big_nodes(self.workflow)) - 1))

    def _render_matched_windows(self, windows: list[dict], empty_text: str = "(未匹配到窗口)") -> None:
        self.matched_windows_list.clear()
        self._matched_window_hwnds = []
        if not windows:
            self.matched_windows_list.addItem(empty_text)
            return

        for item in windows:
            hwnd = int(item.get("hwnd", 0))
            title = str(item.get("title", ""))
            text = f"{title}  [hwnd={hwnd}]"
            lw_item = QListWidgetItem(text)
            lw_item.setData(Qt.UserRole, hwnd)
            self.matched_windows_list.addItem(lw_item)
            self._matched_window_hwnds.append(hwnd)

        self.matched_windows_list.setCurrentRow(0)

    def _refresh_matched_windows_for_current_node(self) -> None:
        if not self.workflows:
            self._render_matched_windows([])
            return
        big_nodes = get_big_nodes(self.workflow)
        if not big_nodes:
            self._render_matched_windows([])
            return
        self.node_index = max(0, min(self.node_index, len(big_nodes) - 1))
        scope = big_nodes[self.node_index].get("scope", {})
        try:
            windows = resolve_scope_windows(scope, self.selected_hwnd, list(self.windows))
        except Exception:
            windows = []
        self._render_matched_windows(windows)

    def _on_runner_windows_resolved(self, payload: dict):
        node_index = self._map_runner_node_index(int(payload.get("node_index", 0)))
        windows = payload.get("windows", []) or []

        # Switch UI context to the running node so step highlight is visible.
        if node_index != self.node_index:
            self.node_index = node_index
            self.step_index = 0
            self.refresh_steps()

        self._render_matched_windows(windows)

    def _on_runner_step_started(self, payload: dict):
        node_index = self._map_runner_node_index(int(payload.get("node_index", 0)))
        step_index = int(payload.get("step_index", 0))
        hwnd = payload.get("hwnd", None)

        if node_index != self.node_index:
            self.node_index = node_index
            self.step_index = 0
            self.refresh_steps()

        self.step_index = step_index
        if 0 <= step_index < self.step_list.topLevelItemCount():
            item = self.step_list.topLevelItem(step_index)
            self.step_list.setCurrentItem(item)
            self.step_list.scrollToItem(item)

        if hwnd is None:
            return
        for row in range(self.matched_windows_list.count()):
            item = self.matched_windows_list.item(row)
            if item is not None and item.data(Qt.UserRole) == hwnd:
                self.matched_windows_list.setCurrentRow(row)
                break

    def edit_step_quickly(self, index: int):
        small_nodes = get_small_nodes(self.node)
        if index < 0 or index >= len(small_nodes):
            return

        self.step_index = index
        step = small_nodes[index]
        updated = edit_step_dialog(self, step)
        if updated is None:
            return

        step.update(updated)
        self.refresh_steps()
        self.log("已更新小节点")

    def edit_big_node_quickly(self, index: int):
        big_nodes = get_big_nodes(self.workflow)
        if index < 0 or index >= len(big_nodes):
            return

        self.node_index = index
        node = big_nodes[index]
        updated = edit_big_node_dialog(self, node)
        if updated is None:
            return

        node.update(updated)
        self.refresh_all()
        self.log("已更新大节点")

    def new_workflow(self):
        self.workflows.append(create_workflow())
        self.workflow_files.append(None)
        self.workflow_index = len(self.workflows) - 1
        self.current_file = None
        self.node_index = 0
        self.step_index = 0
        self.refresh_all()
        self.log("已新建工作流")

    def delete_workflow(self):
        if len(self.workflows) == 1:
            QMessageBox.information(self, "不能删除", "至少保留一个工作流。")
            return
        target_file = self.workflow_files[self.workflow_index] if self.workflow_index < len(self.workflow_files) else None
        if target_file is not None:
            confirm = QMessageBox.question(
                self,
                "删除工作流",
                f"将删除文件：\n{target_file}\n\n确定继续吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if confirm != QMessageBox.Yes:
                return
            try:
                if target_file.exists():
                    target_file.unlink()
            except Exception as exc:
                QMessageBox.warning(self, "删除失败", f"无法删除文件：{exc}")
                return
        self.workflows.pop(self.workflow_index)
        if self.workflow_index < len(self.workflow_files):
            self.workflow_files.pop(self.workflow_index)
        self.workflow_index = max(0, self.workflow_index - 1)
        selected_file = self.workflow_files[self.workflow_index] if self.workflow_files else None
        self.current_file = str(selected_file) if selected_file is not None else None
        self.node_index = 0
        self.step_index = 0
        self.refresh_all()
        self.log("已删除工作流")

    def add_node(self):
        big_nodes = get_big_nodes(self.workflow)
        big_nodes.append(create_big_node(len(big_nodes) + 1))
        self.node_index = len(big_nodes) - 1
        self.step_index = 0
        self.refresh_all()
        self.log("已新增大节点")

    def delete_node(self):
        big_nodes = get_big_nodes(self.workflow)
        if len(big_nodes) == 1:
            QMessageBox.information(self, "不能删除", "至少保留一个大节点。")
            return
        big_nodes.pop(self.node_index)
        self.node_index = max(0, self.node_index - 1)
        self.step_index = 0
        self.refresh_all()
        self.log("已删除大节点")

    def move_node_up(self):
        index = self.node_index
        if index <= 0:
            return
        nodes = get_big_nodes(self.workflow)
        nodes[index - 1], nodes[index] = nodes[index], nodes[index - 1]
        self.node_index -= 1
        self.refresh_all()

    def move_node_down(self):
        index = self.node_index
        nodes = get_big_nodes(self.workflow)
        if index >= len(nodes) - 1:
            return
        nodes[index + 1], nodes[index] = nodes[index], nodes[index + 1]
        self.node_index += 1
        self.refresh_all()

    def save_scope(self):
        node_count = len(get_big_nodes(self.workflow))
        for index, node in enumerate(get_big_nodes(self.workflow), start=1):
            scope = node.get("scope", {})
            node_delay = str(node.get("node_delay", "0") or "0").strip()
            regex = (scope.get("regex") or "").strip()
            if not regex:
                pass
            else:
                try:
                    re.compile(regex)
                except re.error as exc:
                    raise ValueError(f"大节点 {index} 作用域正则无效: {exc}") from exc

            flow = node.get("flow", {})
            mode = str(flow.get("mode", "next") or "next").strip().lower()
            target = str(flow.get("target", "") or "").strip()
            max_loops = str(flow.get("max_loops", "1") or "1").strip()

            if mode in {"jump", "conditional_jump"}:
                if not target:
                    raise ValueError(f"大节点 {index} 跳转目标不能为空。")
                try:
                    target_index = int(target)
                except Exception as exc:
                    raise ValueError(f"大节点 {index} 跳转目标必须是数字。") from exc
                if target_index < 1 or target_index > node_count:
                    raise ValueError(f"大节点 {index} 跳转目标超出范围(1-{node_count})。")

            if mode == "conditional_jump":
                condition = str(flow.get("condition", "last_detected") or "last_detected").strip().lower()
                if condition == "window_text_detected":
                    text = str(flow.get("condition_target_text", "") or "").strip()
                    if not text:
                        raise ValueError(f"大节点 {index} 条件跳转(窗口文字)未配置条件文字。")
                if condition == "window_image_detected":
                    image_path = str(flow.get("condition_image_path", "") or "").strip()
                    if not image_path:
                        raise ValueError(f"大节点 {index} 条件跳转(窗口图片)未配置条件图片路径。")

            if mode == "loop":
                try:
                    loops = int(max_loops)
                except Exception as exc:
                    raise ValueError(f"大节点 {index} 循环次数必须是数字。") from exc
                if loops < 1:
                    raise ValueError(f"大节点 {index} 循环次数必须 >= 1。")

            try:
                node_delay_value = float(node_delay)
            except Exception as exc:
                raise ValueError(f"大节点 {index} 节点间延迟必须是数字。") from exc
            if node_delay_value < 0:
                raise ValueError(f"大节点 {index} 节点间延迟必须 >= 0。")

    def load_scope(self):
        # 预留扩展点：如果未来增加“窗口作用域面板”，可在这里统一回填 UI。
        return

    def _default_coord_mode_for_current_node(self) -> str:
        regex = str(self.node.get("scope", {}).get("regex", "")).strip()
        return "relative" if regex else "absolute"

    def _effective_coord_mode(self, params: dict) -> str:
        raw = str(params.get("coord_mode", "")).strip().lower()
        if raw == "absolute":
            return "absolute"
        return self._default_coord_mode_for_current_node()

    def _apply_default_coord_mode_if_needed(self, step: dict) -> None:
        action = str(step.get("action", "")).strip()
        if action not in self.COORD_ACTIONS:
            return
        params = step.get("params", {})
        if not isinstance(params, dict):
            return
        params["coord_mode"] = self._default_coord_mode_for_current_node()

    def insert_step(self, action_name: str, index: int | None = None):
        new_step = create_small_node(action_name)
        self._apply_default_coord_mode_if_needed(new_step)
        if index is None:
            small_nodes = get_small_nodes(self.node)
            small_nodes.append(new_step)
            self.step_index = len(small_nodes) - 1
        else:
            small_nodes = get_small_nodes(self.node)
            index = max(0, min(index, len(small_nodes)))
            small_nodes.insert(index, new_step)
            self.step_index = index
        self.refresh_steps()
        self._auto_launch_capture_tool_for_new_step(action_name)
        self.log("已新增小节点")

    def insert_step_from_drop(self, action_name: str, index: int):
        self.insert_step(action_name, index=index)

    def edit_default_delay(self):
        value, accepted = QInputDialog.getText(
            self,
            "默认延迟",
            "请输入新增小节点的默认延迟(s):",
            text=self.default_delay,
        )
        if not accepted:
            return

        cleaned = value.strip()
        if not cleaned:
            QMessageBox.warning(self, "设置失败", "默认延迟不能为空。")
            return

        try:
            if float(cleaned) < 0:
                raise ValueError
        except ValueError:
            QMessageBox.warning(self, "设置失败", "默认延迟必须是大于等于 0 的数字。")
            return

        self.default_delay = cleaned
        set_default_delay(cleaned)
        self._save_session_state()
        self.refresh_steps()
        self.log(f"默认延迟已更新为 {cleaned}s")

    def _resolve_preview_window(self) -> dict | None:
        try:
            return resolve_scope_window(self.node.get("scope", {}), self.selected_hwnd, list(self.windows))
        except Exception:
            return None

    def _create_rel_provider(self):
        target = self._resolve_preview_window()
        if not target:
            return None

        def provider(abs_x: int, abs_y: int):
            left, top, _right, _bottom = get_window_frame_rect(int(target["hwnd"]))
            return abs_x - left, abs_y - top

        return provider

    def _auto_launch_capture_tool_for_new_step(self, action_name: str):
        small_nodes = get_small_nodes(self.node)
        if self.step_index < 0 or self.step_index >= len(small_nodes):
            return
        step = small_nodes[self.step_index]
        params = step.get("params", {})

        def apply_point(x_key: str, y_key: str, abs_x: int, abs_y: int) -> None:
            coord_mode = self._effective_coord_mode(params)
            if coord_mode != "absolute":
                rel_provider = self._create_rel_provider()
                rel_pos = rel_provider(abs_x, abs_y) if rel_provider else None
                if rel_pos is not None:
                    params[x_key], params[y_key] = str(rel_pos[0]), str(rel_pos[1])
                    return
                params[x_key], params[y_key] = str(abs_x), str(abs_y)
                params["coord_mode"] = "absolute"
                self.log("未找到预览窗口，相对坐标已自动回退为 absolute。")
                return
            params[x_key], params[y_key] = str(abs_x), str(abs_y)

        if action_name == ACTION_CLICK:
            overlay = PointPickerOverlay(rel_provider=self._create_rel_provider())
            self._active_point_overlay = overlay

            def on_point_picked(abs_x: int, abs_y: int):
                apply_point("x", "y", abs_x, abs_y)
                self.refresh_steps()
                self.log(f"Point captured: ({params.get('x')}, {params.get('y')})")

            def on_canceled():
                self.log("Point capture canceled.")

            overlay.point_picked.connect(on_point_picked)
            overlay.canceled.connect(on_canceled)
            overlay.show()
            overlay.activateWindow()
            return

        if action_name in {ACTION_MOUSE_DRAG, ACTION_MOUSE_SCROLL}:
            is_drag = action_name == ACTION_MOUSE_DRAG
            action_label = "拖动" if is_drag else "滑动"
            start_overlay = PointPickerOverlay(rel_provider=self._create_rel_provider())
            self._active_point_overlay = start_overlay

            def open_end_picker() -> None:
                end_overlay = PointPickerOverlay(rel_provider=self._create_rel_provider())
                self._active_point_overlay = end_overlay

                def on_end_picked(abs_x: int, abs_y: int):
                    apply_point("end_x", "end_y", abs_x, abs_y)
                    self.refresh_steps()
                    self.log(
                        f"{action_label}终点已捕获: ({params.get('end_x')}, {params.get('end_y')})"
                    )

                def on_end_canceled():
                    self.log(f"{action_label}终点捕获已取消。")

                end_overlay.point_picked.connect(on_end_picked)
                end_overlay.canceled.connect(on_end_canceled)
                end_overlay.show()
                end_overlay.activateWindow()

            def on_start_picked(abs_x: int, abs_y: int):
                apply_point("start_x", "start_y", abs_x, abs_y)
                self.refresh_steps()
                self.log(
                    f"{action_label}起点已捕获: ({params.get('start_x')}, {params.get('start_y')})，请继续选择终点。"
                )
                QTimer.singleShot(120, open_end_picker)

            def on_start_canceled():
                self.log(f"{action_label}起点捕获已取消。")

            start_overlay.point_picked.connect(on_start_picked)
            start_overlay.canceled.connect(on_start_canceled)
            start_overlay.show()
            start_overlay.activateWindow()
            return

        if action_name in {ACTION_CLICK_IMAGE, ACTION_DETECT_IMAGE}:
            templates_dir = Path.cwd() / "templates"
            default_target = templates_dir / f"template_{time.strftime('%Y%m%d_%H%M%S')}_{uuid4().hex[:8]}.png"
            overlay = ScreenshotOverlay(default_target=str(default_target), prompt_for_save=False)
            self._active_screenshot_overlay = overlay

            def on_screenshot_saved(path: str):
                data_uri = self._image_file_to_data_uri(path)
                if data_uri is not None:
                    params["image_path"] = data_uri
                    try:
                        Path(path).unlink(missing_ok=True)
                    except Exception:
                        pass
                else:
                    params["image_path"] = path
                self.refresh_steps()
                if data_uri is not None:
                    self.log("已保存模板截图并内嵌为Base64")
                else:
                    self.log(f"已保存模板截图: {path}")

            def on_canceled():
                self.log("截图已取消。")

            overlay.screenshot_saved.connect(on_screenshot_saved)
            overlay.canceled.connect(on_canceled)
            overlay.show()
            overlay.raise_()
            overlay.activateWindow()

    def copy_step(self):
        current = self.step
        get_small_nodes(self.node).insert(self.step_index + 1, clone_payload(current))
        self.step_index += 1
        self.refresh_steps()
        self.log("已复制小节点")

    def delete_step(self):
        small_nodes = get_small_nodes(self.node)
        if len(small_nodes) == 1:
            QMessageBox.information(self, "不能删除", "至少保留一个小节点。")
            return
        small_nodes.pop(self.step_index)
        self.step_index = max(0, self.step_index - 1)
        self.refresh_steps()
        self.log("已删除小节点")

    def move_step_up(self):
        index = self.step_index
        if index <= 0:
            return
        steps = get_small_nodes(self.node)
        steps[index - 1], steps[index] = steps[index], steps[index - 1]
        self.step_index -= 1
        self.refresh_steps()

    def move_step_down(self):
        index = self.step_index
        steps = get_small_nodes(self.node)
        if index >= len(steps) - 1:
            return
        steps[index + 1], steps[index] = steps[index], steps[index + 1]
        self.step_index += 1
        self.refresh_steps()

    @staticmethod
    def _sanitize_workflow_file_name(name: str) -> str:
        cleaned = re.sub(r'[\\/:*?"<>|]+', "_", (name or "").strip())
        cleaned = cleaned.strip(" .")
        return cleaned or "workflow"

    def _read_workflow_from_file(self, path: Path) -> dict:
        payload = json.loads(path.read_text(encoding="utf-8"))
        workflows = payload.get("workflows")
        if not isinstance(workflows, list) or not workflows:
            raise ValueError("missing workflows list")
        workflow = normalize_workflow(workflows[0])
        self._materialize_embedded_images(workflow)
        return workflow

    @staticmethod
    def _is_data_image_uri(value: str) -> bool:
        return value.startswith("data:image/") and ";base64," in value

    def _data_uri_to_cached_file(self, value: str) -> Path:
        if not self._is_data_image_uri(value):
            raise ValueError("invalid data uri")
        header, encoded = value.split(",", 1)
        mime = header[5:].split(";", 1)[0].strip().lower()
        extension = mimetypes.guess_extension(mime) or ".png"
        payload = base64.b64decode(encoded)
        digest = sha256(payload).hexdigest()[:24]
        target = self.EMBEDDED_IMAGE_DIR / f"{digest}{extension}"
        if not target.exists():
            target.write_bytes(payload)
        return target

    def _materialize_embedded_images(self, workflow: dict) -> None:
        for node in get_big_nodes(workflow):
            for step in get_small_nodes(node):
                action = str(step.get("action", ""))
                if action not in {ACTION_CLICK_IMAGE, ACTION_DETECT_IMAGE}:
                    continue
                params = step.get("params", {})
                image_path = str(params.get("image_path", "")).strip()
                if not image_path or not self._is_data_image_uri(image_path):
                    continue
                try:
                    params["image_path"] = str(self._data_uri_to_cached_file(image_path))
                except Exception:
                    self.log("检测到内嵌图片但解析失败，已保留原始内容")

    def _iter_image_steps(self, workflow: dict):
        for node in get_big_nodes(workflow):
            for step in get_small_nodes(node):
                action = str(step.get("action", ""))
                if action in {ACTION_CLICK_IMAGE, ACTION_DETECT_IMAGE}:
                    yield step

    def _image_file_to_data_uri(self, image_value: str) -> str | None:
        if self.image_storage_mode == self.IMAGE_STORAGE_FILE:
            return None
        image_file = Path(image_value).expanduser()
        if not image_file.is_absolute():
            image_file = (Path.cwd() / image_file).resolve()
        if not image_file.exists() or not image_file.is_file():
            return None
        if (
            self.image_storage_mode == self.IMAGE_STORAGE_BASE64_AUTO
            and image_file.stat().st_size >= self.EMBEDDED_IMAGE_MAX_BYTES
        ):
            return None

        mime = mimetypes.guess_type(image_file.name)[0] or "image/png"
        if not mime.startswith("image/"):
            return None
        encoded = base64.b64encode(image_file.read_bytes()).decode("ascii")
        return f"data:{mime};base64,{encoded}"

    def _embed_images_for_save(self, workflow: dict) -> tuple[dict, int]:
        if self.image_storage_mode == self.IMAGE_STORAGE_FILE:
            return workflow, 0
        embedded_count = 0
        for step in self._iter_image_steps(workflow):
            params = step.get("params", {})
            image_value = str(params.get("image_path", "")).strip()
            if not image_value or self._is_data_image_uri(image_value):
                continue

            data_uri = self._image_file_to_data_uri(image_value)
            if data_uri is None:
                continue
            params["image_path"] = data_uri
            embedded_count += 1
        return workflow, embedded_count

    def _load_workflows_from_directory(self, preferred_path: Path | None = None) -> int:
        files = sorted(self.WORKFLOW_DIR.glob("*.json"), key=lambda item: item.name.lower())
        loaded_workflows: list[dict] = []
        loaded_files: list[Path | None] = []
        failed_files: list[tuple[Path, str]] = []
        for file_path in files:
            try:
                loaded_workflows.append(self._read_workflow_from_file(file_path))
                loaded_files.append(file_path)
            except Exception as exc:
                failed_files.append((file_path, str(exc)))

        if not loaded_workflows:
            loaded_workflows = [create_workflow()]
            loaded_files = [None]

        selected_path = preferred_path
        if selected_path is None and self.current_file:
            selected_path = Path(self.current_file)

        selected_index = 0
        if selected_path is not None:
            for index, file_path in enumerate(loaded_files):
                if file_path is not None and file_path.resolve() == selected_path.resolve():
                    selected_index = index
                    break

        self.workflows = loaded_workflows
        self.workflow_files = loaded_files
        self.workflow_index = selected_index
        self.node_index = 0
        self.step_index = 0
        selected_file = self.workflow_files[self.workflow_index] if self.workflow_files else None
        self.current_file = str(selected_file) if selected_file is not None else None
        self.refresh_all()
        for file_path, reason in failed_files:
            self.log(f"跳过无效工作流文件: {file_path.name} ({reason})")
        return len(files)

    def save_file(self):
        target_file = self.workflow_files[self.workflow_index] if self.workflow_index < len(self.workflow_files) else None
        if target_file is None:
            self.save_file_as()
            return
        payload_workflow = clone_payload(self.workflow)
        payload_workflow, embedded_count = self._embed_images_for_save(payload_workflow)
        target_file.write_text(
            json.dumps({"workflows": [payload_workflow]}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        self.current_file = str(target_file)
        self._save_session_state()
        self.refresh_all()
        if embedded_count > 0:
            if self.image_storage_mode == self.IMAGE_STORAGE_BASE64_AUTO:
                self.log(f"已内嵌 {embedded_count} 张小于 3MB 的图片")
            else:
                self.log(f"已内嵌 {embedded_count} 张图片（Base64）")
        self.log(f"已保存到 {target_file}")

    def save_file_as(self):
        default_name = f"{self._sanitize_workflow_file_name(self.workflow.get('name', 'workflow'))}.json"
        filename, accepted = QInputDialog.getText(
            self,
            "另存为",
            f"文件名（固定保存到 {self.WORKFLOW_DIR}）:",
            text=default_name,
        )
        if not accepted:
            return
        filename = filename.strip()
        if not filename:
            QMessageBox.warning(self, "保存失败", "文件名不能为空。")
            return
        if not filename.lower().endswith(".json"):
            filename = f"{filename}.json"

        target_file = self.WORKFLOW_DIR / filename
        current_target = self.workflow_files[self.workflow_index] if self.workflow_index < len(self.workflow_files) else None
        if target_file.exists() and (current_target is None or current_target.resolve() != target_file.resolve()):
            confirm = QMessageBox.question(
                self,
                "覆盖确认",
                f"{target_file.name} 已存在，是否覆盖？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if confirm != QMessageBox.Yes:
                return

        if self.workflow_index >= len(self.workflow_files):
            self.workflow_files.extend([None] * (self.workflow_index - len(self.workflow_files) + 1))
        self.workflow_files[self.workflow_index] = target_file
        self.current_file = str(target_file)
        self.save_file()

    def open_file(self):
        total = self._load_workflows_from_directory()
        self._save_session_state()
        self.log(f"已刷新工作流目录，共发现 {total} 个文件")

    def _save_session_state(self) -> None:
        state = {
            "default_delay": self.default_delay,
            "image_storage_mode": self.image_storage_mode,
            "auto_refresh_windows_each_step": self.auto_refresh_windows_each_step,
        }
        if self.current_file:
            state["last_workflow_file"] = str(Path(self.current_file))
        self.SESSION_STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

    def _load_session_state_path(self) -> Path | None:
        if not self.SESSION_STATE_FILE.exists():
            return None
        try:
            state = json.loads(self.SESSION_STATE_FILE.read_text(encoding="utf-8"))
            delay = str(state.get("default_delay", "")).strip()
            if delay:
                self.default_delay = delay
                set_default_delay(delay)
            storage_mode = str(state.get("image_storage_mode", self.IMAGE_STORAGE_BASE64_AUTO)).strip()
            if storage_mode not in {
                self.IMAGE_STORAGE_FILE,
                self.IMAGE_STORAGE_BASE64_AUTO,
                self.IMAGE_STORAGE_BASE64_ALWAYS,
            }:
                storage_mode = self.IMAGE_STORAGE_BASE64_AUTO
            self.image_storage_mode = storage_mode
            self.auto_refresh_windows_each_step = bool(state.get("auto_refresh_windows_each_step", False))
            self._sync_global_settings_widgets()
            path = state.get("last_workflow_file")
            if path:
                candidate = Path(path)
                if candidate.exists() and candidate.is_file():
                    return candidate
        except Exception:
            return None
        return None

    def _auto_load_last_workflow(self) -> None:
        preferred_path = self._load_session_state_path()
        total = self._load_workflows_from_directory(preferred_path=preferred_path)
        self._save_session_state()
        self.log(f"已加载工作流目录 {self.WORKFLOW_DIR}，共 {total} 个文件")

    def _start_runner_with_workflow(
        self,
        workflow_payload: dict,
        start_message: str,
        runner_node_index_map: list[int] | None = None,
    ) -> None:
        self.matched_windows_list.clear()
        self._matched_window_hwnds = []
        self.refresh_windows()
        self._runner_node_index_map = runner_node_index_map
        thread = QThread(self)
        runner = WorkflowRunner(
            clone_payload(workflow_payload),
            self.selected_hwnd,
            list(self.windows),
            refresh_windows_each_step=self.auto_refresh_windows_each_step,
        )
        runner.moveToThread(thread)
        thread.started.connect(runner.run)
        runner.log_emitted.connect(self.log)
        runner.windows_resolved.connect(self._on_runner_windows_resolved)
        runner.step_started.connect(self._on_runner_step_started)
        runner.finished.connect(thread.quit)
        runner.finished.connect(runner.deleteLater)
        thread.finished.connect(self._on_runner_thread_finished)
        thread.finished.connect(thread.deleteLater)

        self.runner_thread = thread
        self.runner = runner
        thread.start()
        self.log(start_message)

    def start_workflow(self):
        if self._is_runner_thread_running():
            self.log("Workflow is already running")
            return

        try:
            self.save_scope()
        except Exception as exc:
            QMessageBox.warning(self, "Validation Failed", str(exc))
            self.log(f"Pre-run validation failed: {exc}")
            return

        big_nodes = get_big_nodes(self.workflow)
        self._start_runner_with_workflow(
            self.workflow,
            "Workflow started",
            runner_node_index_map=list(range(len(big_nodes))),
        )

    def start_current_node(self):
        if self._is_runner_thread_running():
            self.log("Workflow is already running")
            return

        try:
            self.save_scope()
        except Exception as exc:
            QMessageBox.warning(self, "Validation Failed", str(exc))
            self.log(f"Pre-run validation failed: {exc}")
            return

        current_node = clone_payload(self.node)
        current_node["flow"] = {"mode": "stop"}
        workflow_payload = {
            "name": f"{self.workflow.get('name', 'workflow')}-single-node",
            "big_nodes": [current_node],
        }
        node_alias = self._big_node_display_alias(self.node_index, self.node)
        self._start_runner_with_workflow(
            workflow_payload,
            f"开始运行当前大节点: {node_alias}",
            runner_node_index_map=[self.node_index],
        )

    def _build_test_scope(self) -> dict:
        scope_regex = self.test_scope_regex_edit.text().strip()
        if scope_regex:
            try:
                re.compile(scope_regex)
            except re.error as exc:
                raise ValueError(f"测试窗口正则无效: {exc}") from exc
        return {
            "regex": scope_regex,
            "bring_front": True,
            "multi_window_mode": "first",
        }

    def _run_detached_test_action(self, action_name: str) -> None:
        if action_name in self.WINDOW_REQUIRED_TEST_ACTIONS and not self.test_scope_regex_edit.text().strip():
            QMessageBox.warning(self, "测试失败", f"{action_name} 需要目标窗口，请填写“测试窗口正则”。")
            return

        test_step = create_small_node(action_name)
        edited_step = edit_step_dialog(self, test_step)
        if edited_step is None:
            self.log("已取消测试")
            return

        scope = self._build_test_scope()
        test_node = {
            "alias": "测试节点",
            "scope": scope,
            "small_nodes": [edited_step],
        }
        test_workflow = {
            "name": f"动作测试-{action_name}",
            "big_nodes": [test_node],
        }
        self._start_runner_with_workflow(test_workflow, f"开始测试执行：{action_name}")

    def test_action(self):
        if self._is_runner_thread_running():
            self.log("Workflow is already running")
            return

        action_name = self.test_action_combo.currentData()
        if not isinstance(action_name, str) or action_name not in self.TEST_ACTIONS:
            self.log("未选择有效的动作")
            return

        try:
            self._run_detached_test_action(action_name)
        except Exception as exc:
            QMessageBox.warning(self, "测试失败", str(exc))
            self.log(f"测试失败: {exc}")

    def test_current_step(self):
        # Backward-compatible wrapper for old signal/method name.
        self.test_action()

    def test_detection_action(self):
        if self._is_runner_thread_running():
            self.log("Workflow is already running")
            return

        action_name = self.detect_action_combo.currentData()
        if not isinstance(action_name, str) or action_name not in self.DETECTION_ACTIONS:
            self.log("未选择有效的检测动作")
            return

        try:
            self._run_detached_test_action(action_name)
        except Exception as exc:
            QMessageBox.warning(self, "检测测试失败", str(exc))
            self.log(f"检测测试失败: {exc}")

    def stop_workflow(self):
        if self.runner:
            try:
                self.runner.stop()
            except RuntimeError:
                self.runner = None
        self.log("Stop signal sent")

    def closeEvent(self, event):
        if self._is_runner_thread_running():
            self.stop_workflow()
            if self.runner_thread is not None:
                self.runner_thread.quit()
                self.runner_thread.wait(3000)
        super().closeEvent(event)



