import json
import re
import time
from pathlib import Path
from uuid import uuid4

from PySide6.QtCore import Qt, QThread
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTextEdit,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..constants import ACTION_CLICK_ABS, ACTION_CLICK_IMAGE, ACTION_CLICK_REL, ACTION_DETECT_IMAGE, ACTION_DEFINITIONS
from ..models import (
    clone_payload,
    create_big_node,
    create_small_node,
    create_workflow,
    get_default_delay,
    get_big_nodes,
    get_small_nodes,
    normalize_workflow,
    set_default_delay,
)
from ..services.runner import WorkflowRunner
from ..services.windows import enumerate_windows, get_window_frame_rect, resolve_scope_window
from .dialogs import edit_big_node_dialog, edit_step_dialog
from .overlays import PointPickerOverlay, ScreenshotOverlay
from .panels import ActionPalette, BigNodeListWidget, StepListWidget


class MainWindow(QMainWindow):
    SESSION_STATE_FILE = Path.cwd() / ".macro_app_state.json"

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Macro App")
        self.resize(1560, 920)

        self.windows = []
        self.workflows = [create_workflow()]
        self.workflow_index = 0
        self.node_index = 0
        self.step_index = 0
        self.current_file = None
        self.selected_hwnd = None
        self.runner = None
        self.runner_thread = None
        self._matched_window_hwnds: list[int] = []
        self._active_point_overlay = None
        self._active_screenshot_overlay = None
        self.default_delay = get_default_delay()

        self._build_ui()
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
            return "????"
        if mode == "serial":
            return "????"
        return "?????"

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

        header = QHBoxLayout()
        title = QLabel("工作流控制台")
        title.setStyleSheet("font-size: 20px; font-weight: 700;")
        self.status_label = QLabel("准备就绪")
        self.status_label.setStyleSheet("color: #667085;")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(self.status_label)
        layout.addLayout(header)

        toolbar = QHBoxLayout()
        for text, callback in [
            ("新建", self.new_workflow),
            ("打开", self.open_file),
            ("保存", self.save_file),
            ("另存为", self.save_file_as),
            ("刷新窗口", self.refresh_windows),
            ("开始", self.start_workflow),
            ("停止", self.stop_workflow),
        ]:
            button = QPushButton(text)
            button.clicked.connect(callback)
            toolbar.addWidget(button)
        toolbar.addStretch(1)
        layout.addLayout(toolbar)

        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter, 1)

        splitter.addWidget(self._build_left_panel())
        splitter.addWidget(self._build_center_panel())
        splitter.addWidget(self._build_right_panel())
        splitter.setSizes([280, 700, 400])

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
        workflow_row = QHBoxLayout()
        add_button = QPushButton("New")
        add_button.clicked.connect(self.new_workflow)
        delete_button = QPushButton("删除")
        delete_button.clicked.connect(self.delete_workflow)
        workflow_row.addWidget(add_button)
        workflow_row.addWidget(delete_button)
        workflow_row.addStretch(1)
        workflow_layout.addLayout(workflow_row)
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
        self.node_list = BigNodeListWidget()
        self.node_list.row_changed.connect(self.on_node_selected)
        self.node_list.node_double_clicked.connect(self.edit_big_node_quickly)
        node_layout.addWidget(self.node_list)
        node_row = QHBoxLayout()
        for text, callback in [("新增", self.add_node), ("删除", self.delete_node), ("上移", self.move_node_up), ("下移", self.move_node_down)]:
            button = QPushButton(text)
            button.clicked.connect(callback)
            node_row.addWidget(button)
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
        header.addStretch(1)
        steps_layout.addLayout(header)
        self.step_list = StepListWidget()
        self.step_list.row_changed.connect(self.on_step_selected)
        self.step_list.step_double_clicked.connect(self.edit_step_quickly)
        self.step_list.action_dropped.connect(self.insert_step_from_drop)
        self.step_list.header().setStretchLastSection(True)
        steps_layout.addWidget(self.step_list, 1)
        step_row = QHBoxLayout()
        for text, callback in [("复制", self.copy_step), ("删除", self.delete_step), ("上移", self.move_step_up), ("下移", self.move_step_down)]:
            button = QPushButton(text)
            button.clicked.connect(callback)
            step_row.addWidget(button)
        step_row.addStretch(1)
        steps_layout.addLayout(step_row)
        right_col_layout.addWidget(steps_box, 4)

        palette_box = self._create_section("Add Actions")
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
        self.overview_text.setFixedHeight(100)
        overview_layout.addWidget(self.overview_text)
        layout.addWidget(overview_box)

        matched_box = self._create_section("匹配窗口")
        matched_layout = matched_box.layout()
        self.matched_windows_list = QListWidget()
        matched_layout.addWidget(self.matched_windows_list)
        layout.addWidget(matched_box)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text, 1)
        clear_button = QPushButton("清空日志")
        clear_button.clicked.connect(self.log_text.clear)
        layout.addWidget(clear_button, alignment=Qt.AlignRight)
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
        self.log_text.append(f"[{timestamp}] {message}")
        self.status_label.setText(message)

    def refresh_windows(self):
        self.windows = enumerate_windows()
        self.log(f"刷新窗口列表，共 {len(self.windows)} 个")

    def refresh_all(self):
        self.workflow_index = min(self.workflow_index, len(self.workflows) - 1)
        self.workflow_list.blockSignals(True)
        self.workflow_list.clear()
        for workflow in self.workflows:
            self.workflow_list.addItem(workflow["name"])
        self.workflow_list.setCurrentRow(self.workflow_index)
        self.workflow_list.blockSignals(False)

        big_nodes = get_big_nodes(self.workflow)
        self.node_index = min(self.node_index, len(big_nodes) - 1)
        self.node_list.blockSignals(True)
        self.node_list.clear()
        for index, node in enumerate(big_nodes, start=1):
            scope = node.get("scope", {})
            item = QTreeWidgetItem(
                [
                    self._big_node_display_alias(index - 1, node),
                    scope.get("regex", ""),
                    self._multi_window_mode_label(scope.get("multi_window_mode", "first")),
                ]
            )
            self.node_list.addTopLevelItem(item)
        if self.node_list.topLevelItemCount() > 0:
            self.node_list.setCurrentItem(self.node_list.topLevelItem(self.node_index))
        self.node_list.blockSignals(False)
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
        self.overview_text.setPlainText(
            "\n".join(
                [f"工作流: {self.workflow['name']}"]
                + [f"{self._big_node_display_alias(index, node)} -> {len(get_small_nodes(node))} 个小节点" for index, node in enumerate(big_nodes)]
            )
        )

    def on_workflow_selected(self, index: int):
        if index < 0:
            return
        self.workflow_index = index
        self.node_index = 0
        self.step_index = 0
        self.refresh_all()

    def on_node_selected(self, index: int):
        if index < 0:
            return
        self.node_index = index
        self.step_index = 0
        self.refresh_steps()

    def on_step_selected(self, index: int):
        if index < 0:
            return
        self.step_index = index

    def _on_runner_windows_resolved(self, payload: dict):
        node_index = int(payload.get("node_index", 0))
        windows = payload.get("windows", []) or []

        # Switch UI context to the running node so step highlight is visible.
        if node_index != self.node_index:
            self.node_index = max(0, min(node_index, len(get_big_nodes(self.workflow)) - 1))
            self.step_index = 0
            self.refresh_steps()

        self.matched_windows_list.clear()
        self._matched_window_hwnds = []
        if not windows:
            self.matched_windows_list.addItem("(未匹配到窗口)")
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

    def _on_runner_step_started(self, payload: dict):
        node_index = int(payload.get("node_index", 0))
        step_index = int(payload.get("step_index", 0))
        hwnd = payload.get("hwnd", None)

        if node_index != self.node_index:
            self.node_index = max(0, min(node_index, len(get_big_nodes(self.workflow)) - 1))
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
        self.workflow_index = len(self.workflows) - 1
        self.node_index = 0
        self.step_index = 0
        self.refresh_all()
        self.log("已新建工作流")

    def delete_workflow(self):
        if len(self.workflows) == 1:
            QMessageBox.information(self, "不能删除", "至少保留一个工作流。")
            return
        self.workflows.pop(self.workflow_index)
        self.workflow_index = max(0, self.workflow_index - 1)
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
        for index, node in enumerate(get_big_nodes(self.workflow), start=1):
            scope = node.get("scope", {})
            regex = (scope.get("regex") or "").strip()
            if not regex:
                continue
            try:
                re.compile(regex)
            except re.error as exc:
                raise ValueError(f"大节点 {index} 作用域正则无效: {exc}") from exc

    def load_scope(self):
        # 预留扩展点：如果未来增加“窗口作用域面板”，可在这里统一回填 UI。
        return

    def insert_step(self, action_name: str, index: int | None = None):
        if index is None:
            small_nodes = get_small_nodes(self.node)
            small_nodes.append(create_small_node(action_name))
            self.step_index = len(small_nodes) - 1
        else:
            small_nodes = get_small_nodes(self.node)
            index = max(0, min(index, len(small_nodes)))
            small_nodes.insert(index, create_small_node(action_name))
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

        if action_name in {ACTION_CLICK_ABS, ACTION_CLICK_REL}:
            overlay = PointPickerOverlay(rel_provider=self._create_rel_provider())
            self._active_point_overlay = overlay

            def on_point_picked(abs_x: int, abs_y: int):
                if action_name == ACTION_CLICK_REL:
                    rel_provider = self._create_rel_provider()
                    rel_pos = rel_provider(abs_x, abs_y) if rel_provider else None
                    if rel_pos is not None:
                        params["x"], params["y"] = str(rel_pos[0]), str(rel_pos[1])
                    else:
                        params["x"], params["y"] = str(abs_x), str(abs_y)
                        self.log("No preview window for relative conversion; stored absolute values temporarily.")
                else:
                    params["x"], params["y"] = str(abs_x), str(abs_y)
                self.refresh_steps()
                self.log(f"Point captured: ({params.get('x')}, {params.get('y')})")

            def on_canceled():
                self.log("Point capture canceled.")

            overlay.point_picked.connect(on_point_picked)
            overlay.canceled.connect(on_canceled)
            overlay.show()
            overlay.activateWindow()
            return

        if action_name in {ACTION_CLICK_IMAGE, ACTION_DETECT_IMAGE}:
            templates_dir = Path.cwd() / "templates"
            default_target = templates_dir / f"template_{time.strftime('%Y%m%d_%H%M%S')}_{uuid4().hex[:8]}.png"
            overlay = ScreenshotOverlay(default_target=str(default_target), prompt_for_save=False)
            self._active_screenshot_overlay = overlay

            def on_screenshot_saved(path: str):
                params["image_path"] = path
                self.refresh_steps()
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


    def save_file(self):
        if self.current_file is None:
            self.save_file_as()
            return
        Path(self.current_file).write_text(
            json.dumps({"workflows": self.workflows}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        self._save_session_state()
        self.log(f"Saved to {self.current_file}")

    def save_file_as(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Workflow",
            str(Path.cwd() / f"{self.workflow['name']}.json"),
            "JSON Files (*.json)",
        )
        if not path:
            return
        self.current_file = path
        self.save_file()

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Workflow", "", "JSON Files (*.json)")
        if not path:
            return
        try:
            self._open_workflow_path(Path(path))
        except Exception as exc:
            QMessageBox.critical(self, "Open Failed", str(exc))

    def _open_workflow_path(self, path: Path) -> None:
        payload = json.loads(path.read_text(encoding="utf-8"))
        workflows = payload.get("workflows")
        if not isinstance(workflows, list) or not workflows:
            raise ValueError("Invalid workflow file: missing workflows list")

        self.workflows = [normalize_workflow(item) for item in workflows]
        self.current_file = str(path)
        self.workflow_index = 0
        self.node_index = 0
        self.step_index = 0
        self.refresh_all()
        self._save_session_state()
        self.log(f"Loaded workflow file: {path}")

    def _save_session_state(self) -> None:
        state = {"default_delay": self.default_delay}
        if self.current_file:
            state["last_workflow_file"] = str(Path(self.current_file))
        self.SESSION_STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

    def _load_session_state_path(self) -> Path | None:
        if not self.SESSION_STATE_FILE.exists():
            return self._discover_workflow_file()
        try:
            state = json.loads(self.SESSION_STATE_FILE.read_text(encoding="utf-8"))
            delay = str(state.get("default_delay", "")).strip()
            if delay:
                self.default_delay = delay
                set_default_delay(delay)
            path = state.get("last_workflow_file")
            if not path:
                return self._discover_workflow_file()
            candidate = Path(path)
            if candidate.exists() and candidate.is_file():
                return candidate
        except Exception:
            return self._discover_workflow_file()
        return self._discover_workflow_file()

    def _discover_workflow_file(self) -> Path | None:
        candidates = [item for item in Path.cwd().glob("*.json") if item.name != self.SESSION_STATE_FILE.name]
        if not candidates:
            return None
        candidates.sort(key=lambda item: item.stat().st_mtime, reverse=True)
        return candidates[0]

    def _auto_load_last_workflow(self) -> None:
        path = self._load_session_state_path()
        if path is None:
            return
        try:
            self._open_workflow_path(path)
            self.log(f"Auto-loaded workflow: {path}")
        except Exception as exc:
            self.log(f"Auto-load failed: {exc}")

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

        self.matched_windows_list.clear()
        self._matched_window_hwnds = []
        self.refresh_windows()
        thread = QThread(self)
        runner = WorkflowRunner(clone_payload(self.workflow), self.selected_hwnd, list(self.windows))
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
        self.log("Workflow started")

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


