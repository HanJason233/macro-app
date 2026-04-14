from __future__ import annotations

import shlex
import subprocess
import time
from pathlib import Path
from typing import Any

import pyautogui
from PySide6.QtCore import QObject, Signal

from ..constants import (
    ACTION_CLICK,
    ACTION_CLICK_IMAGE,
    ACTION_CLICK_OCR,
    ACTION_CLOSE_PROGRAM,
    ACTION_DETECT_IMAGE,
    ACTION_DETECT_OCR,
    ACTION_GET_WINDOW_SIZE,
    ACTION_DETECT_WINDOW_SIZE,
    ACTION_KEY_PRESS,
    ACTION_MOUSE_DRAG,
    ACTION_MOUSE_SCROLL,
    ACTION_MINIMIZE_WINDOW,
    ACTION_RESIZE_WINDOW,
    ACTION_START_PROGRAM,
    ACTION_TEXT_INPUT,
    normalize_action_name,
    summarize_step,
)
from ..models import get_big_nodes, get_small_nodes
from .capture import match_in_window
from .ocr import find_text_in_window
from .windows import (
    bring_window_to_front,
    close_window,
    enumerate_windows,
    get_window_frame_rect,
    get_window_rect,
    minimize_window,
    resize_window,
    resolve_scope_windows,
)

pyautogui.FAILSAFE = True

ALLOWED_MOUSE_BUTTONS = {"left", "right", "middle"}


def _num(value: Any, name: str, integer: bool = False, minimum: float | None = None) -> int | float:
    try:
        number = int(value) if integer else float(value)
    except Exception as exc:
        raise ValueError(f"{name} 必须是数字") from exc

    if minimum is not None and number < minimum:
        raise ValueError(f"{name} 不能小于 {minimum}")
    return number


def _mouse_button(value: str) -> str:
    button = (value or "left").strip().lower()
    if button not in ALLOWED_MOUSE_BUTTONS:
        raise ValueError(f"不支持的鼠标按键: {button}")
    return button


def _coord_mode(value: str, scope: dict | None = None) -> str:
    raw = (value or "").strip().lower()
    if raw in {"absolute", "abs", "绝对", "绝对坐标"}:
        return "absolute"
    scope_regex = ""
    if isinstance(scope, dict):
        scope_regex = str(scope.get("regex", "")).strip()
    return "relative" if scope_regex else "absolute"


def _parse_program_command(raw: str) -> list[str]:
    command = raw.strip()
    if not command:
        raise RuntimeError("程序路径不能为空")

    args = shlex.split(command, posix=False)
    if not args:
        raise RuntimeError("程序路径不能为空")

    executable = Path(args[0].strip('"')).expanduser()
    if executable.exists():
        args[0] = str(executable)
    return args


class WorkflowRunner(QObject):
    log_emitted = Signal(str)
    windows_resolved = Signal(object)
    step_started = Signal(object)
    finished = Signal()

    def __init__(
        self,
        workflow: dict,
        selected_hwnd: int | None,
        windows_snapshot: list[dict],
        refresh_windows_each_step: bool = False,
    ):
        super().__init__()
        self.workflow = workflow
        self.selected_hwnd = selected_hwnd
        self.windows_snapshot = windows_snapshot
        self.refresh_windows_each_step = bool(refresh_windows_each_step)
        self.should_stop = False

    def stop(self):
        self.should_stop = True

    def run(self):
        try:
            nodes = get_big_nodes(self.workflow)
            node_count = len(nodes)
            node_index = 0
            flow_guard = 0
            loop_counts: dict[int, int] = {}
            while 0 <= node_index < node_count:
                if self.should_stop:
                    self.log_emitted.emit("工作流已停止")
                    return

                flow_guard += 1
                if flow_guard > 10000:
                    raise RuntimeError("大节点流转次数超过上限，可能存在死循环")

                node = nodes[node_index]
                node_name = (node.get("alias") or node.get("name") or "大节点").strip()
                self.log_emitted.emit(f"进入大节点: {node_name}")

                # Refresh window cache before each big-node scope resolution so matching stays up-to-date.
                self.windows_snapshot = enumerate_windows()
                scope = node.get("scope", {})
                windows = resolve_scope_windows(scope, self.selected_hwnd, self.windows_snapshot)
                if scope.get("regex") and not windows:
                    raise RuntimeError(f"{node_name} 未匹配到目标窗口")

                mode = str(scope.get("multi_window_mode", "first") or "first").strip().lower()
                if mode not in {"first", "sync", "serial"}:
                    mode = "first"

                self.windows_resolved.emit(
                    {
                        "node_index": node_index,
                        "mode": mode,
                        "windows": windows,
                    }
                )

                if mode == "first":
                    targets = windows[:1] if windows else [None]
                    last_detected = self._run_serial(node, node_index, scope, targets, mode="first")
                elif mode == "sync":
                    targets = windows if windows else [None]
                    last_detected = self._run_sync(node, node_index, scope, targets)
                else:
                    targets = windows if windows else [None]
                    last_detected = self._run_serial(node, node_index, scope, targets, mode="serial")

                next_index = self._resolve_next_node_index(
                    node,
                    node_index,
                    node_count,
                    last_detected,
                    targets,
                    loop_counts,
                )
                if next_index != node_index and 0 <= next_index < node_count:
                    node_delay = float(_num(node.get("node_delay", "0"), "大节点间延迟", minimum=0))
                    if node_delay > 0:
                        self.log_emitted.emit(f"大节点间延迟 {node_delay}s")
                        if self._sleep_with_stop(node_delay):
                            self.log_emitted.emit("工作流已停止")
                            return
                if next_index != node_index:
                    loop_counts[node_index] = 0
                node_index = next_index

            self.log_emitted.emit("工作流执行完成")
        except Exception as exc:
            self.log_emitted.emit(f"执行失败: {exc}")
        finally:
            self.finished.emit()

    def _resolve_next_node_index(
        self,
        node: dict,
        node_index: int,
        node_count: int,
        last_detected: bool | None,
        targets: list[dict | None],
        loop_counts: dict[int, int],
    ) -> int:
        flow = node.get("flow", {})
        mode = str(flow.get("mode", "next") or "next").strip().lower()
        target_raw = str(flow.get("target", "") or "").strip()
        condition = str(flow.get("condition", "last_detected") or "last_detected").strip().lower()
        max_loops_raw = str(flow.get("max_loops", "1") or "1").strip()
        try:
            max_loops = max(1, int(max_loops_raw))
        except Exception:
            max_loops = 1

        def parse_target() -> int:
            target = int(target_raw)
            if target < 1 or target > node_count:
                raise RuntimeError(f"大节点 {node_index + 1} 跳转目标超出范围: {target}")
            return target - 1

        if mode == "loop":
            loop_counts[node_index] = loop_counts.get(node_index, 0) + 1
            if loop_counts[node_index] < max_loops:
                self.log_emitted.emit(f"大节点流转: 循环第 {loop_counts[node_index] + 1}/{max_loops} 次")
                return node_index
            return node_index + 1

        if mode == "jump":
            if not target_raw:
                raise RuntimeError(f"大节点 {node_index + 1} 未配置跳转目标")
            target_index = parse_target()
            self.log_emitted.emit(f"大节点流转: 无条件跳转到 {target_index + 1}")
            return target_index

        if mode == "conditional_jump":
            if not target_raw:
                raise RuntimeError(f"大节点 {node_index + 1} 未配置条件跳转目标")
            condition_met = False
            if condition == "last_detected":
                condition_met = last_detected is True
            elif condition == "last_not_detected":
                condition_met = last_detected is False
            elif condition == "window_text_detected":
                condition_met = self._evaluate_flow_text_condition(flow, targets)
            elif condition == "window_image_detected":
                condition_met = self._evaluate_flow_image_condition(flow, targets)
            if condition_met:
                target_index = parse_target()
                self.log_emitted.emit(f"大节点流转: 条件满足，跳转到 {target_index + 1}")
                return target_index
            self.log_emitted.emit("大节点流转: 条件不满足，顺序执行下一个")
            return node_index + 1

        if mode == "stop":
            self.log_emitted.emit("大节点流转: 当前大节点执行完毕，按配置停止工作流")
            return node_count

        return node_index + 1

    def _evaluate_flow_text_condition(self, flow: dict, targets: list[dict | None]) -> bool:
        target_text = str(flow.get("condition_target_text", "") or "").strip()
        if not target_text:
            raise RuntimeError("条件跳转(窗口文字)缺少条件文字")
        text_match_mode = str(flow.get("condition_text_match_mode", "contains") or "contains").strip().lower()
        min_score = float(_num(flow.get("condition_min_score", "0.5"), "条件最低置信度", minimum=0))
        for window in targets:
            if not window:
                continue
            matches = find_text_in_window(
                int(window["hwnd"]),
                target_text,
                text_match_mode,
                min_score,
            )
            if matches:
                return True
        return False

    def _evaluate_flow_image_condition(self, flow: dict, targets: list[dict | None]) -> bool:
        image_path = str(flow.get("condition_image_path", "") or "").strip()
        if not image_path:
            raise RuntimeError("条件跳转(窗口图片)缺少条件图片路径")
        match_mode = str(flow.get("condition_match_mode", "default") or "default").strip().lower()
        threshold = float(_num(flow.get("condition_threshold", "0.8"), "条件图片阈值", minimum=0))
        for window in targets:
            if not window:
                continue
            points = match_in_window(
                int(window["hwnd"]),
                image_path,
                match_mode,
                threshold,
            )
            if points:
                return True
        return False

    def _refresh_windows_for_step(
        self,
        node_index: int,
        scope: dict,
        mode: str,
        current_window: dict | None = None,
    ) -> tuple[list[dict | None], dict | None]:
        self.windows_snapshot = enumerate_windows()
        windows = resolve_scope_windows(scope, self.selected_hwnd, self.windows_snapshot)
        if scope.get("regex") and not windows:
            raise RuntimeError(f"大节点 {node_index + 1} 小节点执行前未匹配到目标窗口")
        self.windows_resolved.emit(
            {
                "node_index": node_index,
                "mode": mode,
                "windows": windows,
            }
        )
        if mode == "first":
            targets: list[dict | None] = windows[:1] if windows else [None]
        else:
            targets = windows if windows else [None]

        refreshed_window: dict | None = None
        if current_window:
            current_hwnd = int(current_window.get("hwnd", 0))
            refreshed_window = next(
                (item for item in windows if int(item.get("hwnd", 0)) == current_hwnd),
                None,
            )
        if refreshed_window is None:
            refreshed_window = targets[0] if targets else None
        return targets, refreshed_window

    def _run_serial(
        self,
        node: dict,
        node_index: int,
        scope: dict,
        targets: list[dict | None],
        mode: str,
    ) -> bool | None:
        last_detected: bool | None = None
        for window in targets:
            if window:
                self.log_emitted.emit(f"目标窗口: {window['title']}")
            for step_index, step in enumerate(get_small_nodes(node)):
                if self.should_stop:
                    self.log_emitted.emit("工作流已停止")
                    return last_detected
                if self.refresh_windows_each_step:
                    _, window = self._refresh_windows_for_step(
                        node_index=node_index,
                        scope=scope,
                        mode=mode,
                        current_window=window,
                    )
                self._emit_step_started(node_index, step_index, window)
                step_detected = self._execute_step(step, scope, window)
                if step_detected is not None:
                    last_detected = step_detected
                delay = float(_num(step.get("delay", "2"), "小节点延迟", minimum=0))
                if self._sleep_with_stop(delay):
                    self.log_emitted.emit("工作流已停止")
                    return last_detected
        return last_detected

    def _run_sync(self, node: dict, node_index: int, scope: dict, targets: list[dict | None]) -> bool | None:
        last_detected: bool | None = None
        for step_index, step in enumerate(get_small_nodes(node)):
            if self.should_stop:
                self.log_emitted.emit("工作流已停止")
                return last_detected
            if self.refresh_windows_each_step:
                targets, _ = self._refresh_windows_for_step(
                    node_index=node_index,
                    scope=scope,
                    mode="sync",
                )

            for window in targets:
                if self.should_stop:
                    self.log_emitted.emit("工作流已停止")
                    return last_detected
                self._emit_step_started(node_index, step_index, window)
                step_detected = self._execute_step(step, scope, window)
                if step_detected is not None:
                    last_detected = step_detected

            delay = float(_num(step.get("delay", "2"), "小节点延迟", minimum=0))
            if self._sleep_with_stop(delay):
                self.log_emitted.emit("工作流已停止")
                return last_detected
        return last_detected

    def _emit_step_started(self, node_index: int, step_index: int, window: dict | None) -> None:
        self.step_started.emit(
            {
                "node_index": node_index,
                "step_index": step_index,
                "hwnd": window["hwnd"] if window else None,
                "window_title": window["title"] if window else "",
            }
        )

    def _sleep_with_stop(self, seconds: float) -> bool:
        end_at = time.time() + max(0.0, seconds)
        while time.time() < end_at:
            if self.should_stop:
                return True
            time.sleep(0.05)
        return False

    def _poll_until(self, duration: float, interval: float, check: Any) -> bool | None:
        deadline = time.monotonic() + max(0.0, duration)
        while True:
            result = check()
            if result:
                return True
            if self.should_stop:
                return None
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return False
            if self._sleep_with_stop(min(interval, remaining)):
                return None

    @staticmethod
    def _require_window(window: dict | None, action: str) -> dict:
        if not window:
            raise RuntimeError(f"{action} 需要目标窗口")
        return window

    def _resolve_point(self, window: dict | None, action: str, x: int, y: int, mode: str) -> tuple[int, int]:
        if mode == "absolute":
            return x, y
        target = self._require_window(window, action)
        left, top, _right, _bottom = get_window_frame_rect(target["hwnd"])
        return left + x, top + y

    def _execute_step(self, step: dict, scope: dict, window: dict | None) -> bool | None:
        action = normalize_action_name(step.get("action", ACTION_TEXT_INPUT))
        params = step.get("params", {})
        self.log_emitted.emit(f"执行 {summarize_step(step)}")

        if window and scope.get("bring_front", True):
            bring_window_to_front(window["hwnd"])
            time.sleep(0.2)

        if action == ACTION_KEY_PRESS:
            repeat = int(_num(params.get("repeat", "1"), "重复次数", integer=True, minimum=1))
            interval = float(_num(params.get("interval", "0.1"), "按键间隔", minimum=0))
            key_name = (params.get("key", "enter") or "enter").strip()
            for _ in range(repeat):
                pyautogui.press(key_name)
                if self._sleep_with_stop(interval):
                    return
            return

        if action == ACTION_TEXT_INPUT:
            interval = float(_num(params.get("interval", "0.02"), "字间隔", minimum=0))
            pyautogui.write(str(params.get("text", "")), interval=interval)
            return

        if action == ACTION_CLICK:
            mode = _coord_mode(str(params.get("coord_mode", "")), scope)
            x = int(_num(params.get("x", "0"), "X", integer=True))
            y = int(_num(params.get("y", "0"), "Y", integer=True))
            click_x, click_y = self._resolve_point(window, action, x, y, mode)
            pyautogui.click(
                x=click_x,
                y=click_y,
                button=_mouse_button(str(params.get("button", "left"))),
                clicks=int(_num(params.get("clicks", "1"), "点击次数", integer=True, minimum=1)),
            )
            return

        if action == ACTION_CLICK_IMAGE:
            target = self._require_window(window, action)
            trigger_raw = str(params.get("click_trigger", "检测到")).strip().lower()
            if trigger_raw in {"检测到", "detected", "found"}:
                trigger_mode = "detected"
            elif trigger_raw in {"检测不到", "not_detected", "not found", "not_found"}:
                trigger_mode = "not_detected"
            else:
                raise ValueError("检测结果触发仅支持：检测到 / 检测不到")

            points = match_in_window(
                target["hwnd"],
                str(params.get("image_path", "")),
                str(params.get("match_mode", "default")),
                float(_num(params.get("threshold", "0.8"), "阈值", minimum=0)),
            )
            if not points:
                if trigger_mode == "detected":
                    raise RuntimeError("未识别到目标图片")
                pyautogui.click(
                    button=_mouse_button(str(params.get("button", "left"))),
                    clicks=int(_num(params.get("clicks", "1"), "点击次数", integer=True, minimum=1)),
                )
                self.log_emitted.emit("未识别到目标图片，按“检测不到”条件执行点击")
                return
            if trigger_mode == "not_detected":
                self.log_emitted.emit("已识别到目标图片，当前为“检测不到”触发，跳过点击")
                return

            x, y, width, height = points[0]
            left, top, _right, _bottom = get_window_rect(target["hwnd"])
            pyautogui.click(
                x=left + x + width // 2 + int(_num(params.get("offset_x", "0"), "偏移X", integer=True)),
                y=top + y + height // 2 + int(_num(params.get("offset_y", "0"), "偏移Y", integer=True)),
                button=_mouse_button(str(params.get("button", "left"))),
                clicks=int(_num(params.get("clicks", "1"), "点击次数", integer=True, minimum=1)),
            )
            return

        if action == ACTION_CLICK_OCR:
            target = self._require_window(window, action)
            matches = find_text_in_window(
                target["hwnd"],
                str(params.get("target_text", "")),
                str(params.get("text_match_mode", "contains")),
                float(_num(params.get("min_score", "0.5"), "最低置信度", minimum=0)),
            )
            if not matches:
                raise RuntimeError("未识别到目标文字")

            item = matches[0]
            left, top, _right, _bottom = get_window_rect(target["hwnd"])
            pyautogui.click(
                x=left + int(item["x"]) + int(item["width"]) // 2 + int(_num(params.get("offset_x", "0"), "偏移X", integer=True)),
                y=top + int(item["y"]) + int(item["height"]) // 2 + int(_num(params.get("offset_y", "0"), "偏移Y", integer=True)),
                button=_mouse_button(str(params.get("button", "left"))),
                clicks=int(_num(params.get("clicks", "1"), "点击次数", integer=True, minimum=1)),
            )
            return

        if action == ACTION_MOUSE_DRAG:
            mode = _coord_mode(str(params.get("coord_mode", "")), scope)
            start_x = int(_num(params.get("start_x", "500"), "起点X", integer=True))
            start_y = int(_num(params.get("start_y", "500"), "起点Y", integer=True))
            end_x = int(_num(params.get("end_x", "700"), "终点X", integer=True))
            end_y = int(_num(params.get("end_y", "700"), "终点Y", integer=True))
            hold_duration = float(_num(params.get("hold_duration", "0.2"), "按住时长", minimum=0))
            move_duration = float(_num(params.get("move_duration", "0.4"), "拖动时长", minimum=0))
            drag_start_x, drag_start_y = self._resolve_point(window, action, start_x, start_y, mode)
            drag_end_x, drag_end_y = self._resolve_point(window, action, end_x, end_y, mode)
            self.log_emitted.emit(f"拖动坐标: ({drag_start_x}, {drag_start_y}) -> ({drag_end_x}, {drag_end_y})")

            pyautogui.moveTo(drag_start_x, drag_start_y)
            pyautogui.mouseDown(button="left")
            try:
                if self._sleep_with_stop(hold_duration):
                    return
                pyautogui.moveTo(drag_end_x, drag_end_y, duration=move_duration)
            finally:
                pyautogui.mouseUp(button="left")
            return

        if action == ACTION_MOUSE_SCROLL:
            mode = _coord_mode(str(params.get("coord_mode", "")), scope)
            start_x = int(_num(params.get("start_x", params.get("x", "500")), "起点X", integer=True))
            start_y = int(_num(params.get("start_y", params.get("y", "500")), "起点Y", integer=True))
            end_x = int(_num(params.get("end_x", params.get("x", "700")), "终点X", integer=True))
            end_y = int(_num(params.get("end_y", params.get("y", "700")), "终点Y", integer=True))
            move_duration = float(_num(params.get("move_duration", "0.4"), "滑动时长", minimum=0))
            slide_start_x, slide_start_y = self._resolve_point(window, action, start_x, start_y, mode)
            slide_end_x, slide_end_y = self._resolve_point(window, action, end_x, end_y, mode)
            self.log_emitted.emit(f"滑动坐标: ({slide_start_x}, {slide_start_y}) -> ({slide_end_x}, {slide_end_y})")
            pyautogui.moveTo(slide_start_x, slide_start_y)
            pyautogui.moveTo(slide_end_x, slide_end_y, duration=move_duration)
            return

        if action == ACTION_DETECT_IMAGE:
            if not window:
                self.log_emitted.emit("没检测到")
                return False

            duration = float(_num(params.get("detect_duration", "0"), "检测时长", minimum=0))
            interval = float(_num(params.get("detect_interval", "0.5"), "检测频率", minimum=0))

            def check_image() -> bool:
                points = match_in_window(
                    window["hwnd"],
                    str(params.get("image_path", "")),
                    str(params.get("match_mode", "default")),
                    float(_num(params.get("threshold", "0.8"), "阈值", minimum=0)),
                )
                return bool(points)

            if duration > 0:
                detected = self._poll_until(duration, interval, check_image)
                if detected is None:
                    return None
            else:
                detected = check_image()

            if detected:
                self.log_emitted.emit("检测到")
            else:
                self.log_emitted.emit("没检测到")
            return bool(detected)

        if action == ACTION_DETECT_OCR:
            if not window:
                self.log_emitted.emit("没检测到")
                return False

            duration = float(_num(params.get("detect_duration", "0"), "检测时长", minimum=0))
            interval = float(_num(params.get("detect_interval", "0.5"), "检测频率", minimum=0))

            def check_text() -> bool:
                matches = find_text_in_window(
                    window["hwnd"],
                    str(params.get("target_text", "")),
                    str(params.get("text_match_mode", "contains")),
                    float(_num(params.get("min_score", "0.5"), "最低置信度", minimum=0)),
                )
                return bool(matches)

            if duration > 0:
                detected = self._poll_until(duration, interval, check_text)
                if detected is None:
                    return None
            else:
                detected = check_text()

            if detected:
                self.log_emitted.emit("检测到")
            else:
                self.log_emitted.emit("没检测到")
            return bool(detected)

        if action == ACTION_DETECT_WINDOW_SIZE:
            if not window:
                self.log_emitted.emit("没检测到")
                return False

            left, top, right, bottom = get_window_rect(window["hwnd"])
            current_width = right - left
            current_height = bottom - top
            expected_width = int(_num(params.get("width", "1280"), "宽", integer=True, minimum=1))
            expected_height = int(_num(params.get("height", "720"), "高", integer=True, minimum=1))

            if current_width == expected_width and current_height == expected_height:
                self.log_emitted.emit("检测到")
                return True
            else:
                self.log_emitted.emit("没检测到")
                return False

        if action == ACTION_GET_WINDOW_SIZE:
            if not window:
                self.log_emitted.emit("未匹配到目标窗口")
                return

            left, top, right, bottom = get_window_rect(window["hwnd"])
            current_width = right - left
            current_height = bottom - top
            self.log_emitted.emit(f"当前窗口分辨率: {current_width} x {current_height}")
            return

        if action == ACTION_START_PROGRAM:
            args = _parse_program_command(str(params.get("program_path", "")))
            count = int(_num(params.get("count", "1"), "启动数量", integer=True, minimum=1))
            delay = float(_num(params.get("delay", "1"), "启动间隔", minimum=0))
            for _ in range(count):
                subprocess.Popen(args, shell=False)
                if self._sleep_with_stop(delay):
                    return
            return

        if action == ACTION_CLOSE_PROGRAM:
            target = self._require_window(window, action)
            close_window(target["hwnd"])
            return

        if action == ACTION_RESIZE_WINDOW:
            target = self._require_window(window, action)
            resize_window(
                target["hwnd"],
                int(_num(params.get("width", "1280"), "宽", integer=True, minimum=1)),
                int(_num(params.get("height", "720"), "高", integer=True, minimum=1)),
            )
            return

        if action == ACTION_MINIMIZE_WINDOW:
            target = self._require_window(window, action)
            minimize_window(target["hwnd"])
            return

        raise RuntimeError(f"不支持的操作: {action}")
