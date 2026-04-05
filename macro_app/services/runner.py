from __future__ import annotations

import shlex
import subprocess
import time
from pathlib import Path
from typing import Any

import pyautogui
from PySide6.QtCore import QObject, Signal

from ..constants import (
    ACTION_CLICK_ABS,
    ACTION_CLICK_IMAGE,
    ACTION_CLICK_OCR,
    ACTION_CLICK_REL,
    ACTION_CLOSE_PROGRAM,
    ACTION_DETECT_IMAGE,
    ACTION_DETECT_OCR,
    ACTION_GET_WINDOW_SIZE,
    ACTION_DETECT_WINDOW_SIZE,
    ACTION_KEY_PRESS,
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

    def __init__(self, workflow: dict, selected_hwnd: int | None, windows_snapshot: list[dict]):
        super().__init__()
        self.workflow = workflow
        self.selected_hwnd = selected_hwnd
        self.windows_snapshot = windows_snapshot
        self.should_stop = False

    def stop(self):
        self.should_stop = True

    def run(self):
        try:
            for node_index, node in enumerate(get_big_nodes(self.workflow)):
                if self.should_stop:
                    self.log_emitted.emit("工作流已停止")
                    return

                node_name = (node.get("alias") or node.get("name") or "大节点").strip()
                self.log_emitted.emit(f"进入大节点: {node_name}")

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
                    self._run_serial(node, node_index, scope, targets)
                elif mode == "sync":
                    targets = windows if windows else [None]
                    self._run_sync(node, node_index, scope, targets)
                else:
                    targets = windows if windows else [None]
                    self._run_serial(node, node_index, scope, targets)

            self.log_emitted.emit("工作流执行完成")
        except Exception as exc:
            self.log_emitted.emit(f"执行失败: {exc}")
        finally:
            self.finished.emit()

    def _run_serial(self, node: dict, node_index: int, scope: dict, targets: list[dict | None]) -> None:
        for window in targets:
            if window:
                self.log_emitted.emit(f"目标窗口: {window['title']}")
            for step_index, step in enumerate(get_small_nodes(node)):
                if self.should_stop:
                    self.log_emitted.emit("工作流已停止")
                    return
                self._emit_step_started(node_index, step_index, window)
                self._execute_step(step, scope, window)
                delay = float(_num(step.get("delay", "2"), "小节点延迟", minimum=0))
                if self._sleep_with_stop(delay):
                    self.log_emitted.emit("工作流已停止")
                    return

    def _run_sync(self, node: dict, node_index: int, scope: dict, targets: list[dict | None]) -> None:
        for step_index, step in enumerate(get_small_nodes(node)):
            if self.should_stop:
                self.log_emitted.emit("工作流已停止")
                return

            for window in targets:
                if self.should_stop:
                    self.log_emitted.emit("工作流已停止")
                    return
                self._emit_step_started(node_index, step_index, window)
                self._execute_step(step, scope, window)

            delay = float(_num(step.get("delay", "2"), "小节点延迟", minimum=0))
            if self._sleep_with_stop(delay):
                self.log_emitted.emit("工作流已停止")
                return

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

    def _execute_step(self, step: dict, scope: dict, window: dict | None):
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

        if action == ACTION_CLICK_ABS:
            pyautogui.click(
                x=int(_num(params.get("x", "0"), "绝对X", integer=True)),
                y=int(_num(params.get("y", "0"), "绝对Y", integer=True)),
                button=_mouse_button(str(params.get("button", "left"))),
                clicks=int(_num(params.get("clicks", "1"), "点击次数", integer=True, minimum=1)),
            )
            return

        if action == ACTION_CLICK_REL:
            target = self._require_window(window, action)
            left, top, _right, _bottom = get_window_frame_rect(target["hwnd"])
            pyautogui.click(
                x=left + int(_num(params.get("x", "0"), "相对X", integer=True)),
                y=top + int(_num(params.get("y", "0"), "相对Y", integer=True)),
                button=_mouse_button(str(params.get("button", "left"))),
                clicks=int(_num(params.get("clicks", "1"), "点击次数", integer=True, minimum=1)),
            )
            return

        if action == ACTION_CLICK_IMAGE:
            target = self._require_window(window, action)
            points = match_in_window(
                target["hwnd"],
                str(params.get("image_path", "")),
                str(params.get("match_mode", "default")),
                float(_num(params.get("threshold", "0.8"), "阈值", minimum=0)),
            )
            if not points:
                raise RuntimeError("未识别到目标图片")

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

        if action == ACTION_DETECT_IMAGE:
            if not window:
                self.log_emitted.emit("没检测到")
                return

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
                    return
            else:
                detected = check_image()

            if detected:
                self.log_emitted.emit("检测到")
            else:
                self.log_emitted.emit("没检测到")
            return

        if action == ACTION_DETECT_OCR:
            if not window:
                self.log_emitted.emit("没检测到")
                return

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
                    return
            else:
                detected = check_text()

            if detected:
                self.log_emitted.emit("检测到")
            else:
                self.log_emitted.emit("没检测到")
            return

        if action == ACTION_DETECT_WINDOW_SIZE:
            if not window:
                self.log_emitted.emit("没检测到")
                return

            left, top, right, bottom = get_window_rect(window["hwnd"])
            current_width = right - left
            current_height = bottom - top
            expected_width = int(_num(params.get("width", "1280"), "宽", integer=True, minimum=1))
            expected_height = int(_num(params.get("height", "720"), "高", integer=True, minimum=1))

            if current_width == expected_width and current_height == expected_height:
                self.log_emitted.emit("检测到")
            else:
                self.log_emitted.emit("没检测到")
            return

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
