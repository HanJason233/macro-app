from __future__ import annotations

from pathlib import Path
from typing import Final

ActionFieldDef = tuple[str, str, str]
ActionDefMap = dict[str, list[ActionFieldDef]]

ACTION_WAIT: Final[str] = "等待"
ACTION_KEY_PRESS: Final[str] = "键盘按键"
ACTION_TEXT_INPUT: Final[str] = "输入文本"
ACTION_CLICK: Final[str] = "点击"
ACTION_CLICK_ABS: Final[str] = "绝对坐标点击"
ACTION_CLICK_REL: Final[str] = "窗口相对点击"
ACTION_CLICK_IMAGE: Final[str] = "识图点击"
ACTION_CLICK_OCR: Final[str] = "OCR文字点击"
ACTION_MOUSE_DRAG: Final[str] = "鼠标拖动"
ACTION_MOUSE_SCROLL: Final[str] = "鼠标滑动"
ACTION_DETECT_IMAGE: Final[str] = "检测图片"
ACTION_DETECT_OCR: Final[str] = "检测文字"
ACTION_DETECT_WINDOW_SIZE: Final[str] = "检测窗口分辨率"
ACTION_GET_WINDOW_SIZE: Final[str] = "获取窗口分辨率"
ACTION_START_PROGRAM: Final[str] = "启动程序"
ACTION_CLOSE_PROGRAM: Final[str] = "关闭程序"
ACTION_RESIZE_WINDOW: Final[str] = "设定窗口大小"
ACTION_MINIMIZE_WINDOW: Final[str] = "最小化窗口"

LEGACY_ACTION_ALIASES: Final[dict[str, str]] = {
    "绛夊緟": ACTION_TEXT_INPUT,
    ACTION_WAIT: ACTION_TEXT_INPUT,
    "閿洏鎸夐敭": ACTION_KEY_PRESS,
    "杈撳叆鏂囨湰": ACTION_TEXT_INPUT,
    "缁濆鍧愭爣鐐瑰嚮": ACTION_CLICK,
    "绐楀彛鐩稿鐐瑰嚮": ACTION_CLICK,
    "璇嗗浘鐐瑰嚮": ACTION_CLICK_IMAGE,
    "OCR鏂囧瓧鐐瑰嚮": ACTION_CLICK_OCR,
    "榧犳爣鎷栧姩": ACTION_MOUSE_DRAG,
    "榧犳爣婊戝姩": ACTION_MOUSE_SCROLL,
    "妫€娴嬪浘鐗": ACTION_DETECT_IMAGE,
    "妫€娴嬫枃瀛�": ACTION_DETECT_OCR,
    "妫€娴嬬獥鍙ｅ垎杈ㄧ巼": ACTION_DETECT_WINDOW_SIZE,
    "鑾峰彇绐楀彛鍒嗚鲸鐜": ACTION_GET_WINDOW_SIZE,
    "鍚姩绋嬪簭": ACTION_START_PROGRAM,
    "鍏抽棴绋嬪簭": ACTION_CLOSE_PROGRAM,
    "窗口大小": ACTION_RESIZE_WINDOW,
    "绐楀彛澶у皬": ACTION_RESIZE_WINDOW,
    "鏈€灏忓寲绐楀彛": ACTION_MINIMIZE_WINDOW,
}

ACTION_DEFINITIONS: Final[ActionDefMap] = {
    ACTION_KEY_PRESS: [("key", "按键", "enter"), ("repeat", "重复次数", "1"), ("interval", "按键间隔", "0.1")],
    ACTION_TEXT_INPUT: [("text", "文本内容", ""), ("interval", "字间隔", "0.02")],
    ACTION_CLICK: [
        ("coord_mode", "坐标模式(relative/absolute)", "relative"),
        ("x", "X", "100"),
        ("y", "Y", "100"),
        ("button", "按键", "left"),
        ("clicks", "点击次数", "1"),
    ],
    ACTION_CLICK_IMAGE: [
        ("image_path", "模板图片", ""),
        ("match_mode", "匹配模式(default/binary/contour)", "default"),
        ("threshold", "阈值", "0.8"),
        ("click_trigger", "检测结果触发(检测到/检测不到)", "检测到"),
        ("offset_x", "偏移X", "0"),
        ("offset_y", "偏移Y", "0"),
        ("button", "按键", "left"),
        ("clicks", "点击次数", "1"),
    ],
    ACTION_CLICK_OCR: [
        ("target_text", "目标文字", ""),
        ("text_match_mode", "文字匹配(exact/contains/regex)", "contains"),
        ("min_score", "最低置信度", "0.5"),
        ("offset_x", "偏移X", "0"),
        ("offset_y", "偏移Y", "0"),
        ("button", "按键", "left"),
        ("clicks", "点击次数", "1"),
    ],
    ACTION_MOUSE_DRAG: [
        ("coord_mode", "坐标模式(relative/absolute)", "relative"),
        ("start_x", "起点X", "500"),
        ("start_y", "起点Y", "500"),
        ("end_x", "终点X", "700"),
        ("end_y", "终点Y", "700"),
        ("hold_duration", "按住时长(s)", "0.2"),
        ("move_duration", "拖动时长(s)", "0.4"),
    ],
    ACTION_MOUSE_SCROLL: [
        ("coord_mode", "坐标模式(relative/absolute)", "relative"),
        ("start_x", "起点X", "500"),
        ("start_y", "起点Y", "500"),
        ("end_x", "终点X", "700"),
        ("end_y", "终点Y", "700"),
        ("move_duration", "滑动时长(s)", "0.4"),
    ],
    ACTION_DETECT_IMAGE: [
        ("image_path", "模板图片", ""),
        ("match_mode", "匹配模式(default/binary/contour)", "default"),
        ("threshold", "阈值", "0.8"),
        ("detect_duration", "检测时长(s)", "0"),
        ("detect_interval", "检测频率(s)", "0.5"),
    ],
    ACTION_DETECT_OCR: [
        ("target_text", "目标文字", ""),
        ("text_match_mode", "文字匹配(exact/contains/regex)", "contains"),
        ("min_score", "最低置信度", "0.5"),
        ("detect_duration", "检测时长(s)", "0"),
        ("detect_interval", "检测频率(s)", "0.5"),
    ],
    ACTION_DETECT_WINDOW_SIZE: [("width", "宽", "1280"), ("height", "高", "720")],
    ACTION_GET_WINDOW_SIZE: [],
    ACTION_START_PROGRAM: [("program_path", "程序路径", ""), ("count", "启动数量", "1"), ("delay", "启动间隔", "1")],
    ACTION_CLOSE_PROGRAM: [],
    ACTION_RESIZE_WINDOW: [("width", "宽", "1280"), ("height", "高", "720")],
    ACTION_MINIMIZE_WINDOW: [],
}

ACTION_GROUPS: Final[dict[str, list[str]]] = {
    "键盘": [ACTION_KEY_PRESS, ACTION_TEXT_INPUT],
    "鼠标操作": [
        ACTION_CLICK,
        ACTION_CLICK_IMAGE,
        ACTION_CLICK_OCR,
        ACTION_MOUSE_DRAG,
        ACTION_MOUSE_SCROLL,
    ],
    "窗口与程序": [ACTION_START_PROGRAM, ACTION_CLOSE_PROGRAM, ACTION_RESIZE_WINDOW, ACTION_MINIMIZE_WINDOW],
}

WINDOW_SCOPE_MODES: Final[list[str]] = ["当前选中窗口", "标题匹配正则"]


def normalize_action_name(action_name: str) -> str:
    if action_name == ACTION_WAIT:
        return ACTION_TEXT_INPUT
    if action_name in {ACTION_CLICK_ABS, ACTION_CLICK_REL}:
        return ACTION_CLICK
    if action_name in ACTION_DEFINITIONS:
        return action_name
    return LEGACY_ACTION_ALIASES.get(action_name, ACTION_TEXT_INPUT)


def build_default_params(action_name: str) -> dict[str, str]:
    action_name = normalize_action_name(action_name)
    fields = ACTION_DEFINITIONS[action_name]
    return {field: default for field, _label, default in fields}


def summarize_step(step: dict) -> str:
    params = step.get("params", {})
    action = normalize_action_name(step.get("action", ACTION_TEXT_INPUT))
    suffix = f" | 延迟 {step.get('delay', '2')}s"
    coord_suffix = f" [{params.get('coord_mode', 'relative')}]"
    if action == ACTION_KEY_PRESS:
        return f"按键 {params.get('key', 'enter')} x{params.get('repeat', '1')}{suffix}"
    if action == ACTION_TEXT_INPUT:
        return f"输入文本 {params.get('text', '')[:12] or '(空)'}{suffix}"
    if action == ACTION_CLICK:
        return f"坐标点击 ({params.get('x')}, {params.get('y')}){coord_suffix}{suffix}"
    if action == ACTION_CLICK_IMAGE:
        return f"识图点击 {Path(params.get('image_path') or '未设置').name}{suffix}"
    if action == ACTION_CLICK_OCR:
        return f"OCR点击 {params.get('target_text', '') or '未设置'}{suffix}"
    if action == ACTION_MOUSE_DRAG:
        return f"鼠标拖动 ({params.get('start_x')}, {params.get('start_y')}) -> ({params.get('end_x')}, {params.get('end_y')}){coord_suffix}{suffix}"
    if action == ACTION_MOUSE_SCROLL:
        return f"鼠标滑动 ({params.get('start_x')}, {params.get('start_y')}) -> ({params.get('end_x')}, {params.get('end_y')}){coord_suffix}{suffix}"
    if action == ACTION_DETECT_IMAGE:
        return f"检测图片 {Path(params.get('image_path') or '未设置').name}{suffix}"
    if action == ACTION_DETECT_OCR:
        return f"检测文字 {params.get('target_text', '') or '未设置'}{suffix}"
    if action == ACTION_DETECT_WINDOW_SIZE:
        return f"检测分辨率 {params.get('width')} x {params.get('height')}{suffix}"
    if action == ACTION_GET_WINDOW_SIZE:
        return f"获取窗口分辨率{suffix}"
    if action == ACTION_START_PROGRAM:
        return f"启动程序 {Path(params.get('program_path') or '未设置').name}{suffix}"
    if action == ACTION_CLOSE_PROGRAM:
        return f"关闭程序{suffix}"
    if action == ACTION_RESIZE_WINDOW:
        return f"设定窗口大小 {params.get('width')} x {params.get('height')}{suffix}"
    if action == ACTION_MINIMIZE_WINDOW:
        return f"最小化窗口{suffix}"
    return f"{action}{suffix}"
