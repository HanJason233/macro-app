from __future__ import annotations

from pathlib import Path

import cv2
import numpy as np
from PIL import Image, ImageGrab

from .windows import get_window_rect

SUPPORTED_MATCH_MODES = {"default", "binary", "contour"}


def capture_window(hwnd: int) -> Image.Image:
    left, top, right, bottom = get_window_rect(hwnd)
    return ImageGrab.grab(bbox=(left, top, right, bottom))


def save_cropped_screenshot(left: int, top: int, right: int, bottom: int, target: str) -> str:
    if right <= left or bottom <= top:
        raise ValueError("截图区域无效：宽高必须大于 0")

    image = ImageGrab.grab(bbox=(left, top, right, bottom))
    target_path = Path(target)
    target_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(target_path)
    return str(target_path)


def _validate_template(template_path: str) -> Path:
    path = Path(template_path).expanduser()
    if not path.exists() or not path.is_file():
        raise RuntimeError(f"模板图片不存在: {path}")
    return path


def _normalize_threshold(threshold: float) -> float:
    if threshold < 0 or threshold > 1:
        raise ValueError("阈值必须在 0 到 1 之间")
    return threshold


def match_in_window(hwnd: int, template_path: str, mode: str, threshold: float) -> list[tuple[int, int, int, int]]:
    mode = (mode or "default").strip().lower()
    if mode not in SUPPORTED_MATCH_MODES:
        raise ValueError(f"不支持的匹配模式: {mode}")

    template_file = _validate_template(template_path)
    threshold = _normalize_threshold(float(threshold))

    frame = cv2.cvtColor(np.array(capture_window(hwnd)), cv2.COLOR_RGB2BGR)
    template = cv2.imread(str(template_file))
    if template is None:
        raise RuntimeError(f"模板图片读取失败: {template_file}")

    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

    if mode == "binary":
        _, gray = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
        _, template_gray = cv2.threshold(template_gray, 127, 255, cv2.THRESH_BINARY)

    if mode in {"default", "binary"}:
        result = cv2.matchTemplate(gray, template_gray, cv2.TM_CCOEFF_NORMED)
        ys, xs = np.where(result >= threshold)
        height, width = template_gray.shape
        return [(int(x), int(y), int(width), int(height)) for y, x in zip(ys, xs)]

    contours_frame, _ = cv2.findContours(cv2.threshold(gray, 127, 255, 0)[1], cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours_template, _ = cv2.findContours(cv2.threshold(template_gray, 127, 255, 0)[1], cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours_template:
        return []

    target = max(contours_template, key=cv2.contourArea)
    matches: list[tuple[int, int, int, int]] = []
    for contour in contours_frame:
        if cv2.matchShapes(contour, target, cv2.CONTOURS_MATCH_I1, 0.0) < 0.1:
            matches.append(tuple(map(int, cv2.boundingRect(contour))))
    return matches
