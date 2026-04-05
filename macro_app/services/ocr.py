from __future__ import annotations

import re
import sys
from functools import lru_cache
from typing import Any

import numpy as np
from PIL import Image

from .capture import capture_window

SUPPORTED_TEXT_MATCH_MODES = {"exact", "contains", "regex"}


def _normalize_score(value: float) -> float:
    score = float(value)
    if score < 0 or score > 1:
        raise ValueError("最低置信度必须在 0 到 1 之间")
    return score


def _normalize_match_mode(mode: str) -> str:
    cleaned = (mode or "contains").strip().lower()
    if cleaned not in SUPPORTED_TEXT_MATCH_MODES:
        raise ValueError(f"不支持的文字匹配模式: {cleaned}")
    return cleaned


@lru_cache(maxsize=1)
def _get_ocr_engine():
    try:
        from rapidocr_onnxruntime import RapidOCR
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "未安装 OCR 依赖。"
            f" 当前解释器: {sys.executable}。"
            " 请在同一个解释器下执行: python -m pip install rapidocr_onnxruntime"
        ) from exc
    return RapidOCR()


def _preprocess_image(image: Image.Image) -> tuple[np.ndarray, int]:
    rgb = image.convert("RGB")
    frame = np.array(rgb)
    height, width = frame.shape[:2]
    if max(width, height) < 1200:
        scale = 2
        frame = np.repeat(np.repeat(frame, scale, axis=0), scale, axis=1)
        return frame, scale
    return frame, 1


def _polygon_bounds(box: Any, scale: int) -> tuple[int, int, int, int]:
    points = np.array(box, dtype=np.float32).reshape(-1, 2) / max(scale, 1)
    min_x = int(np.floor(points[:, 0].min()))
    min_y = int(np.floor(points[:, 1].min()))
    max_x = int(np.ceil(points[:, 0].max()))
    max_y = int(np.ceil(points[:, 1].max()))
    return min_x, min_y, max_x - min_x, max_y - min_y


def recognize_window_text(hwnd: int) -> list[dict[str, Any]]:
    engine = _get_ocr_engine()
    frame, scale = _preprocess_image(capture_window(hwnd))
    result, _elapsed = engine(frame)
    if not result:
        return []

    matches: list[dict[str, Any]] = []
    for item in result:
        if not isinstance(item, (list, tuple)) or len(item) < 3:
            continue
        box, text, score = item[0], str(item[1] or ""), float(item[2])
        if not text.strip():
            continue
        x, y, width, height = _polygon_bounds(box, scale)
        matches.append(
            {
                "text": text,
                "score": score,
                "x": x,
                "y": y,
                "width": width,
                "height": height,
            }
        )
    return matches


def find_text_in_window(
    hwnd: int,
    target_text: str,
    match_mode: str = "contains",
    min_score: float = 0.5,
) -> list[dict[str, Any]]:
    keyword = str(target_text or "").strip()
    if not keyword:
        raise ValueError("目标文字不能为空")

    normalized_mode = _normalize_match_mode(match_mode)
    min_confidence = _normalize_score(min_score)
    candidates = recognize_window_text(hwnd)

    found: list[dict[str, Any]] = []
    for item in candidates:
        text = str(item["text"])
        score = float(item["score"])
        if score < min_confidence:
            continue

        matched = False
        if normalized_mode == "exact":
            matched = text == keyword
        elif normalized_mode == "contains":
            matched = keyword in text
        else:
            matched = re.search(keyword, text) is not None

        if matched:
            found.append(item)

    found.sort(key=lambda item: (-float(item["score"]), int(item["y"]), int(item["x"])))
    return found
