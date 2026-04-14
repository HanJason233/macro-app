import ctypes
import re
from typing import TypedDict


user32 = ctypes.windll.user32
dwmapi = getattr(ctypes.windll, "dwmapi", None)
SW_RESTORE = 9
SW_MINIMIZE = 6
WM_CLOSE = 0x0010
DWMWA_EXTENDED_FRAME_BOUNDS = 9


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


class POINT(ctypes.Structure):
    _fields_ = [
        ("x", ctypes.c_long),
        ("y", ctypes.c_long),
    ]


class WindowInfo(TypedDict):
    hwnd: int
    title: str
    rect: tuple[int, int, int, int]


def _to_tuple(rect: RECT) -> tuple[int, int, int, int]:
    return rect.left, rect.top, rect.right, rect.bottom


def get_window_text(hwnd: int) -> str:
    length = user32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value


def get_window_rect(hwnd: int) -> tuple[int, int, int, int]:
    rect = RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        raise RuntimeError("Failed to get window rectangle.")
    return _to_tuple(rect)


def get_window_frame_rect(hwnd: int) -> tuple[int, int, int, int]:
    if dwmapi is not None and hasattr(dwmapi, "DwmGetWindowAttribute"):
        rect = RECT()
        result = dwmapi.DwmGetWindowAttribute(
            hwnd,
            DWMWA_EXTENDED_FRAME_BOUNDS,
            ctypes.byref(rect),
            ctypes.sizeof(rect),
        )
        if result == 0:
            return _to_tuple(rect)
    return get_window_rect(hwnd)


def get_client_rect(hwnd: int) -> tuple[int, int, int, int]:
    rect = RECT()
    if not user32.GetClientRect(hwnd, ctypes.byref(rect)):
        raise RuntimeError("Failed to get client rectangle.")
    return _to_tuple(rect)


def get_client_origin(hwnd: int) -> tuple[int, int]:
    point = POINT(0, 0)
    if not user32.ClientToScreen(hwnd, ctypes.byref(point)):
        raise RuntimeError("Failed to convert client coordinates to screen coordinates.")
    return int(point.x), int(point.y)


def get_cursor_pos() -> tuple[int, int]:
    point = POINT()
    if not user32.GetCursorPos(ctypes.byref(point)):
        raise RuntimeError("Failed to get cursor position.")
    return int(point.x), int(point.y)


def enumerate_windows() -> list[WindowInfo]:
    windows: list[WindowInfo] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def enum_proc(hwnd, _l_param):
        if not user32.IsWindowVisible(hwnd):
            return True

        title = get_window_text(hwnd).strip()
        if not title:
            return True

        rect = RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            return True

        windows.append({"hwnd": int(hwnd), "title": title, "rect": _to_tuple(rect)})
        return True

    user32.EnumWindows(enum_proc, 0)
    return sorted(windows, key=lambda item: item["title"].lower())


def get_foreground_window() -> WindowInfo | None:
    hwnd = int(user32.GetForegroundWindow())
    if hwnd <= 0:
        return None
    if not user32.IsWindowVisible(hwnd):
        return None
    title = get_window_text(hwnd).strip()
    if not title:
        return None
    try:
        rect = get_window_rect(hwnd)
    except Exception:
        return None
    return {"hwnd": hwnd, "title": title, "rect": rect}


def bring_window_to_front(hwnd: int) -> None:
    user32.ShowWindow(hwnd, SW_RESTORE)
    user32.SetForegroundWindow(hwnd)


def minimize_window(hwnd: int) -> None:
    user32.ShowWindow(hwnd, SW_MINIMIZE)


def close_window(hwnd: int) -> None:
    if not user32.PostMessageW(hwnd, WM_CLOSE, 0, 0):
        raise RuntimeError("Failed to close target window.")


def resize_window(hwnd: int, width: int, height: int) -> None:
    left, top, _right, _bottom = get_window_rect(hwnd)
    user32.MoveWindow(hwnd, left, top, width, height, True)


def resolve_scope_windows(
    scope: dict, selected_hwnd: int | None, cached_windows: list[WindowInfo]
) -> list[WindowInfo]:
    scope_regex = (scope.get("regex") or "").strip()
    if not scope_regex:
        if selected_hwnd is None:
            foreground = get_foreground_window()
            return [foreground] if foreground else []
        selected = next((item for item in cached_windows if item["hwnd"] == selected_hwnd), None)
        return [selected] if selected else []

    try:
        pattern = re.compile(scope_regex)
    except re.error as exc:
        raise RuntimeError(f"Invalid scope regex: {exc}") from exc

    return [item for item in enumerate_windows() if pattern.search(item["title"])]


def resolve_scope_window(
    scope: dict, selected_hwnd: int | None, cached_windows: list[WindowInfo]
) -> WindowInfo | None:
    windows = resolve_scope_windows(scope, selected_hwnd, cached_windows)
    return windows[0] if windows else None
