from __future__ import annotations

import ctypes
import os
import sys

from PySide6.QtWidgets import QApplication

from .ui.main_window import MainWindow


def _enable_windows_dpi_awareness() -> None:
    """Use real monitor pixels so Windows scaling does not skew coordinates."""
    if sys.platform != "win32":
        return

    user32 = getattr(ctypes.windll, "user32", None)
    shcore = getattr(ctypes.windll, "shcore", None)

    try:
        if user32 is not None and hasattr(user32, "SetProcessDpiAwarenessContext"):
            # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
            user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
            return
    except Exception:
        pass

    try:
        if shcore is not None and hasattr(shcore, "SetProcessDpiAwareness"):
            # PROCESS_PER_MONITOR_DPI_AWARE
            shcore.SetProcessDpiAwareness(2)
            return
    except Exception:
        pass

    try:
        if user32 is not None and hasattr(user32, "SetProcessDPIAware"):
            user32.SetProcessDPIAware()
    except Exception:
        pass


def main() -> int:
    # Qt6 already configures DPI awareness on Windows.
    # Calling Windows DPI APIs here can cause Access Denied on Qt's own call.
    if os.getenv("MACRO_APP_FORCE_LEGACY_DPI_AWARENESS") == "1":
        _enable_windows_dpi_awareness()
    app = QApplication.instance() or QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()
