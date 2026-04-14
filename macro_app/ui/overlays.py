from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QPoint, QRect, Qt, Signal
from PySide6.QtGui import QColor, QCursor, QPainter, QPen
from PySide6.QtWidgets import QFileDialog, QWidget

from ..services.capture import save_cropped_screenshot
from ..services.windows import get_cursor_pos


class PointPickerOverlay(QWidget):
    point_picked = Signal(int, int)
    canceled = Signal()

    def __init__(self, rel_provider=None):
        super().__init__()
        self.rel_provider = rel_provider
        self.current_pos = QCursor.pos()
        self.native_pos = QPoint(*get_cursor_pos())
        self._overlay_text = ""
        self._crosshair_pen = QPen(QColor("#36cfc9"), 1)
        self._marker_pen = QPen(QColor("#ff4d4f"), 2)
        self.setWindowFlag(Qt.FramelessWindowHint, True)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setAttribute(Qt.WA_NoSystemBackground, True)
        self.setAutoFillBackground(False)
        self.setWindowState(Qt.WindowFullScreen)
        self.setCursor(Qt.CrossCursor)
        self.setMouseTracking(True)
        self._refresh_state()

    def _build_overlay_text(self) -> str:
        text = f"绝对坐标: ({self.native_pos.x()}, {self.native_pos.y()})"
        if self.rel_provider:
            rel = self.rel_provider(self.native_pos.x(), self.native_pos.y())
            if rel is not None:
                text += f"\n窗口相对: ({rel[0]}, {rel[1]})"
        return text + "\n左键确认  右键/Esc取消"

    def _refresh_state(self, visual_pos: QPoint | None = None) -> bool:
        native_pos = QPoint(*get_cursor_pos())
        visual = visual_pos if visual_pos is not None else QCursor.pos()
        changed = native_pos != self.native_pos or visual != self.current_pos
        self.native_pos = native_pos
        self.current_pos = visual
        if changed or not self._overlay_text:
            self._overlay_text = self._build_overlay_text()
        return changed

    def mouseMoveEvent(self, event):
        if self._refresh_state(event.globalPosition().toPoint()):
            self.update()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            point = QPoint(*get_cursor_pos())
            self.point_picked.emit(point.x(), point.y())
            self.close()
        elif event.button() == Qt.RightButton:
            self.canceled.emit()
            self.close()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.canceled.emit()
            self.close()

    def paintEvent(self, _event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor(0, 0, 0, 35))
        painter.setPen(self._crosshair_pen)

        cursor_pos = self.mapFromGlobal(self.current_pos)
        painter.drawLine(0, cursor_pos.y(), self.width(), cursor_pos.y())
        painter.drawLine(cursor_pos.x(), 0, cursor_pos.x(), self.height())

        painter.setPen(self._marker_pen)
        painter.drawEllipse(cursor_pos, 6, 6)

        painter.setPen(Qt.white)
        text_rect = QRect(cursor_pos + QPoint(18, 18), QPoint(cursor_pos.x() + 260, cursor_pos.y() + 90))
        painter.fillRect(text_rect.adjusted(-8, -6, 8, 6), QColor(0, 0, 0, 170))
        painter.drawText(text_rect, Qt.TextWordWrap, self._overlay_text)


class ScreenshotOverlay(QWidget):
    screenshot_saved = Signal(str)
    canceled = Signal()

    def __init__(self, parent=None, *, default_target: str | None = None, prompt_for_save: bool = True):
        super().__init__(parent)
        self.start_point = QPoint()
        self.end_point = QPoint()
        self.visual_start_point = QPoint()
        self.visual_end_point = QPoint()
        self.dragging = False
        self.default_target = default_target
        self.prompt_for_save = prompt_for_save
        self._selection_pen = QPen(QColor("#ff4d4f"), 2)
        self.setWindowFlag(Qt.FramelessWindowHint, True)
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        self.setWindowFlag(Qt.Tool, True)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setAttribute(Qt.WA_NoSystemBackground, True)
        self.setAutoFillBackground(False)
        self.setWindowState(Qt.WindowFullScreen)
        self.setCursor(Qt.CrossCursor)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragging = True
            self.start_point = QPoint(*get_cursor_pos())
            self.end_point = self.start_point
            self.visual_start_point = event.globalPosition().toPoint()
            self.visual_end_point = self.visual_start_point
            self.update()
        elif event.button() == Qt.RightButton:
            self.canceled.emit()
            self.close()

    def mouseMoveEvent(self, event):
        if self.dragging:
            native_end = QPoint(*get_cursor_pos())
            visual_end = event.globalPosition().toPoint()
            if native_end != self.end_point or visual_end != self.visual_end_point:
                self.end_point = native_end
                self.visual_end_point = visual_end
                self.update()

    def mouseReleaseEvent(self, event):
        if event.button() != Qt.LeftButton or not self.dragging:
            return

        self.dragging = False
        self.end_point = QPoint(*get_cursor_pos())
        self.visual_end_point = event.globalPosition().toPoint()
        rect = QRect(self.start_point, self.end_point).normalized()
        if rect.width() < 4 or rect.height() < 4:
            self.close()
            return

        if self.prompt_for_save:
            target, _ = QFileDialog.getSaveFileName(
                self,
                "保存模板图片",
                str(Path(self.default_target) if self.default_target else Path.cwd() / "template.png"),
                "PNG 图片 (*.png)",
            )
        else:
            target = self.default_target or str(Path.cwd() / "template.png")
        if target:
            path = save_cropped_screenshot(rect.left(), rect.top(), rect.right(), rect.bottom(), target)
            self.screenshot_saved.emit(path)
        else:
            self.canceled.emit()
        self.close()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.canceled.emit()
            self.close()

    def paintEvent(self, _event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor(0, 0, 0, 45))
        if self.dragging:
            rect = QRect(
                self.mapFromGlobal(self.visual_start_point),
                self.mapFromGlobal(self.visual_end_point),
            ).normalized()
            painter.fillRect(rect, QColor(255, 255, 255, 40))
            painter.setPen(self._selection_pen)
            painter.drawRect(rect)
