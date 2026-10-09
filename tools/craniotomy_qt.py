import math
import ctypes
import json
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QEvent, QPoint, QPointF, QRectF, QSize, Qt, QProcess, QTimer, Signal
from PySide6.QtGui import QColor, QFont, QImage, QKeySequence, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QFileDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QKeySequenceEdit,
    QSizePolicy,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from stereodrive_controller import StereoDriveController, StereoDriveError


user32 = ctypes.WinDLL("user32", use_last_error=True)
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)

SRCCOPY = 0x00CC0020
BI_RGB = 0
DIB_RGB_COLORS = 0
PW_RENDERFULLCONTENT = 0x00000002


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", ctypes.c_uint32),
        ("biWidth", ctypes.c_int32),
        ("biHeight", ctypes.c_int32),
        ("biPlanes", ctypes.c_uint16),
        ("biBitCount", ctypes.c_uint16),
        ("biCompression", ctypes.c_uint32),
        ("biSizeImage", ctypes.c_uint32),
        ("biXPelsPerMeter", ctypes.c_int32),
        ("biYPelsPerMeter", ctypes.c_int32),
        ("biClrUsed", ctypes.c_uint32),
        ("biClrImportant", ctypes.c_uint32),
    ]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [
        ("bmiHeader", BITMAPINFOHEADER),
        ("bmiColors", ctypes.c_uint32 * 3),
    ]


class NumericLineEdit(QLineEdit):
    valueChanged = Signal(int)

    def __init__(
        self,
        value: float | int = 0,
        minimum: float | int = -100.0,
        maximum: float | int = 100.0,
        integer: bool = False,
    ) -> None:
        super().__init__()
        self._minimum = minimum
        self._maximum = maximum
        self._integer = integer
        self.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.setValue(value)
        self.editingFinished.connect(self._commit_text)

    def setRange(self, minimum: float | int, maximum: float | int) -> None:  # noqa: N802
        self._minimum = minimum
        self._maximum = maximum
        self.setValue(self.value())

    def value(self) -> float | int:
        text = self.text().strip()
        try:
            value = float(text)
        except ValueError:
            value = float(self._minimum)
        value = max(float(self._minimum), min(float(self._maximum), value))
        if self._integer:
            return int(round(value))
        return value

    def setValue(self, value: float | int) -> None:  # noqa: N802
        value = max(float(self._minimum), min(float(self._maximum), float(value)))
        if self._integer:
            text = str(int(round(value)))
        else:
            text = f"{value:g}"
        changed = self.text() != text
        self.setText(text)
        if changed and not self.signalsBlocked():
            self.valueChanged.emit(int(round(value)))

    def _commit_text(self) -> None:
        self.setValue(self.value())


user32.GetDC.argtypes = [ctypes.c_void_p]
user32.GetDC.restype = ctypes.c_void_p
user32.ReleaseDC.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
user32.ReleaseDC.restype = ctypes.c_int
user32.PrintWindow.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint]
user32.PrintWindow.restype = ctypes.c_bool
gdi32.CreateCompatibleDC.argtypes = [ctypes.c_void_p]
gdi32.CreateCompatibleDC.restype = ctypes.c_void_p
gdi32.DeleteDC.argtypes = [ctypes.c_void_p]
gdi32.DeleteDC.restype = ctypes.c_bool
gdi32.CreateCompatibleBitmap.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
gdi32.CreateCompatibleBitmap.restype = ctypes.c_void_p
gdi32.DeleteObject.argtypes = [ctypes.c_void_p]
gdi32.DeleteObject.restype = ctypes.c_bool
gdi32.SelectObject.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
gdi32.SelectObject.restype = ctypes.c_void_p
gdi32.GetDIBits.argtypes = [
    ctypes.c_void_p,
    ctypes.c_void_p,
    ctypes.c_uint,
    ctypes.c_uint,
    ctypes.c_void_p,
    ctypes.POINTER(BITMAPINFO),
    ctypes.c_uint,
]
gdi32.GetDIBits.restype = ctypes.c_int


MOVE_SPEED_OPTIONS_MM = [0.001, 0.005, 0.01, 0.02, 0.05, 0.1, 0.2, 0.5, 1.0, 2.0, 5.0]
DEFAULT_MOVE_SPEED_MM = 0.05
INJECTION_VOLUME_OPTIONS_NL = [10, 20, 50, 100, 200, 500, 1000, 2000]
DEFAULT_INJECTION_VOLUME_NL = 100
SYRINGE_MIN_NL = 0.0
SYRINGE_MAX_NL = 5000.0


@dataclass
class SeedPoint:
    index: int
    angle_deg: float
    ap: float
    ml: float
    dv: float | None = None
    sampled_ap: float | None = None
    sampled_ml: float | None = None


@dataclass
class InjectionSite:
    ap: float
    ml: float
    dv: float | None = None
    generated: bool = False


@dataclass
class StoredLocation:
    ap: float
    ml: float
    dv: float
    coordinate_system: str = "bregma"


@dataclass
class InjectionProtocolSettings:
    main_volume_nl: int
    insertion_rate_nl_min: float
    main_rate_nl_min: float
    injection_depth_mm: float
    insert_retract_speed_um_s: float
    overshoot_mm: float
    post_inject_pause_s: float


@dataclass
class CraniotomyConfig:
    diameter_mm: float
    seed_count: int
    trajectory_points: int
    cut_offset_dv_mm: float
    max_depth_mm: float
    depth_per_round_mm: float
    skull_thickness_mm: float
    round_time_seconds: float
    drill_rate_mm_per_s: float
    auto_start_rounds: bool


class ProjectionWidget(QWidget):
    freeze_drawn = Signal(int)
    unfreeze_drawn = Signal(int)
    location_double_clicked = Signal(float, float)

    def __init__(self, x_label: str, y_label: str, invert_y: bool = False, parent: QWidget | None = None):
        super().__init__(parent)
        self.x_label = x_label
        self.y_label = y_label
        self.invert_y = invert_y
        self.trajectory: list[tuple[float, float, float]] = []
        self.seed_points: list[tuple[float, float, bool]] = []
        self.injection_site_points: list[tuple[float, float]] = []
        self.anchor_point: tuple[float, float] | None = None
        self.frozen_points: list[bool] = []
        self.current_point: tuple[float, float] | None = None
        self.freeze_mode = False
        self.unfreeze_mode = False
        self._trajectory_screen_points: list[QPointF] = []
        self._inner_ring_screen_points: list[QPointF] = []
        self._coordinate_bounds: tuple[float, float, float, float] | None = None
        self.overlay_image: QImage | None = None
        self.overlay_calibration: dict[str, object] | None = None
        self.coordinate_mode_bregma = False
        self.zoom_level = 1.0
        self.view_focus_points: list[tuple[float, float]] | None = None
        self.navigation_enabled = False
        self.navigation_zoom = 1.0
        self._minimum_navigation_zoom = 1.0
        self.navigation_pan = QPointF(0.0, 0.0)
        self._pan_anchor: QPointF | None = None
        self._pan_start = QPointF(0.0, 0.0)
        self._draw_rect: QRectF | None = None
        self.setMinimumHeight(360)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

    def set_overlay_image(self, image: QImage | None, calibration: dict[str, object] | None = None) -> None:
        self.overlay_image = image
        self.overlay_calibration = calibration
        self.update()

    def set_coordinate_mode_bregma(self, enabled: bool) -> None:
        self.coordinate_mode_bregma = enabled
        self.update()

    def set_zoom_level(self, level: float) -> None:
        self.zoom_level = max(0.0, min(1.0, level))
        self.update()

    def set_view_focus_points(self, points: list[tuple[float, float]] | None) -> None:
        self.view_focus_points = points
        self.update()

    def set_navigation_enabled(self, enabled: bool) -> None:
        self.navigation_enabled = enabled
        self.setCursor(Qt.OpenHandCursor if enabled else Qt.ArrowCursor)

    def reset_navigation(self) -> None:
        self.navigation_zoom = 1.0
        self.navigation_pan = QPointF(0.0, 0.0)
        self.update()

    def hasHeightForWidth(self) -> bool:  # noqa: N802
        return True

    def heightForWidth(self, width: int) -> int:  # noqa: N802
        return width

    def set_data(
        self,
        trajectory: list[tuple[float, float, float]],
        seed_points: list[tuple[float, float, bool]],
        frozen_points: list[bool] | None = None,
        current_point: tuple[float, float] | None = None,
        injection_sites: list[tuple[float, float]] | None = None,
        anchor_point: tuple[float, float] | None = None,
    ) -> None:
        self.trajectory = trajectory
        self.seed_points = seed_points
        self.injection_site_points = injection_sites or []
        self.anchor_point = anchor_point
        self.frozen_points = frozen_points or [False] * len(trajectory)
        self.current_point = current_point
        self.update()

    def set_freeze_mode(self, enabled: bool) -> None:
        self.freeze_mode = enabled
        if enabled:
            self.unfreeze_mode = False
        self.update()

    def set_unfreeze_mode(self, enabled: bool) -> None:
        self.unfreeze_mode = enabled
        if enabled:
            self.freeze_mode = False
        self.update()

    def mousePressEvent(self, event) -> None:  # noqa: N802
        if self.freeze_mode and event.button() == Qt.LeftButton:
            self._emit_nearest_trajectory_index(event.position(), freeze=True)
            event.accept()
            return
        if self.unfreeze_mode and event.button() == Qt.LeftButton:
            self._emit_nearest_trajectory_index(event.position(), freeze=False)
            event.accept()
            return
        if self.navigation_enabled and event.button() == Qt.LeftButton:
            self._pan_anchor = event.position()
            self._pan_start = QPointF(self.navigation_pan)
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if self._pan_anchor is not None and event.button() == Qt.LeftButton:
            self._pan_anchor = None
            self.setCursor(Qt.OpenHandCursor)
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def wheelEvent(self, event) -> None:  # noqa: N802
        if self.navigation_enabled and event.angleDelta().y():
            steps = event.angleDelta().y() / 120.0
            minimum_zoom = self._minimum_navigation_zoom
            self.navigation_zoom = max(minimum_zoom, min(20.0, self.navigation_zoom * (1.25 ** steps)))
            # At the fully zoomed-out limit show the complete skull, rather
            # than a panned crop of it.
            if self.navigation_zoom <= minimum_zoom + 1e-6:
                self.navigation_pan = QPointF(0.0, 0.0)
            self.update()
            event.accept()
            return
        super().wheelEvent(event)

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        if self._pan_anchor is not None and self._coordinate_bounds is not None and self._draw_rect is not None:
            min_x, max_x, min_y, max_y = self._coordinate_bounds
            dx = event.position().x() - self._pan_anchor.x()
            dy = event.position().y() - self._pan_anchor.y()
            span_x = max_x - min_x
            span_y = max_y - min_y
            self.navigation_pan = QPointF(
                self._pan_start.x() - dx / max(1.0, self._draw_rect.width()) * span_x,
                self._pan_start.y() + dy / max(1.0, self._draw_rect.height()) * span_y * (-1.0 if self.invert_y else 1.0),
            )
            self.update()
            event.accept()
            return
        if self.freeze_mode and event.buttons() & Qt.LeftButton:
            self._emit_nearest_trajectory_index(event.position(), freeze=True)
            event.accept()
            return
        if self.unfreeze_mode and event.buttons() & Qt.LeftButton:
            self._emit_nearest_trajectory_index(event.position(), freeze=False)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseDoubleClickEvent(self, event) -> None:  # noqa: N802
        if self._coordinate_bounds is not None and event.button() == Qt.LeftButton:
            min_x, max_x, min_y, max_y = self._coordinate_bounds
            x = min_x + (event.position().x() - 24) / max(1.0, self.width() - 48) * (max_x - min_x)
            normalized_y = (event.position().y() - 24) / max(1.0, self.height() - 48)
            y = max_y - normalized_y * (max_y - min_y)
            self.location_double_clicked.emit(float(x), float(y))
            event.accept()
            return
        super().mouseDoubleClickEvent(event)

    def _emit_nearest_trajectory_index(self, position, freeze: bool) -> None:
        if not self._trajectory_screen_points:
            return
        nearest_index = None
        nearest_distance = 18.0
        for index, point in enumerate(self._trajectory_screen_points):
            dx = point.x() - position.x()
            dy = point.y() - position.y()
            distance = math.hypot(dx, dy)
            if distance <= nearest_distance:
                nearest_distance = distance
                nearest_index = index
        if nearest_index is not None:
            if freeze:
                self.freeze_drawn.emit(nearest_index)
            else:
                self.unfreeze_drawn.emit(nearest_index)

    def paintEvent(self, _event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QColor("#fbfcfa"))

        pad = 24
        available_rect = self.rect().adjusted(pad, pad, -pad, -pad)
        side = min(available_rect.width(), available_rect.height())
        draw_rect = QRectF(
            available_rect.left() + (available_rect.width() - side) / 2.0,
            available_rect.top() + (available_rect.height() - side) / 2.0,
            side,
            side,
        )
        self._draw_rect = draw_rect

        painter.setPen(QPen(QColor("#cad7cb"), 1))
        painter.setBrush(QColor("#ffffff"))
        painter.drawRoundedRect(draw_rect, 16, 16)
        overlay_visible = self.overlay_image is not None and self.coordinate_mode_bregma
        overlay_hidden_for_axis = self.overlay_image is not None and not self.coordinate_mode_bregma
        if (
            not self.trajectory
            and not self.seed_points
            and not self.injection_site_points
            and self.anchor_point is None
            and self.current_point is None
            and not overlay_visible
        ):
            painter.setPen(QColor("#8b9a8d"))
            message = "No trajectory yet"
            if overlay_hidden_for_axis:
                message += "\nSwitch to Bregma reference mode to display overlay."
            painter.drawText(self.rect(), Qt.AlignCenter, message)
            return

        xs = [p[0] for p in self.trajectory] + [s[0] for s in self.seed_points] + [p[0] for p in self.injection_site_points]
        ys = [p[1] for p in self.trajectory] + [s[1] for s in self.seed_points] + [p[1] for p in self.injection_site_points]
        if self.anchor_point is not None:
            xs.append(self.anchor_point[0])
            ys.append(self.anchor_point[1])
        if self.view_focus_points:
            xs = [point[0] for point in self.view_focus_points]
            ys = [point[1] for point in self.view_focus_points]
        overlay_bounds = None
        if overlay_visible and self.overlay_calibration:
            bregma = self.overlay_calibration.get("bregma_pixel")
            lambda_pixel = self.overlay_calibration.get("lambda_pixel")
            distance_mm = float(self.overlay_calibration.get("bregma_to_lambda_mm", 3.9))
            if isinstance(bregma, list) and isinstance(lambda_pixel, list) and len(bregma) == 2 and len(lambda_pixel) == 2:
                pixel_distance = math.hypot(lambda_pixel[0] - bregma[0], lambda_pixel[1] - bregma[1])
                if pixel_distance > 0:
                    mm_per_pixel = distance_mm / pixel_distance
                    overlay_bounds = (
                        -bregma[0] * mm_per_pixel,
                        (self.overlay_image.width() - bregma[0]) * mm_per_pixel,
                        -(self.overlay_image.height() - bregma[1]) * mm_per_pixel,
                        bregma[1] * mm_per_pixel,
                    )
        if overlay_bounds is not None and (not xs or self.zoom_level < 1.0):
            if not xs:
                xs = [overlay_bounds[0], overlay_bounds[1]]
                ys = [overlay_bounds[2], overlay_bounds[3]]
            elif self.zoom_level < 1.0:
                trajectory_bounds = (min(xs), max(xs), min(ys), max(ys))
                xs = [
                    overlay_bounds[0] + self.zoom_level * (trajectory_bounds[0] - overlay_bounds[0]),
                    overlay_bounds[1] + self.zoom_level * (trajectory_bounds[1] - overlay_bounds[1]),
                ]
                ys = [
                    overlay_bounds[2] + self.zoom_level * (trajectory_bounds[2] - overlay_bounds[2]),
                    overlay_bounds[3] + self.zoom_level * (trajectory_bounds[3] - overlay_bounds[3]),
                ]
        if self.current_point is not None:
            xs.append(self.current_point[0])
            ys.append(self.current_point[1])
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)
        if math.isclose(min_x, max_x):
            min_x -= 1.0
            max_x += 1.0
        if math.isclose(min_y, max_y):
            min_y -= 1.0
            max_y += 1.0

        span_x = max_x - min_x
        span_y = max_y - min_y
        base_uniform_span = max(span_x, span_y) * 1.3
        base_cx = (min_x + max_x) / 2.0
        base_cy = (min_y + max_y) / 2.0
        self._minimum_navigation_zoom = 1.0
        if overlay_bounds is not None:
            overlay_min_x, overlay_max_x, overlay_min_y, overlay_max_y = overlay_bounds
            full_skull_span = 2.0 * max(
                abs(overlay_min_x - base_cx),
                abs(overlay_max_x - base_cx),
                abs(overlay_min_y - base_cy),
                abs(overlay_max_y - base_cy),
            )
            if full_skull_span > 0:
                self._minimum_navigation_zoom = min(1.0, base_uniform_span / full_skull_span)
        uniform_span = base_uniform_span / self.navigation_zoom
        cx = base_cx + self.navigation_pan.x()
        cy = base_cy + self.navigation_pan.y()
        min_x = cx - uniform_span / 2.0
        max_x = cx + uniform_span / 2.0
        min_y = cy - uniform_span / 2.0
        max_y = cy + uniform_span / 2.0
        self._coordinate_bounds = (min_x, max_x, min_y, max_y)

        def map_point(x: float, y: float) -> QPointF:
            px = draw_rect.left() + (x - min_x) / (max_x - min_x) * draw_rect.width()
            normalized_y = (y - min_y) / (max_y - min_y)
            if self.invert_y:
                py = draw_rect.top() + normalized_y * draw_rect.height()
            else:
                py = draw_rect.bottom() - normalized_y * draw_rect.height()
            return QPointF(px, py)

        if overlay_visible and self.overlay_calibration:
            calibration = self.overlay_calibration
            bregma = calibration.get("bregma_pixel")
            lambda_pixel = calibration.get("lambda_pixel")
            distance_mm = float(calibration.get("bregma_to_lambda_mm", 3.9))
            if isinstance(bregma, list) and isinstance(lambda_pixel, list) and len(bregma) == 2 and len(lambda_pixel) == 2:
                pixel_distance = math.hypot(lambda_pixel[0] - bregma[0], lambda_pixel[1] - bregma[1])
                if pixel_distance > 0 and distance_mm > 0:
                    mm_per_pixel = distance_mm / pixel_distance
                    left_mm = -bregma[0] * mm_per_pixel
                    right_mm = (self.overlay_image.width() - bregma[0]) * mm_per_pixel
                    top_mm = bregma[1] * mm_per_pixel
                    bottom_mm = -(self.overlay_image.height() - bregma[1]) * mm_per_pixel
                    overlay_top_left = map_point(left_mm, top_mm)
                    overlay_bottom_right = map_point(right_mm, bottom_mm)
                    painter.save()
                    painter.setOpacity(0.42)
                    painter.drawImage(QRectF(overlay_top_left, overlay_bottom_right), self.overlay_image)
                    painter.restore()

        self._trajectory_screen_points = [map_point(point[0], point[1]) for point in self.trajectory]
        if self.trajectory:
            center_x = sum(point[0] for point in self.trajectory) / len(self.trajectory)
            center_y = sum(point[1] for point in self.trajectory) / len(self.trajectory)
            self._inner_ring_screen_points = []
            for point in self.trajectory:
                dx = point[0] - center_x
                dy = point[1] - center_y
                self._inner_ring_screen_points.append(map_point(center_x + dx * 0.88, center_y + dy * 0.88))
        else:
            self._inner_ring_screen_points = []

        if len(self.trajectory) > 1:
            for index, (start, end) in enumerate(zip(self.trajectory[:-1], self.trajectory[1:])):
                progress = max(0.0, min(1.0, (start[2] + end[2]) / 2.0))
                red = int(235 + progress * 20)
                green = int(215 - progress * 160)
                blue = int(70 - progress * 50)
                painter.setPen(QPen(QColor(red, green, max(0, blue)), 6))
                painter.drawLine(map_point(start[0], start[1]), map_point(end[0], end[1]))

        if len(self._inner_ring_screen_points) > 1:
            for index, (start, end) in enumerate(zip(self._inner_ring_screen_points[:-1], self._inner_ring_screen_points[1:])):
                frozen = index < len(self.frozen_points) and self.frozen_points[index]
                painter.setPen(QPen(QColor("#2563eb" if frozen else "#16a34a"), 4))
                painter.drawLine(start, end)

        for idx, (x, y, sampled) in enumerate(self.seed_points, start=1):
            pt = map_point(x, y)
            color = QColor("#0d8a63" if sampled else "#dd6e42")
            painter.setPen(Qt.NoPen)
            painter.setBrush(color)
            painter.drawEllipse(pt, 6, 6)
            painter.setPen(color)
            painter.drawText(pt + QPointF(8, -8), f"{idx} [{x:.2f}, {y:.2f}]")

        for idx, (x, y) in enumerate(self.injection_site_points, start=1):
            pt = map_point(x, y)
            painter.setPen(QPen(QColor("#6b21a8"), 2))
            painter.setBrush(QColor("#d8b4fe"))
            painter.drawEllipse(pt, 8, 8)
            painter.setPen(QColor("#3b0764"))
            painter.drawText(QRectF(pt.x() - 6, pt.y() - 8, 12, 16), Qt.AlignCenter, str(idx))

        if self.anchor_point is not None:
            pt = map_point(self.anchor_point[0], self.anchor_point[1])
            anchor_color = QColor("#2563eb")
            painter.setPen(QPen(anchor_color, 3, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
            painter.setBrush(Qt.NoBrush)
            # A font-independent anchor icon: ring, shank, stock, and flukes.
            painter.drawEllipse(pt + QPointF(0, -7), 2.5, 2.5)
            painter.drawLine(pt + QPointF(0, -4), pt + QPointF(0, 9))
            painter.drawLine(pt + QPointF(-8, 1), pt + QPointF(8, 1))
            painter.drawLine(pt + QPointF(0, 9), pt + QPointF(-8, 4))
            painter.drawLine(pt + QPointF(0, 9), pt + QPointF(8, 4))
            painter.setPen(QColor("#1d4ed8"))
            painter.drawText(pt + QPointF(11, -10), "Anchor")

        if self.current_point is not None:
            pt = map_point(self.current_point[0], self.current_point[1])
            marker_pen = QPen(QColor("#dc2626" if self.coordinate_mode_bregma else "#1f2937"), 4)
            painter.setPen(marker_pen)
            painter.drawLine(pt + QPointF(-10, -10), pt + QPointF(10, 10))
            painter.drawLine(pt + QPointF(-10, 10), pt + QPointF(10, -10))

        if overlay_hidden_for_axis:
            painter.setPen(QColor("#6b7280"))
            painter.drawText(
                self.rect().adjusted(10, 8, -10, -8),
                Qt.AlignHCenter | Qt.AlignTop,
                "Switch to Bregma reference mode to display overlay.",
            )

        if self.freeze_mode:
            painter.setPen(QColor("#b23a48"))
            painter.drawText(self.rect().adjusted(0, 0, -10, -10), Qt.AlignRight | Qt.AlignTop, "Draw Freeze")
        elif self.unfreeze_mode:
            painter.setPen(QColor("#1d4ed8"))
            painter.drawText(self.rect().adjusted(0, 0, -10, -10), Qt.AlignRight | Qt.AlignTop, "Draw Unfreeze")


class DepthLegendWidget(QWidget):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.skull_thickness_mm = 0.25
        self.current_depth_ratio: float | None = None
        self.setMinimumWidth(90)
        self.setMaximumWidth(90)
        self.setMinimumHeight(260)

    def set_skull_thickness_mm(self, skull_thickness_mm: float) -> None:
        self.skull_thickness_mm = skull_thickness_mm
        self.update()

    def set_current_depth_ratio(self, current_depth_ratio: float | None) -> None:
        self.current_depth_ratio = current_depth_ratio
        self.update()

    def paintEvent(self, _event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QColor("#fbfcfa"))
        bar_rect = QRectF(24, 24, 18, max(120, self.height() - 56))
        for idx in range(int(bar_rect.height())):
            ratio = idx / max(1.0, bar_rect.height() - 1.0)
            red = int(235 + ratio * 20)
            green = int(215 - ratio * 160)
            blue = int(70 - ratio * 50)
            painter.setPen(QPen(QColor(red, green, max(0, blue)), 1))
            y = bar_rect.top() + idx
            painter.drawLine(QPointF(bar_rect.left(), y), QPointF(bar_rect.right(), y))
        painter.setPen(QColor("#496052"))
        painter.drawRect(bar_rect)
        painter.drawText(QRectF(0, 4, self.width(), 18), Qt.AlignHCenter, "Depth")
        painter.drawText(QRectF(48, bar_rect.top() - 6, self.width() - 50, 20), Qt.AlignLeft | Qt.AlignVCenter, "0 mm")
        painter.drawText(
            QRectF(48, bar_rect.bottom() - 10, self.width() - 50, 28),
            Qt.AlignLeft | Qt.AlignVCenter,
            f"{self.skull_thickness_mm:.2f} mm",
        )
        if self.current_depth_ratio is not None:
            ratio = max(0.0, min(1.0, self.current_depth_ratio))
            y = bar_rect.top() + ratio * bar_rect.height()
            painter.setPen(QPen(QColor("#111111"), 2))
            painter.drawLine(QPointF(bar_rect.left() - 6, y), QPointF(bar_rect.right() + 6, y))


class PlungerGaugeWidget(QWidget):
    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.position_nl: float | None = None
        self.maximum_nl = 5000.0
        self.setMinimumWidth(92)
        self.setMaximumWidth(92)
        self.setMinimumHeight(520)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)

    def set_position(self, position_nl: float | None) -> None:
        self.position_nl = position_nl
        self.update()

    def paintEvent(self, _event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QColor("#ffffff"))

        top = 28.0
        bottom = max(top + 120.0, self.height() - 48.0)
        axis_x = 36.0
        painter.setPen(QPen(QColor("#1f2937"), 1))
        painter.drawLine(QPointF(axis_x, top), QPointF(axis_x, bottom))

        for value in range(0, int(self.maximum_nl) + 1, 100):
            ratio = value / self.maximum_nl
            y = bottom - ratio * (bottom - top)
            is_major = value % 500 == 0
            tick_length = 10 if is_major else 5
            painter.setPen(QPen(QColor("#1f2937"), 1))
            painter.drawLine(QPointF(axis_x, y), QPointF(axis_x + tick_length, y))
            if is_major:
                painter.drawText(QRectF(axis_x + 14, y - 8, 40, 16), Qt.AlignLeft | Qt.AlignVCenter, str(value))

        painter.drawText(QRectF(axis_x + 12, 4, 42, 18), Qt.AlignRight | Qt.AlignVCenter, "nl")

        if self.position_nl is not None:
            value = max(0.0, min(self.maximum_nl, self.position_nl))
            ratio = value / self.maximum_nl
            y = bottom - ratio * (bottom - top)
            painter.setPen(QPen(QColor("#1d4ed8"), 6))
            painter.drawLine(QPointF(axis_x - 6, top), QPointF(axis_x - 6, y))
            painter.setPen(QColor("#1d4ed8"))
            painter.drawText(QRectF(4, self.height() - 28, self.width() - 8, 24), Qt.AlignCenter, f"{self.position_nl:.0f}")
        else:
            painter.setPen(QColor("#6b7280"))
            painter.drawText(QRectF(4, self.height() - 28, self.width() - 8, 24), Qt.AlignCenter, "--")


class CraniotomyWindow(QMainWindow):
    status_signal = Signal(str)
    redraw_signal = Signal()
    drill_progress_signal = Signal(int)
    injection_progress_signal = Signal(int, str)
    injection_site_progress_signal = Signal(int)
    sequence_step_signal = Signal(int)
    active_injection_site_signal = Signal(int)
    injection_finished_signal = Signal(str)
    drill_round_finished_signal = Signal(str)
    benchmark_finished_signal = Signal(str)
    syringe_position_signal = Signal(object)
    syringe_limit_warning_signal = Signal(str)
    block_prompt_signal = Signal()
    usb_probe_log_signal = Signal(str)
    usb_probe_finished_signal = Signal(str)
    validation_move_position_signal = Signal(object)

    def __init__(self) -> None:
        super().__init__()
        self.controller = StereoDriveController()
        self.seeds: list[SeedPoint] = []
        self.trajectory: list[tuple[float, float, float]] = []
        self.drilled_depths: list[float] = []
        self.frozen_points: list[bool] = []
        self.current_seed_index: int | None = None
        self.current_action = "No trajectory yet"
        self.move_speed_step_mm = DEFAULT_MOVE_SPEED_MM
        self.movement_key_bindings = {
            "speed_decrease": int(Qt.Key.Key_Shift),
            "speed_increase": int(Qt.Key.Key_Ccedilla),
            "ml_left": int(Qt.Key.Key_Left),
            "ml_right": int(Qt.Key.Key_Right),
            "ap_anterior": int(Qt.Key.Key_Up),
            "ap_posterior": int(Qt.Key.Key_Down),
            "dv_up": int(Qt.Key.Key_PageUp),
            "dv_down": int(Qt.Key.Key_PageDown),
        }
        self.syringe_key_bindings = {
            "volume_down": int(Qt.Key.Key_F1),
            "volume_up": int(Qt.Key.Key_F2),
            "syringe_up": int(Qt.Key.Key_F3),
            "syringe_down": int(Qt.Key.Key_F4),
            "stop_injection": int(Qt.Key.Key_Escape),
        }
        self.drill_pause_requested = threading.Event()
        self.drill_stop_requested = threading.Event()
        self.drill_thread: threading.Thread | None = None
        self.drill_completed_points = 0
        self.drill_round_started_at: float | None = None
        self.drill_round_target_seconds: float = 0.0
        self.active_surface_dv: float | None = None
        self.active_depth_ratio: float | None = None
        self.active_drill_depth_mm: float | None = None
        self.current_target_depth_mm = 0.0
        self.drilling_paused = False
        self.benchmark_thread: threading.Thread | None = None
        self.usb_probe_thread: threading.Thread | None = None
        self.manual_injection_volume_nl = DEFAULT_INJECTION_VOLUME_NL
        self.syringe_position_nl: float | None = None
        self.syringe_position_lock = threading.Lock()
        self.coordinate_mode = "axis"
        self.craniotomy_coordinate_system = "bregma"
        self.injection_sites_coordinate_system = "bregma"
        self.bregma_axis: tuple[float, float, float] | None = None
        self.anchor_axis: tuple[float, float, float] | None = None
        self.anchor_bregma: tuple[float, float, float] | None = None
        self.injection_sites: list[InjectionSite] = []
        self.recent_injection_grid_configs: list[dict[str, object]] = []
        self.nudge_all_sites_active = False
        self.validation_modal_active = False
        self.validation_move_active = False
        self.validation_move_cancel_callback = None
        self.quick_locations: dict[str, StoredLocation] = {}
        self.named_locations: dict[str, StoredLocation] = {}
        self.injection_thread: threading.Thread | None = None
        self.injection_pause_requested = threading.Event()
        self.injection_stop_requested = threading.Event()
        self.block_prompt_event: threading.Event | None = None
        self.block_prompt_result = "clear"
        self.warning_auto_confirm_stop = threading.Event()
        self.setWindowTitle("Craniotomy Planner")
        self.setFocusPolicy(Qt.StrongFocus)
        self.resize(1120, 760)
        self.status_signal.connect(self.set_status)
        self.redraw_signal.connect(self.redraw_views)
        self.drill_progress_signal.connect(self.set_drill_completed_points)
        self.injection_progress_signal.connect(self.set_injection_progress)
        self.injection_site_progress_signal.connect(self.set_injection_site_progress)
        self.sequence_step_signal.connect(self.set_active_sequence_step)
        self.active_injection_site_signal.connect(self.set_active_injection_site)
        self.injection_finished_signal.connect(self.finish_injection)
        self.drill_round_finished_signal.connect(self.on_drill_round_finished)
        self.benchmark_finished_signal.connect(self.show_benchmark_results)
        self.syringe_position_signal.connect(self.set_syringe_position)
        self.syringe_limit_warning_signal.connect(self.show_syringe_limit_warning)
        self.block_prompt_signal.connect(self.show_block_prompt)
        self.usb_probe_log_signal.connect(self._append_usb_probe_log)
        self.usb_probe_finished_signal.connect(self._finish_usb_probe)
        self.validation_move_position_signal.connect(self._set_validation_move_position)
        self._build_ui()
        self._load_last_used_configs()
        self._load_general_settings()
        self._offer_project_session_restore()
        self.restore_saved_window_geometry()
        QApplication.instance().installEventFilter(self)
        self.refresh_live_position()
        QTimer.singleShot(250, self.update_syringe_position_from_scale)
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.refresh_live_position)
        self.refresh_timer.start(50)
        self.project_session_timer = QTimer(self)
        self.project_session_timer.timeout.connect(self._autosave_project_session)
        self.project_session_timer.start(1000)
        self._start_warning_auto_confirm_watcher()

    def _start_warning_auto_confirm_watcher(self) -> None:
        def worker() -> None:
            while not self.warning_auto_confirm_stop.is_set():
                try:
                    self.controller.confirm_below_skull_warning(timeout_seconds=0.0)
                    self.controller.confirm_no_actual_movement_dialog(timeout_seconds=0.0)
                except Exception:
                    pass
                self.warning_auto_confirm_stop.wait(0.05)

        threading.Thread(target=worker, daemon=True).start()

    def closeEvent(self, event) -> None:  # noqa: N802
        if self._motion_is_active():
            self.stop_motion()
            try:
                self.controller.wait_until_stopped()
            except Exception as exc:
                QMessageBox.warning(self, "Close", str(exc))
                event.ignore()
                return
            # Keep the process alive until cancelled workers have unwound.
            workers = (self.drill_thread, self.injection_thread, self.benchmark_thread, self.usb_probe_thread)
            for worker in workers:
                if worker is not None:
                    worker.join(timeout=0.5)
            if any(worker is not None and worker.is_alive() for worker in workers):
                self.set_status("Stopping operations. Close again once they have finished.")
                event.ignore()
                return
        try:
            self.window_geometry = {
                "x": self.x(), "y": self.y(),
                "width": self.width(), "height": self.height(),
            }
            self._save_last_used_configs()
            self._save_general_settings()
            self._autosave_project_session()
        except Exception:
            pass
        self.warning_auto_confirm_stop.set()
        super().closeEvent(event)

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        root.setStyleSheet(
            """
            QWidget {
                background: #eef3ea;
                color: #173122;
                font-family: 'Segoe UI';
                font-size: 11px;
            }
            QGroupBox {
                border: 1px solid #d4ded3;
                border-radius: 7px;
                margin-top: 4px;
                background: rgba(255,255,255,0.92);
                font-weight: 600;
                padding-top: 4px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 8px;
                padding: 0 3px;
            }
            QPushButton {
                border: none;
                border-radius: 6px;
                padding: 1px 6px;
                background: #dceae0;
                min-height: 17px;
            }
            QPushButton:hover {
                background: #cfe3d6;
            }
            QPushButton[variant="primary"] {
                background: #0d8a63;
                color: white;
            }
            QPushButton[variant="danger"] {
                background: #b23a48;
                color: white;
            }
            QPushButton[variant="quick-green"] {
                background: #108a54;
                color: white;
            }
            QPushButton[variant="quick-blue"] {
                background: #1f6fbf;
                color: white;
            }
            QPushButton[variant="quick-yellow"] {
                background: #f1c232;
                color: #2d2600;
            }
            QLabel[role="hero"] {
                font-size: 34px;
                font-weight: 700;
            }
            QLabel[role="muted"] {
                color: #5e7064;
            }
            QLabel[role="coord"] {
                background: rgba(255,255,255,0.92);
                border: 1px solid #d4ded3;
                border-radius: 8px;
                font-size: 12px;
                font-weight: 700;
                padding: 2px 8px;
            }
            QLineEdit {
                border: 1px solid #cfdbcf;
                border-radius: 6px;
                padding: 1px 3px;
                background: white;
                min-height: 14px;
            }
            """
        )

        layout = QVBoxLayout(root)
        layout.setContentsMargins(8, 6, 8, 8)
        layout.setSpacing(4)

        self.current_ap_label = QLabel("-")
        self.current_ml_label = QLabel("-")
        self.current_dv_label = QLabel("-")
        self.action_status_label = QLabel(self.current_action)
        self.action_status_label.setWordWrap(True)
        self.action_status_label.setMinimumHeight(34)
        self.action_status_label.setStyleSheet(
            "font-size: 22px; font-weight: 800; color: #173122; "
            "background: rgba(255,255,255,0.72); border: 1px solid #d4ded3; "
            "border-radius: 7px; padding: 4px 8px;"
        )
        self.move_speed_label = QLabel()
        for label in (self.current_ap_label, self.current_ml_label, self.current_dv_label):
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            label.setProperty("role", "coord")
            label.setMinimumWidth(82)

        header_container = QVBoxLayout()
        header_container.setSpacing(3)
        position_layout = QHBoxLayout()
        position_layout.setSpacing(5)
        header_layout = QHBoxLayout()
        header_layout.setSpacing(5)
        set_bregma_btn = QPushButton("Set Bregma")
        set_bregma_btn.clicked.connect(self.set_current_location_to_bregma)
        bregma_btn = QPushButton("Bregma")
        bregma_btn.clicked.connect(self.goto_bregma)
        home_btn = QPushButton("Home")
        home_btn.clicked.connect(self.goto_home)
        work_btn = QPushButton("Work")
        work_btn.clicked.connect(self.goto_work)
        go_to_btn = QPushButton("Go to")
        go_to_btn.clicked.connect(self.open_goto_dialog)
        header_stop_btn = QPushButton("Stop")
        header_stop_btn.setProperty("variant", "danger")
        header_stop_btn.style().unpolish(header_stop_btn)
        header_stop_btn.style().polish(header_stop_btn)
        header_stop_btn.clicked.connect(self.stop_motion)
        drill_toggle_btn = QPushButton("Drill On/Off")
        drill_toggle_btn.clicked.connect(self.activate_stereodrive_drill)
        clear_project_btn = QPushButton("Clear Project")
        clear_project_btn.clicked.connect(self.clear_project)
        set_bregma_local_btn = QPushButton("Set Bregma")
        set_bregma_local_btn.clicked.connect(self.set_local_bregma)
        set_anchor_btn = QPushButton("Set Anchor")
        set_anchor_btn.clicked.connect(self.set_anchor)
        at_anchor_btn = QPushButton("At Anchor")
        at_anchor_btn.clicked.connect(self.at_anchor)
        self.axis_mode_btn = QPushButton("Axis")
        self.bregma_mode_btn = QPushButton("Bregma")
        self.axis_mode_btn.clicked.connect(lambda: self.set_coordinate_mode("axis"))
        self.bregma_mode_btn.clicked.connect(lambda: self.set_coordinate_mode("bregma"))
        position_layout.addWidget(self.axis_mode_btn)
        position_layout.addWidget(self.bregma_mode_btn)
        quick_specs = (
            ("A", "quick-green"),
            ("B", "quick-blue"),
            ("C", "quick-yellow"),
        )
        quick_buttons: list[QPushButton] = []
        for name, variant in quick_specs:
            set_btn = QPushButton(f"Set {name}")
            set_btn.setProperty("variant", variant)
            set_btn.clicked.connect(lambda _checked=False, slot=name: self.set_quick_location(slot))
            quick_goto_btn = QPushButton(name)
            quick_goto_btn.setProperty("variant", variant)
            quick_goto_btn.clicked.connect(lambda _checked=False, slot=name: self.goto_quick_location(slot))
            for button in (set_btn, quick_goto_btn):
                button.style().unpolish(button)
                button.style().polish(button)
                quick_buttons.append(button)
        position_layout.addWidget(QLabel("AP"))
        position_layout.addWidget(self.current_ap_label)
        position_layout.addWidget(QLabel("ML"))
        position_layout.addWidget(self.current_ml_label)
        position_layout.addWidget(QLabel("DV"))
        position_layout.addWidget(self.current_dv_label)
        position_layout.addWidget(self.move_speed_label)
        position_layout.addWidget(header_stop_btn)
        position_layout.addWidget(drill_toggle_btn)
        position_layout.addWidget(clear_project_btn)
        position_layout.addStretch(1)
        header_layout.addWidget(set_bregma_local_btn)
        header_layout.addWidget(set_anchor_btn)
        header_layout.addWidget(at_anchor_btn)
        header_layout.addWidget(QLabel("Go to:"))
        header_layout.addWidget(bregma_btn)
        header_layout.addWidget(home_btn)
        header_layout.addWidget(work_btn)
        header_layout.addWidget(go_to_btn)
        for button in quick_buttons:
            header_layout.addWidget(button)
        header_layout.addStretch(1)
        options_btn = QPushButton("Options")
        options_btn.clicked.connect(self.open_options_dialog)
        header_layout.addWidget(options_btn)
        update_btn = QPushButton("Update")
        update_btn.clicked.connect(self.update_from_github)
        header_layout.addWidget(update_btn)
        header_layout.addWidget(QLabel("Overlay:"))
        self.overlay_combo = QComboBox()
        self.overlay_combo.addItem("None", None)
        overlay_dir = Path(__file__).resolve().parents[1] / "assets" / "background_images"
        for image_path in sorted(overlay_dir.glob("*.png")):
            self.overlay_combo.addItem(image_path.stem, str(image_path))
        self.overlay_combo.currentIndexChanged.connect(self.select_overlay)
        header_layout.addWidget(self.overlay_combo)
        header_container.addLayout(position_layout)
        header_container.addLayout(header_layout)
        header_container.addWidget(self.action_status_label)
        layout.addLayout(header_container)

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, 1)

        craniotomy_tab = QWidget()
        content = QVBoxLayout(craniotomy_tab)
        content.setSpacing(4)
        self.tabs.addTab(craniotomy_tab, "Craniotomy")

        self._build_options_dialog()

        setup_box = QGroupBox("Setup")
        setup_layout = QGridLayout(setup_box)
        setup_layout.setContentsMargins(7, 6, 7, 7)
        setup_layout.setHorizontalSpacing(8)
        setup_layout.setVerticalSpacing(3)
        content.addWidget(setup_box)

        self.mid_ap = self._double_spinbox()
        self.mid_ml = self._double_spinbox()
        self.diameter = self._double_spinbox(value=3.2, minimum=0.1, maximum=20.0)
        self.seed_count = self._spinbox(value=6, minimum=3, maximum=24)
        self.trajectory_points = self._spinbox(value=60, minimum=12, maximum=360)
        self.cut_offset = self._double_spinbox(value=0.0, minimum=-5.0, maximum=5.0)
        self.drill_depth = self._double_spinbox(value=0.20, minimum=0.0, maximum=5.0)
        self.depth_per_round = self._double_spinbox(value=0.05, minimum=0.001, maximum=5.0)
        self.skull_thickness_mm = self._double_spinbox(value=0.25, minimum=0.001, maximum=5.0)
        self.round_time_seconds = self._double_spinbox(value=60.0, minimum=1.0, maximum=3600.0)
        self.drill_rate_mm_per_s = self._double_spinbox(value=0.01, minimum=0.001, maximum=5.0)
        self.auto_start_rounds = QCheckBox("Auto start next round")
        self.auto_start_rounds.setChecked(True)
        self.current_seed_spin = self._spinbox(value=1, minimum=1, maximum=1)
        self.current_seed_spin.valueChanged.connect(self.on_seed_spin_changed)
        self.current_seed_coords = QLabel("Seed: -")
        self.current_seed_coords.setWordWrap(True)
        self.current_seed_coords.setProperty("role", "muted")
        self.craniotomy_save_btn = QPushButton("Save Craniotomy Config")
        self.craniotomy_save_btn.clicked.connect(self.save_craniotomy_config)
        self.craniotomy_load_btn = QPushButton("Load Craniotomy Config")
        self.craniotomy_load_btn.clicked.connect(self.load_craniotomy_config)
        self.set_center_btn = QPushButton("Set Center")
        self.set_center_btn.clicked.connect(self.set_craniotomy_center)

        setup_layout.addWidget(QLabel("Mid AP"), 0, 0)
        setup_layout.addWidget(self.mid_ap, 0, 1)
        setup_layout.addWidget(QLabel("Mid ML"), 0, 2)
        setup_layout.addWidget(self.mid_ml, 0, 3)
        setup_layout.addWidget(self.craniotomy_load_btn, 0, 4)
        setup_layout.addWidget(self.set_center_btn, 0, 5)
        setup_layout.addWidget(self.craniotomy_save_btn, 0, 6)

        setup_layout.addWidget(QLabel("Diameter (mm)"), 1, 0)
        setup_layout.addWidget(self.diameter, 1, 1)
        setup_layout.addWidget(QLabel("Seed Points"), 1, 2)
        setup_layout.addWidget(self.seed_count, 1, 3)
        setup_layout.addWidget(QLabel("Trajectory Points"), 1, 4)
        setup_layout.addWidget(self.trajectory_points, 1, 5)

        setup_layout.addWidget(QLabel("Cut Offset DV"), 2, 0)
        setup_layout.addWidget(self.cut_offset, 2, 1)
        setup_layout.addWidget(QLabel("Max Depth"), 2, 2)
        setup_layout.addWidget(self.drill_depth, 2, 3)
        setup_layout.addWidget(QLabel("Skull Thickness (mm)"), 2, 4)
        setup_layout.addWidget(self.skull_thickness_mm, 2, 5)

        setup_layout.addWidget(QLabel("Current Seed"), 3, 0)
        setup_layout.addWidget(self.current_seed_spin, 3, 1)
        setup_layout.addWidget(QLabel("Drill Rate (mm/s)"), 3, 2)
        setup_layout.addWidget(self.drill_rate_mm_per_s, 3, 3)
        setup_layout.addWidget(QLabel("Round Time (s)"), 3, 4)
        setup_layout.addWidget(self.round_time_seconds, 3, 5)
        setup_layout.addWidget(QLabel("Depth per round"), 4, 0)
        setup_layout.addWidget(self.depth_per_round, 4, 1)
        setup_layout.addWidget(self.auto_start_rounds, 4, 2, 1, 2)

        generate_btn = QPushButton("Generate Seeds")
        generate_btn.clicked.connect(self.generate_seeds)
        self.move_seed_btn = QPushButton("Next seed")
        self.move_seed_btn.clicked.connect(self.move_to_current_seed)
        self.capture_surface_btn = QPushButton("Set Surface")
        self.capture_surface_btn.setProperty("variant", "primary")
        self.capture_surface_btn.style().unpolish(self.capture_surface_btn)
        self.capture_surface_btn.style().polish(self.capture_surface_btn)
        self.capture_surface_btn.clicked.connect(self.capture_surface)
        stop_btn = QPushButton("Stop Motion")
        stop_btn.setProperty("variant", "danger")
        stop_btn.style().unpolish(stop_btn)
        stop_btn.style().polish(stop_btn)
        stop_btn.clicked.connect(self.stop_motion)
        clear_btn = QPushButton("Clear Surface Measurements")
        clear_btn.clicked.connect(self.clear_surface_measurements)
        clear_craniotomy_btn = QPushButton("Clear Craniotomy")
        clear_craniotomy_btn.clicked.connect(self.clear_craniotomy)
        self.start_round_btn = QPushButton("Start Drilling")
        self.start_round_btn.setProperty("variant", "primary")
        self.start_round_btn.style().unpolish(self.start_round_btn)
        self.start_round_btn.style().polish(self.start_round_btn)
        self.start_round_btn.clicked.connect(self.start_drilling_round)
        self.freeze_draw_btn = QPushButton("Draw Freeze")
        self.freeze_draw_btn.setCheckable(True)
        self.freeze_draw_btn.toggled.connect(self.toggle_freeze_mode)
        self.unfreeze_draw_btn = QPushButton("Draw Unfreeze")
        self.unfreeze_draw_btn.setCheckable(True)
        self.unfreeze_draw_btn.toggled.connect(self.toggle_unfreeze_mode)
        self.clear_freeze_btn = QPushButton("Clear Freeze")
        self.clear_freeze_btn.clicked.connect(self.clear_frozen_points)
        setup_layout.addWidget(self.current_seed_coords, 5, 0, 1, 6)

        button_layout = QGridLayout()
        button_layout.setHorizontalSpacing(6)
        button_layout.setVerticalSpacing(3)
        button_layout.addWidget(generate_btn, 0, 0)
        button_layout.addWidget(clear_btn, 0, 1)
        button_layout.addWidget(clear_craniotomy_btn, 0, 2)
        button_layout.addWidget(self.move_seed_btn, 1, 0)
        button_layout.addWidget(self.capture_surface_btn, 1, 1)
        button_layout.addWidget(self.start_round_btn, 1, 2)
        button_layout.addWidget(self.freeze_draw_btn, 2, 0)
        button_layout.addWidget(self.clear_freeze_btn, 2, 1)
        button_layout.addWidget(self.unfreeze_draw_btn, 2, 2)
        button_layout.addWidget(stop_btn, 3, 0, 1, 3)
        setup_layout.addLayout(button_layout, 6, 0, 1, 6)

        views_box = QGroupBox()
        views_layout = QGridLayout(views_box)
        views_layout.setContentsMargins(7, 6, 7, 7)
        content.addWidget(views_box, 1)

        self.top_view = ProjectionWidget("ML", "AP")
        self.top_view.freeze_drawn.connect(self.mark_frozen_point)
        self.top_view.unfreeze_drawn.connect(self.unmark_frozen_point)
        self.top_view.location_double_clicked.connect(self.move_to_map_location)
        self.top_view.set_navigation_enabled(True)
        self.top_view.setMinimumSize(420, 420)
        self.top_view.setMaximumWidth(620)
        map_layout = QVBoxLayout()
        map_layout.addWidget(self.top_view)
        self.zoom_mode_combo = QComboBox()
        self.zoom_mode_combo.addItems(["Zoom to craniotomy", "Zoom to mid-range", "Zoom to skull"])
        self.zoom_mode_combo.currentIndexChanged.connect(self.set_zoom_mode)
        map_layout.addWidget(self.zoom_mode_combo)
        views_layout.addLayout(map_layout, 0, 0)
        legend_layout = QVBoxLayout()
        legend_layout.setSpacing(3)
        self.depth_legend = DepthLegendWidget()
        self.depth_legend.set_skull_thickness_mm(self.skull_thickness_mm.value())
        legend_layout.addWidget(self.depth_legend, 0, Qt.AlignTop)
        self.round_elapsed_label = QLabel("Elapsed: --:--")
        self.round_remaining_label = QLabel("Remaining: --:--")
        self.round_percent_label = QLabel("Complete: --%")
        self.current_target_depth_label = QLabel("Current Target Depth: -- mm")
        self.change_target_depth_btn = QPushButton("Change")
        self.change_target_depth_btn.clicked.connect(self.change_current_target_depth)
        self.round_elapsed_label.setProperty("role", "muted")
        self.round_remaining_label.setProperty("role", "muted")
        self.round_percent_label.setProperty("role", "muted")
        self.current_target_depth_label.setProperty("role", "muted")
        legend_layout.addWidget(self.round_elapsed_label)
        legend_layout.addWidget(self.round_remaining_label)
        legend_layout.addWidget(self.round_percent_label)
        legend_layout.addWidget(self.current_target_depth_label)
        legend_layout.addWidget(self.change_target_depth_btn)
        legend_layout.addStretch(1)
        views_layout.addLayout(legend_layout, 0, 1)
        self.update_current_target_depth_label()
        self.update_move_speed_label()
        self._build_injection_tab()
        default_overlay = "skull_bregma_lambda_reference"
        default_index = self.overlay_combo.findText(default_overlay)
        if default_index >= 0:
            self.overlay_combo.setCurrentIndex(default_index)

    def _build_injection_tab(self) -> None:
        injection_tab = QWidget()
        outer_layout = QHBoxLayout(injection_tab)
        outer_layout.setContentsMargins(7, 6, 7, 7)
        outer_layout.setSpacing(6)
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        outer_layout.addLayout(layout, 1)
        self.plunger_gauge = PlungerGaugeWidget()
        outer_layout.addWidget(self.plunger_gauge)
        self.tabs.addTab(injection_tab, "Injection")

        status_box = QGroupBox("Manual Control")
        status_layout = QGridLayout(status_box)
        status_layout.setContentsMargins(7, 6, 7, 7)
        status_layout.setHorizontalSpacing(8)
        status_layout.setVerticalSpacing(3)
        layout.addWidget(status_box)

        self.manual_volume_label = QLabel()
        self.syringe_position_label = QLabel("Syringe position = -- nl")
        self.syringe_position_label.setProperty("role", "muted")
        self.injection_rate_label = QLabel("Current volume rate = -- nl/min")
        self.injection_rate_label.setProperty("role", "muted")
        self.manual_volume_combo = QComboBox()
        for volume in INJECTION_VOLUME_OPTIONS_NL:
            self.manual_volume_combo.addItem(f"{volume} nl", volume)
        self.manual_volume_combo.setCurrentIndex(INJECTION_VOLUME_OPTIONS_NL.index(self.manual_injection_volume_nl))
        self.manual_volume_combo.currentIndexChanged.connect(self.on_manual_volume_combo_changed)
        self.inject_up_btn = QPushButton("Step Syringe Up (F3)")
        self.inject_up_btn.clicked.connect(lambda: self.manual_syringe_step(up=True))
        self.inject_down_btn = QPushButton("Step Syringe Down (F4)")
        self.inject_down_btn.clicked.connect(lambda: self.manual_syringe_step(up=False))
        self.manual_stop_btn = QPushButton("Stop")
        self.manual_stop_btn.setProperty("variant", "danger")
        self.manual_stop_btn.style().unpolish(self.manual_stop_btn)
        self.manual_stop_btn.style().polish(self.manual_stop_btn)
        self.manual_stop_btn.clicked.connect(self.stop_injection)
        self.empty_syringe_btn = QPushButton("Empty Syringe")
        self.empty_syringe_btn.clicked.connect(self.empty_syringe)
        update_syringe_position_btn = QPushButton("Update Syringe Position")
        update_syringe_position_btn.clicked.connect(self.update_syringe_position_from_scale)
        test_blockage_btn = QPushButton("Test for Blockage")
        test_blockage_btn.clicked.connect(self.test_for_blockage)

        status_layout.addWidget(self.syringe_position_label, 0, 0, 1, 4)
        status_layout.addWidget(QLabel("Current manual injection volume"), 1, 0)
        status_layout.addWidget(self.manual_volume_label, 1, 1, 1, 3)
        status_layout.addWidget(self.injection_rate_label, 2, 0, 1, 4)
        status_layout.addWidget(self.manual_volume_combo, 3, 0, 1, 4)
        status_layout.addWidget(self.inject_up_btn, 4, 0)
        status_layout.addWidget(self.inject_down_btn, 4, 1)
        status_layout.addWidget(self.manual_stop_btn, 4, 2)
        status_layout.addWidget(self.empty_syringe_btn, 4, 3)
        status_layout.addWidget(update_syringe_position_btn, 5, 0, 1, 2)
        status_layout.addWidget(test_blockage_btn, 5, 2, 1, 2)

        single_box = QGroupBox("Injection")
        single_layout = QGridLayout(single_box)
        single_layout.setContentsMargins(7, 6, 7, 7)
        single_layout.setHorizontalSpacing(8)
        single_layout.setVerticalSpacing(3)
        layout.addWidget(single_box)

        self.single_injection_volume_nl = self._number_edit(100)
        self.single_injection_volume_nl.editingFinished.connect(self.round_single_injection_volume_up)
        self.insertion_injection_rate_nl_min = self._number_edit(100.0)
        self.main_injection_rate_nl_min = self._number_edit(100.0)
        self.injection_depth_mm = self._number_edit(0.2)
        self.insert_retract_speed_um_s = self._number_edit(20.0)
        self.movement_overshoot_mm = self._number_edit(0.05)
        self.post_inject_pause_s = self._number_edit(5.0)
        self.single_injection_volume_nl.textChanged.connect(self.update_injection_rate_label)
        self.insertion_injection_rate_nl_min.textChanged.connect(self.update_injection_rate_label)
        self.main_injection_rate_nl_min.textChanged.connect(self.update_injection_rate_label)
        self.block_test_volume_nl = self._number_edit(50)
        self.block_test_volume_nl.editingFinished.connect(self.round_test_volume_to_supported)
        for widget in (
            self.single_injection_volume_nl,
            self.insertion_injection_rate_nl_min,
            self.main_injection_rate_nl_min,
            self.injection_depth_mm,
            self.insert_retract_speed_um_s,
            self.movement_overshoot_mm,
            self.post_inject_pause_s,
            self.block_test_volume_nl,
        ):
            widget.textChanged.connect(self.refresh_injection_sequence_summary)
        self.injection_progress = QProgressBar()
        self.injection_progress.setRange(0, 100)
        self.injection_progress.setValue(0)
        self.injection_site_progress = QProgressBar()
        self.injection_site_progress.setRange(0, 100)
        self.injection_site_progress.setValue(0)
        self.start_injection_btn = QPushButton("Go")
        self.start_injection_btn.setProperty("variant", "primary")
        self.start_injection_btn.style().unpolish(self.start_injection_btn)
        self.start_injection_btn.style().polish(self.start_injection_btn)
        self.start_injection_btn.clicked.connect(self.start_single_injection)
        self.pause_injection_btn = QPushButton("Pause")
        self.pause_injection_btn.clicked.connect(self.pause_resume_injection)
        self.stop_injection_btn = QPushButton("Stop")
        self.stop_injection_btn.setProperty("variant", "danger")
        self.stop_injection_btn.style().unpolish(self.stop_injection_btn)
        self.stop_injection_btn.style().polish(self.stop_injection_btn)
        self.stop_injection_btn.clicked.connect(self.stop_injection)
        self.injection_save_btn = QPushButton("Save Injection Settings")
        self.injection_save_btn.setToolTip("Saves protocol settings only. Injection Sites are not included.")
        self.injection_save_btn.clicked.connect(self.save_injection_config)
        self.injection_load_btn = QPushButton("Load Injection Settings")
        self.injection_load_btn.setToolTip("Loads protocol settings only. Existing Injection Sites are unchanged.")
        self.injection_load_btn.clicked.connect(self.load_injection_config)
        self.sequence_steps_list = QListWidget()
        self.sequence_steps_list.setSpacing(0)
        self.sequence_steps_list.setUniformItemSizes(True)
        self.sequence_steps_list.setStyleSheet(
            """
            QListWidget {
                font-size: 8pt;
            }
            QListWidget::item {
                margin: 0px;
                padding: 0px 2px;
                min-height: 14px;
            }
            """
        )
        single_layout.addWidget(QLabel("Main injection volume (nl)"), 0, 0)
        single_layout.addWidget(self.single_injection_volume_nl, 0, 1)
        single_layout.addWidget(QLabel("Injection rate (nl/min)"), 0, 2)
        single_layout.addWidget(self.main_injection_rate_nl_min, 0, 3)
        single_layout.addWidget(QLabel("Insertion injection rate (nl/min)"), 0, 4)
        single_layout.addWidget(self.insertion_injection_rate_nl_min, 0, 5)
        single_layout.addWidget(QLabel("Injection depth (mm)"), 1, 0)
        single_layout.addWidget(self.injection_depth_mm, 1, 1)
        single_layout.addWidget(QLabel("Insert/retract speed (um/sec)"), 1, 2)
        single_layout.addWidget(self.insert_retract_speed_um_s, 1, 3)
        single_layout.addWidget(QLabel("Overshoot (mm)"), 1, 4)
        single_layout.addWidget(self.movement_overshoot_mm, 1, 5)
        single_layout.addWidget(QLabel("Post inject pause (s)"), 2, 0)
        single_layout.addWidget(self.post_inject_pause_s, 2, 1)
        single_layout.addWidget(QLabel("Test volume (nl)"), 2, 2)
        single_layout.addWidget(self.block_test_volume_nl, 2, 3)
        single_layout.addWidget(self.injection_load_btn, 2, 4)
        single_layout.addWidget(self.injection_save_btn, 2, 5)
        settings_note = QLabel("Load/save applies to injection settings only; Injection Sites are never changed.")
        settings_note.setProperty("role", "muted")
        single_layout.addWidget(settings_note, 3, 0, 1, 6)
        single_layout.addWidget(QLabel("Program sequence"), 4, 0, 1, 6)
        single_layout.addWidget(self.sequence_steps_list, 5, 0, 1, 6)
        single_layout.addWidget(QLabel("Overall sequence progress"), 6, 0)
        single_layout.addWidget(self.injection_progress, 6, 1, 1, 5)
        single_layout.addWidget(QLabel("Current injection/movement"), 7, 0)
        single_layout.addWidget(self.injection_site_progress, 7, 1, 1, 5)
        single_layout.addWidget(self.start_injection_btn, 8, 0)
        single_layout.addWidget(self.pause_injection_btn, 8, 1)
        single_layout.addWidget(self.stop_injection_btn, 8, 2, 1, 4)

        self._build_injection_sites_section(layout)
        self.update_manual_volume_label()
        self.update_injection_rate_label()
        self.refresh_injection_sequence_summary()

    def _build_injection_sites_section(self, parent_layout: QVBoxLayout) -> None:
        """Build the bottom half of Injection with map and site list side by side."""
        sites_section = QWidget()
        sites_outer_layout = QHBoxLayout(sites_section)
        sites_outer_layout.setContentsMargins(0, 0, 0, 0)
        sites_outer_layout.setSpacing(8)
        parent_layout.addWidget(sites_section, 1)

        map_box = QGroupBox("Map")
        map_layout = QVBoxLayout(map_box)
        self.injection_sites_view = ProjectionWidget("ML", "AP")
        self.injection_sites_view.location_double_clicked.connect(self.move_to_map_location)
        self.injection_sites_view.set_navigation_enabled(True)
        self.injection_sites_view.set_coordinate_mode_bregma(self.coordinate_mode == "bregma")
        self.injection_sites_view.set_overlay_image(self.top_view.overlay_image, self.top_view.overlay_calibration)
        # This lives in the lower half of the Injection tab, so keep it usable
        # without forcing the tab taller than a normal application window.
        self.injection_sites_view.setMinimumSize(260, 260)
        map_layout.addWidget(self.injection_sites_view, 1)
        injection_map_controls = QHBoxLayout()
        self.show_craniotomy_on_injection_map = QCheckBox("Show craniotomy")
        self.show_craniotomy_on_injection_map.setChecked(True)
        self.show_craniotomy_on_injection_map.toggled.connect(lambda _checked: self.redraw_views())
        injection_map_controls.addWidget(self.show_craniotomy_on_injection_map)
        injection_map_controls.addStretch(1)
        self.injection_sites_zoom_combo = QComboBox()
        self.injection_sites_zoom_combo.addItems([
            "Zoom To Craniotomy",
            "Zoom To Injection Map",
            "Zoom To Skull",
        ])
        self.injection_sites_zoom_combo.currentIndexChanged.connect(self.set_injection_sites_zoom_mode)
        injection_map_controls.addWidget(self.injection_sites_zoom_combo)
        map_layout.addLayout(injection_map_controls)
        sites_outer_layout.addWidget(map_box, 1)

        sites_box = QGroupBox("Injection Sites")
        sites_layout = QGridLayout(sites_box)
        sites_layout.setContentsMargins(7, 6, 7, 7)
        sites_layout.setHorizontalSpacing(8)
        sites_layout.setVerticalSpacing(3)
        sites_outer_layout.addWidget(sites_box, 1)

        self.injection_sites_list = QListWidget()
        add_site_btn = QPushButton("Add Injection Site")
        add_site_btn.clicked.connect(self.add_injection_site)
        add_grid_btn = QPushButton("Add Grid")
        add_grid_btn.clicked.connect(self.add_injection_site_grid)
        remove_site_btn = QPushButton("Remove Selected Site")
        remove_site_btn.clicked.connect(self.remove_selected_injection_site)
        validate_sites_btn = QPushButton("Validate Sites")
        validate_sites_btn.setProperty("variant", "primary")
        validate_sites_btn.style().unpolish(validate_sites_btn)
        validate_sites_btn.style().polish(validate_sites_btn)
        validate_sites_btn.clicked.connect(self.start_injection_site_validation)
        clear_sites_btn = QPushButton("Clear Sites")
        clear_sites_btn.clicked.connect(self.clear_injection_sites)
        save_site_set_btn = QPushButton("Save Site Set")
        save_site_set_btn.clicked.connect(self.save_injection_site_set)
        load_site_set_btn = QPushButton("Load Site Set")
        load_site_set_btn.clicked.connect(self.load_injection_site_set)
        self.nudge_all_sites_btn = QPushButton("Nudge All Sites")
        self.nudge_all_sites_btn.setCheckable(True)
        self.nudge_all_sites_btn.setToolTip(
            "When enabled, AP/ML arrow shortcuts shift every injection site by the current move speed. "
            "The manipulator does not move."
        )
        self.nudge_all_sites_btn.toggled.connect(self.set_nudge_all_sites_active)
        resume_selected_btn = QPushButton("Start From Selected")
        resume_selected_btn.setProperty("variant", "quick-green")
        resume_selected_btn.style().unpolish(resume_selected_btn)
        resume_selected_btn.style().polish(resume_selected_btn)
        resume_selected_btn.clicked.connect(self.resume_injection_from_selected)
        self.block_check = QCheckBox("Check blockage after each site")
        self.block_check.setChecked(True)
        self.block_check.toggled.connect(self.refresh_injection_sequence_summary)
        sites_layout.addWidget(add_site_btn, 0, 0)
        sites_layout.addWidget(add_grid_btn, 0, 1)
        sites_layout.addWidget(remove_site_btn, 0, 2)
        sites_layout.addWidget(save_site_set_btn, 1, 0)
        sites_layout.addWidget(load_site_set_btn, 1, 1)
        sites_layout.addWidget(self.nudge_all_sites_btn, 1, 2)
        sites_layout.addWidget(validate_sites_btn, 2, 0)
        sites_layout.addWidget(clear_sites_btn, 2, 1)
        sites_layout.addWidget(resume_selected_btn, 2, 2)
        sites_layout.addWidget(self.block_check, 3, 0, 1, 3)
        sites_layout.addWidget(self.injection_sites_list, 4, 0, 1, 3)

    def _build_options_dialog(self) -> None:
        self.options_dialog = QDialog(self)
        self.options_dialog.setWindowTitle("Options")
        self.options_dialog.resize(760, 860)
        options_layout = QVBoxLayout(self.options_dialog)
        options_box = QGroupBox("Keyboard Controls")
        options_grid = QGridLayout(options_box)
        options_grid.addWidget(QLabel("Assign a single key or key combination. Changes save when the app closes."), 0, 0, 1, 2)
        self.movement_key_edits = {}
        key_options = (("speed_decrease", "Decrease movement step"), ("speed_increase", "Increase movement step"), ("ml_left", "ML left"), ("ml_right", "ML right"), ("ap_anterior", "AP anterior"), ("ap_posterior", "AP posterior"), ("dv_up", "DV up"), ("dv_down", "DV down"), ("volume_down", "Decrease injection volume"), ("volume_up", "Increase injection volume"), ("syringe_up", "Syringe step up"), ("syringe_down", "Syringe step down"), ("stop_injection", "Stop injection"))
        for row, (name, label) in enumerate(key_options, start=1):
            bindings = self.syringe_key_bindings if name in self.syringe_key_bindings else self.movement_key_bindings
            edit = QKeySequenceEdit(QKeySequence(bindings[name]))
            edit.setMaximumSequenceLength(1)
            callback = self._syringe_key_sequence_changed if name in self.syringe_key_bindings else self._movement_key_sequence_changed
            edit.keySequenceChanged.connect(lambda sequence, binding_name=name, handler=callback: handler(binding_name, sequence))
            self.movement_key_edits[name] = edit
            options_grid.addWidget(QLabel(label), row, 0)
            options_grid.addWidget(edit, row, 1)
        reset_keys_btn = QPushButton("Reset keyboard shortcuts")
        reset_keys_btn.clicked.connect(self.reset_movement_key_bindings)
        options_grid.addWidget(reset_keys_btn, len(key_options) + 1, 0, 1, 2)
        options_layout.addWidget(options_box)
        scan_box = QGroupBox("StereoDrive Control Scan")
        scan_layout = QVBoxLayout(scan_box)
        scan_layout.addWidget(QLabel("Scan the current StereoDrive window to identify control IDs and labels."))
        scan_btn = QPushButton("Scan StereoDrive")
        scan_btn.clicked.connect(self.scan_stereodrive)
        scan_layout.addWidget(scan_btn)
        benchmark_btn = QPushButton("Benchmark Axis Moves")
        benchmark_btn.clicked.connect(self.start_axis_benchmark)
        scan_layout.addWidget(benchmark_btn)
        update_btn = QPushButton("Update from GitHub (discard local changes)")
        update_btn.clicked.connect(self.update_from_github)
        scan_layout.addWidget(update_btn)
        self.stereodrive_scan_output = QPlainTextEdit()
        self.stereodrive_scan_output.setPlaceholderText("Scan results will appear here for copying into chat.")
        scan_layout.addWidget(self.stereodrive_scan_output, 1)
        options_layout.addWidget(scan_box, 1)

        usb_probe_box = QGroupBox("USB Controller Probe")
        usb_probe_layout = QVBoxLayout(usb_probe_box)
        usb_probe_layout.addWidget(QLabel(
            "Use this only while recording with USBPcap/Wireshark. The app logs timestamps and performs one "
            "small, confirmed Axis nudge followed by its reversal; it does not capture, replay, or inject USB traffic."
        ))
        probe_controls = QHBoxLayout()
        probe_controls.addWidget(QLabel("Axis"))
        self.usb_probe_axis_combo = QComboBox()
        for axis in ("AP", "ML", "DV"):
            self.usb_probe_axis_combo.addItem(axis, axis)
        probe_controls.addWidget(self.usb_probe_axis_combo)
        probe_controls.addWidget(QLabel("Step"))
        self.usb_probe_step_combo = QComboBox()
        for step_mm in (0.01, 0.02, 0.05):
            self.usb_probe_step_combo.addItem(f"{step_mm:g} mm", step_mm)
        probe_controls.addWidget(self.usb_probe_step_combo)
        self.usb_probe_button = QPushButton("Run one out-and-back probe")
        self.usb_probe_button.clicked.connect(self.start_usb_controller_probe)
        probe_controls.addWidget(self.usb_probe_button)
        probe_controls.addStretch(1)
        usb_probe_layout.addLayout(probe_controls)
        self.usb_probe_output = QPlainTextEdit()
        self.usb_probe_output.setReadOnly(True)
        self.usb_probe_output.setPlaceholderText(
            "Probe log will appear here. Start USBPcap/Wireshark capture before running one probe."
        )
        self.usb_probe_output.setMinimumHeight(115)
        usb_probe_layout.addWidget(self.usb_probe_output)
        options_layout.addWidget(usb_probe_box)

    def open_options_dialog(self) -> None:
        self.options_dialog.show()
        self.options_dialog.raise_()
        self.options_dialog.activateWindow()

    def scan_stereodrive(self) -> None:
        try:
            self.stereodrive_scan_output.setPlainText(self.controller.scan_stereodrive_controls())
            self.set_status("StereoDrive control scan complete.")
        except Exception as exc:
            self.stereodrive_scan_output.setPlainText(f"Scan failed: {exc}")
            self.set_status("StereoDrive control scan failed.")

    @staticmethod
    def _usb_probe_timestamp() -> str:
        return time.strftime("%Y-%m-%d %H:%M:%S") + f".{int((time.time() % 1) * 1000):03d}"

    def _append_usb_probe_log(self, message: str) -> None:
        self.usb_probe_output.appendPlainText(message)

    def _motion_is_active(self) -> bool:
        return bool(
            (self.drill_thread is not None and self.drill_thread.is_alive())
            or (self.injection_thread is not None and self.injection_thread.is_alive())
            or (self.benchmark_thread is not None and self.benchmark_thread.is_alive())
            or (self.usb_probe_thread is not None and self.usb_probe_thread.is_alive())
            or self.validation_move_active
            or self.validation_modal_active
        )

    def _require_idle(self, title: str = "Movement") -> bool:
        if self._motion_is_active():
            QMessageBox.information(self, title, "Finish or stop the current operation first.")
            return False
        return True

    def get_bregma_position(self) -> tuple[float, float, float]:
        return self._axis_to_bregma(self.controller.get_current_axis_position())

    def _require_project_coordinates(self, kind: str) -> None:
        if self.bregma_axis is None:
            raise StereoDriveError("Set GUI Bregma before using project positions.")
        frame = getattr(self, f"{kind}_coordinate_system")
        if frame != "bregma":
            raise StereoDriveError("This older session has ambiguous coordinate references. "
                                   "Recreate the craniotomy, or clear/load a new injection site set before moving.")

    def _axis_clearance_path(self, target: tuple[float, float, float], clearance_dv: float) -> list[tuple[float, float, float]]:
        current = self.controller.get_current_axis_position()
        if all(abs(a - b) <= 0.02 for a, b in zip(current, target)):
            return [target]
        safe_dv = min(current[2], clearance_dv, target[2])
        return [(current[0], current[1], safe_dv), (target[0], target[1], safe_dv), target]

    def _approach_axis_position(self, target: tuple[float, float, float], clearance_dv: float, stop_requested) -> None:
        for position in self._axis_clearance_path(target, clearance_dv):
            self.controller.goto_axis_position(*position, delay_seconds=0.5, stop_requested=stop_requested)
            self.controller.wait_for_axis_position(*position, stop_requested=stop_requested)

    def start_usb_controller_probe(self) -> None:
        if self._motion_is_active():
            QMessageBox.warning(
                self,
                "USB Controller Probe",
                "Wait for the current drill, injection, benchmark, or probe operation to finish first.",
            )
            return
        axis = str(self.usb_probe_axis_combo.currentData())
        step_mm = float(self.usb_probe_step_combo.currentData())
        response = QMessageBox.warning(
            self,
            "Run USB Controller Probe?",
            f"Start the USBPcap/Wireshark capture first. This will nudge {axis} by {step_mm:g} mm, "
            "verify the change from StereoDrive's Axis display, then nudge it back. "
            "If the forward move cannot be verified, the app will stop and will not guess a reversal. Continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if response != QMessageBox.Yes:
            return
        self.usb_probe_button.setEnabled(False)
        self.usb_probe_output.clear()
        self.controller.prepare_motion()
        self.usb_probe_thread = threading.Thread(
            target=self._run_usb_controller_probe,
            args=(axis, step_mm),
            daemon=True,
        )
        self.usb_probe_thread.start()

    def _wait_for_probe_axis_value(
        self,
        axis: str,
        expected: float,
        tolerance_mm: float,
        timeout_seconds: float = 8.0,
    ) -> float:
        deadline = time.monotonic() + timeout_seconds
        while time.monotonic() < deadline:
            self.controller._check_motion_cancelled()
            current = self.controller.get_current_axis(axis)
            if abs(current - expected) <= tolerance_mm:
                return current
            time.sleep(0.05)
        current = self.controller.get_current_axis(axis)
        raise StereoDriveError(
            f"{axis} did not reach {expected:.4f} mm within {timeout_seconds:g} seconds "
            f"(last reading {current:.4f} mm)."
        )

    def _run_usb_controller_probe(self, axis: str, step_mm: float) -> None:
        tolerance_mm = max(0.003, step_mm * 0.25)
        try:
            start = self.controller.get_current_axis(axis)
            self.usb_probe_log_signal.emit(
                f"{self._usb_probe_timestamp()} START axis={axis} step_mm={step_mm:g} start_axis_mm={start:.4f}"
            )
            self.controller.set_nudge_step(axis, step_mm)
            self.usb_probe_log_signal.emit(f"{self._usb_probe_timestamp()} FORWARD_NUDGE axis={axis}")
            self.controller.nudge_axis(axis, True)
            forward = self._wait_for_probe_axis_value(axis, start + step_mm, tolerance_mm)
            self.usb_probe_log_signal.emit(
                f"{self._usb_probe_timestamp()} FORWARD_CONFIRMED axis={axis} axis_mm={forward:.4f}"
            )
            self.usb_probe_log_signal.emit(f"{self._usb_probe_timestamp()} REVERSE_NUDGE axis={axis}")
            self.controller.nudge_axis(axis, False)
            returned = self._wait_for_probe_axis_value(axis, start, tolerance_mm)
            self.usb_probe_log_signal.emit(
                f"{self._usb_probe_timestamp()} RETURN_CONFIRMED axis={axis} axis_mm={returned:.4f}"
            )
            self.usb_probe_finished_signal.emit("USB controller probe completed; position returned to its starting value.")
        except Exception as exc:
            try:
                self.controller.stop()
                self.controller.wait_until_stopped()
            except Exception as stop_exc:
                self.usb_probe_log_signal.emit(f"Stop confirmation failed: {stop_exc}")
            self.usb_probe_log_signal.emit(f"{self._usb_probe_timestamp()} ERROR {exc}")
            self.usb_probe_finished_signal.emit(
                "USB controller probe stopped. Check the Axis display before moving again; no unverified reversal was issued."
            )

    def _finish_usb_probe(self, status: str) -> None:
        self.usb_probe_button.setEnabled(True)
        self.set_status(status)

    def update_from_github(self) -> None:
        if not self._require_idle("Update"):
            return
        answer = QMessageBox.warning(
            self,
            "Discard Local Changes?",
            "This will fetch origin/main and discard all local tracked and untracked changes. Continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        repo_dir = Path(__file__).resolve().parents[1]
        try:
            subprocess.run(["git", "fetch", "origin"], cwd=repo_dir, check=True, capture_output=True, text=True)
            subprocess.run(["git", "reset", "--hard", "origin/main"], cwd=repo_dir, check=True, capture_output=True, text=True)
            subprocess.run(["git", "clean", "-fd"], cwd=repo_dir, check=True, capture_output=True, text=True)
            restart = QMessageBox.question(
                self,
                "Update Complete",
                "The repository was updated to origin/main. Restart the app now?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if restart == QMessageBox.Yes:
                QProcess.startDetached(sys.executable, [str(Path(__file__).resolve())])
                self.close()
        except subprocess.CalledProcessError as exc:
            detail = (exc.stderr or exc.stdout or str(exc)).strip()
            QMessageBox.critical(self, "Update Failed", detail)

    def _movement_key_sequence_changed(self, name: str, sequence: QKeySequence) -> None:
        if not sequence.isEmpty():
            self.movement_key_bindings[name] = int(sequence[0].toCombined())

    def _syringe_key_sequence_changed(self, name: str, sequence: QKeySequence) -> None:
        if not sequence.isEmpty():
            self.syringe_key_bindings[name] = int(sequence[0].toCombined())

    def _double_spinbox(self, value: float = 0.0, minimum: float = -100.0, maximum: float = 100.0) -> NumericLineEdit:
        return NumericLineEdit(value=value, minimum=minimum, maximum=maximum)

    def _spinbox(self, value: int, minimum: int, maximum: int) -> NumericLineEdit:
        return NumericLineEdit(value=value, minimum=minimum, maximum=maximum, integer=True)

    def _number_edit(self, value: float | int) -> QLineEdit:
        widget = QLineEdit()
        widget.setText(f"{value:g}")
        widget.setAlignment(Qt.AlignmentFlag.AlignRight)
        return widget

    def _line_float(self, widget: QLineEdit, default: float, minimum: float | None = None, maximum: float | None = None) -> float:
        try:
            value = float(widget.text().strip())
        except ValueError:
            value = default
        if minimum is not None:
            value = max(minimum, value)
        if maximum is not None:
            value = min(maximum, value)
        return value

    def _line_int(self, widget: QLineEdit, default: int, minimum: int | None = None, maximum: int | None = None) -> int:
        value = int(round(self._line_float(widget, float(default), minimum, maximum)))
        return value

    def _set_number_edit(self, widget: QLineEdit, value: float | int) -> None:
        if isinstance(value, int):
            widget.setText(str(value))
        else:
            widget.setText(f"{value:g}")

    def _config_root_dir(self) -> Path:
        return Path.home() / "Documents" / "Neurostar_Master" / "Configs"

    def _config_dir(self, kind: str) -> Path:
        mapping = {
            "injection": self._config_root_dir() / "Injection",
            "injection_sites": self._config_root_dir() / "Injection Sites",
            "craniotomy": self._config_root_dir() / "Craniotomy",
        }
        directory = mapping[kind]
        directory.mkdir(parents=True, exist_ok=True)
        return directory

    def _last_used_config_path(self, kind: str) -> Path:
        return self._config_dir(kind) / "last_used.json"

    def _general_settings_path(self) -> Path:
        return self._config_root_dir() / "settings.json"

    def _project_session_path(self) -> Path:
        return self._config_root_dir() / "project_session.json"

    def _has_recoverable_project_state(self) -> bool:
        return bool(self.seeds or self.trajectory or self.injection_sites or self.quick_locations or self.named_locations)

    def _project_session_dict(self) -> dict[str, object]:
        return {
            "format": "neurostar-project-session-v2",
            "craniotomy_coordinate_system": self.craniotomy_coordinate_system,
            "injection_sites_coordinate_system": self.injection_sites_coordinate_system,
            "coordinate_mode": self.coordinate_mode,
            "bregma_axis": self.bregma_axis,
            "anchor_axis": self.anchor_axis,
            "anchor_bregma": self.anchor_bregma,
            "craniotomy_config": {
                "diameter_mm": self._craniotomy_config().diameter_mm,
                "seed_count": self._craniotomy_config().seed_count,
                "trajectory_points": self._craniotomy_config().trajectory_points,
                "cut_offset_dv_mm": self._craniotomy_config().cut_offset_dv_mm,
                "max_depth_mm": self._craniotomy_config().max_depth_mm,
                "depth_per_round_mm": self._craniotomy_config().depth_per_round_mm,
                "skull_thickness_mm": self._craniotomy_config().skull_thickness_mm,
                "round_time_seconds": self._craniotomy_config().round_time_seconds,
                "drill_rate_mm_per_s": self._craniotomy_config().drill_rate_mm_per_s,
                "auto_start_rounds": self._craniotomy_config().auto_start_rounds,
            },
            "craniotomy_center": {"ap": self.mid_ap.value(), "ml": self.mid_ml.value()},
            "seeds": [
                {
                    "index": seed.index, "angle_deg": seed.angle_deg, "ap": seed.ap, "ml": seed.ml,
                    "dv": seed.dv, "sampled_ap": seed.sampled_ap, "sampled_ml": seed.sampled_ml,
                }
                for seed in self.seeds
            ],
            "trajectory": self.trajectory,
            "drilled_depths": self.drilled_depths,
            "frozen_points": self.frozen_points,
            "current_seed_index": self.current_seed_index,
            "current_target_depth_mm": self.current_target_depth_mm,
            "injection_config": self._injection_config_dict(),
            "injection_sites": [
                {"ap": site.ap, "ml": site.ml, "dv": site.dv, "generated": site.generated}
                for site in self.injection_sites
            ],
            "quick_locations": {
                name: {"ap": location.ap, "ml": location.ml, "dv": location.dv, "coordinate_system": location.coordinate_system}
                for name, location in self.quick_locations.items()
            },
            "named_locations": {
                name: {"ap": location.ap, "ml": location.ml, "dv": location.dv}
                for name, location in self.named_locations.items()
            },
            "overlay_name": Path(str(self.overlay_combo.currentData())).name if self.overlay_combo.currentData() else None,
            "top_zoom_index": self.zoom_mode_combo.currentIndex(),
            "injection_zoom_index": self.injection_sites_zoom_combo.currentIndex(),
            "active_tab": self.tabs.currentIndex(),
        }

    def _autosave_project_session(self) -> None:
        path = self._project_session_path()
        try:
            if not self._has_recoverable_project_state():
                if path.exists():
                    path.unlink()
                return
            path.parent.mkdir(parents=True, exist_ok=True)
            temporary_path = path.with_suffix(".tmp")
            temporary_path.write_text(json.dumps(self._project_session_dict(), indent=2), encoding="utf-8")
            temporary_path.replace(path)
        except Exception:
            # A session backup must never interrupt a live procedure or UI action.
            pass

    @staticmethod
    def _session_axis(value: object) -> tuple[float, float, float] | None:
        if not isinstance(value, list) or len(value) != 3:
            return None
        try:
            position = tuple(float(item) for item in value)
            return position if all(math.isfinite(item) for item in position) else None
        except (TypeError, ValueError):
            return None

    def _offer_project_session_restore(self) -> None:
        path = self._project_session_path()
        if not path.exists():
            return
        try:
            payload = self._read_config_file(path)
            if not isinstance(payload, dict) or not payload.get("format", "").startswith("neurostar-project-session-"):
                raise ValueError("not a project session")
        except (OSError, ValueError, TypeError, json.JSONDecodeError):
            return
        response = QMessageBox.question(
            self,
            "Restore Previous Project?",
            "Restore the previous craniotomy/injection project session?\n\n"
            "Active movements, drilling, and injections are never resumed automatically.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if response != QMessageBox.Yes:
            try:
                path.unlink()
            except OSError:
                pass
            return
        try:
            self._restore_project_session(payload)
            self.set_status("Restored the previous project session.")
        except Exception as exc:
            QMessageBox.warning(self, "Restore Previous Project", f"Could not restore the saved project. {exc}")

    def _restore_project_session(self, payload: dict[str, object]) -> None:
        config = payload.get("craniotomy_config")
        if isinstance(config, dict):
            self._apply_craniotomy_config(
                CraniotomyConfig(
                    diameter_mm=float(config.get("diameter_mm", self.diameter.value())),
                    seed_count=int(config.get("seed_count", self.seed_count.value())),
                    trajectory_points=int(config.get("trajectory_points", self.trajectory_points.value())),
                    cut_offset_dv_mm=float(config.get("cut_offset_dv_mm", self.cut_offset.value())),
                    max_depth_mm=float(config.get("max_depth_mm", self.drill_depth.value())),
                    depth_per_round_mm=float(config.get("depth_per_round_mm", self.depth_per_round.value())),
                    skull_thickness_mm=float(config.get("skull_thickness_mm", self.skull_thickness_mm.value())),
                    round_time_seconds=float(config.get("round_time_seconds", self.round_time_seconds.value())),
                    drill_rate_mm_per_s=float(config.get("drill_rate_mm_per_s", self.drill_rate_mm_per_s.value())),
                    auto_start_rounds=bool(config.get("auto_start_rounds", self.auto_start_rounds.isChecked())),
                )
            )
        injection_config = payload.get("injection_config")
        if isinstance(injection_config, dict):
            self._apply_injection_config_dict(injection_config)
        self.bregma_axis = self._session_axis(payload.get("bregma_axis"))
        self.craniotomy_coordinate_system = str(payload.get("craniotomy_coordinate_system", "unknown" if payload.get("seeds") or payload.get("trajectory") else "bregma"))
        self.injection_sites_coordinate_system = str(payload.get("injection_sites_coordinate_system", "unknown" if payload.get("injection_sites") else "bregma"))
        self.anchor_axis = self._session_axis(payload.get("anchor_axis"))
        self.anchor_bregma = self._session_axis(payload.get("anchor_bregma"))
        self.coordinate_mode = "bregma" if payload.get("coordinate_mode") == "bregma" and self.bregma_axis else "axis"
        center = payload.get("craniotomy_center")
        if isinstance(center, dict):
            self.mid_ap.setValue(float(center.get("ap", self.mid_ap.value())))
            self.mid_ml.setValue(float(center.get("ml", self.mid_ml.value())))
        self.seeds = []
        for raw_seed in payload.get("seeds", []):
            if not isinstance(raw_seed, dict):
                continue
            self.seeds.append(SeedPoint(
                index=int(raw_seed["index"]), angle_deg=float(raw_seed["angle_deg"]),
                ap=float(raw_seed["ap"]), ml=float(raw_seed["ml"]),
                dv=None if raw_seed.get("dv") is None else float(raw_seed["dv"]),
                sampled_ap=None if raw_seed.get("sampled_ap") is None else float(raw_seed["sampled_ap"]),
                sampled_ml=None if raw_seed.get("sampled_ml") is None else float(raw_seed["sampled_ml"]),
            ))
        self.trajectory = [tuple(float(value) for value in point) for point in payload.get("trajectory", []) if isinstance(point, list) and len(point) == 3]
        self.drilled_depths = [float(value) for value in payload.get("drilled_depths", [])]
        self.frozen_points = [bool(value) for value in payload.get("frozen_points", [])]
        self.current_seed_index = payload.get("current_seed_index") if isinstance(payload.get("current_seed_index"), int) else None
        if self.current_seed_index is not None and not 0 <= self.current_seed_index < len(self.seeds):
            self.current_seed_index = None
        self.current_target_depth_mm = float(payload.get("current_target_depth_mm", self._initial_target_depth()))
        self.injection_sites = []
        for raw_site in payload.get("injection_sites", []):
            if isinstance(raw_site, dict):
                self.injection_sites.append(InjectionSite(
                    ap=float(raw_site["ap"]), ml=float(raw_site["ml"]),
                    dv=None if raw_site.get("dv") is None else float(raw_site["dv"]),
                    generated=bool(raw_site.get("generated", False)),
                ))
        self.quick_locations = {}
        raw_locations = payload.get("quick_locations")
        if isinstance(raw_locations, dict):
            for name, raw_location in raw_locations.items():
                if isinstance(name, str) and isinstance(raw_location, dict):
                    self.quick_locations[name] = StoredLocation(
                        ap=float(raw_location["ap"]), ml=float(raw_location["ml"]), dv=float(raw_location["dv"]),
                        coordinate_system=str(raw_location.get("coordinate_system", "unknown")),
                    )
        self.named_locations = {}
        raw_named_locations = payload.get("named_locations")
        if isinstance(raw_named_locations, dict):
            for name, raw_location in raw_named_locations.items():
                if isinstance(name, str) and isinstance(raw_location, dict):
                    self.named_locations[name] = StoredLocation(
                        ap=float(raw_location["ap"]), ml=float(raw_location["ml"]), dv=float(raw_location["dv"]),
                    )
        overlay_name = payload.get("overlay_name")
        if isinstance(overlay_name, str):
            for index in range(self.overlay_combo.count()):
                candidate = self.overlay_combo.itemData(index)
                if candidate and Path(str(candidate)).name == overlay_name:
                    self.overlay_combo.setCurrentIndex(index)
                    break
        self.current_seed_spin.blockSignals(True)
        self.current_seed_spin.setRange(1, max(1, len(self.seeds)))
        self.current_seed_spin.setValue((self.current_seed_index or 0) + 1)
        self.current_seed_spin.blockSignals(False)
        self.update_coordinate_mode_buttons()
        self.top_view.set_coordinate_mode_bregma(self.coordinate_mode == "bregma")
        self.injection_sites_view.set_coordinate_mode_bregma(self.coordinate_mode == "bregma")
        self.update_seed_selector_label()
        self.update_current_target_depth_label()
        self.refresh_injection_sites_list()
        self.zoom_mode_combo.setCurrentIndex(max(0, min(2, int(payload.get("top_zoom_index", self.zoom_mode_combo.currentIndex())))))
        self.injection_sites_zoom_combo.setCurrentIndex(max(0, min(2, int(payload.get("injection_zoom_index", self.injection_sites_zoom_combo.currentIndex())))))
        self.tabs.setCurrentIndex(max(0, min(self.tabs.count() - 1, int(payload.get("active_tab", self.tabs.currentIndex())))))
        self.redraw_views()

        if self.craniotomy_coordinate_system != "bregma" or self.injection_sites_coordinate_system != "bregma":
            QMessageBox.information(self, "Previous Project Coordinates",
                "This older project did not record coordinate references. Its saved data is retained, "
                "but recreate the craniotomy and clear/load injection sites before using them for movement. "
                "Re-save A/B/C positions before using them. Named Bregma positions remain available.")

    def _save_general_settings(self) -> None:
        bindings = {}
        for name, edit in self.movement_key_edits.items():
            sequence = edit.keySequence()
            if not sequence.isEmpty():
                bindings[name] = int(sequence[0].toCombined())
        self._config_root_dir().mkdir(parents=True, exist_ok=True)
        payload = {
            "movement_keys": {name: bindings[name] for name in self.movement_key_bindings if name in bindings},
            "syringe_keys": {name: bindings[name] for name in self.syringe_key_bindings if name in bindings},
            "coordinate_mode": self.coordinate_mode,
            "bregma_axis": self.bregma_axis,
            "anchor_axis": self.anchor_axis,
            "anchor_bregma": self.anchor_bregma,
            "recent_injection_grid_configs": self.recent_injection_grid_configs,
            "window_geometry": getattr(self, "window_geometry", None),
        }
        self._write_config_file(self._general_settings_path(), payload)

    def _load_general_settings(self) -> None:
        path = self._general_settings_path()
        if not path.exists():
            return
        try:
            payload = self._read_config_file(path)
            bindings = payload.get("movement_keys", {})
            for name, edit in self.movement_key_edits.items():
                value = bindings.get(name)
                if isinstance(value, int) and value:
                    self.movement_key_bindings[name] = value
                    edit.setKeySequence(QKeySequence(value))
            syringe_bindings = payload.get("syringe_keys", {})
            for name in self.syringe_key_bindings:
                value = syringe_bindings.get(name)
                if isinstance(value, int) and value:
                    self.syringe_key_bindings[name] = value
                    self.movement_key_edits[name].setKeySequence(QKeySequence(value))
            mode = payload.get("coordinate_mode")
            if mode in {"axis", "bregma"}:
                self.coordinate_mode = mode
            for name in ("bregma_axis", "anchor_axis", "anchor_bregma"):
                value = payload.get(name)
                if isinstance(value, list) and len(value) == 3:
                    setattr(self, name, tuple(float(item) for item in value))
            self.window_geometry = payload.get("window_geometry")
            recent_grids = payload.get("recent_injection_grid_configs", [])
            if isinstance(recent_grids, list):
                self.recent_injection_grid_configs = [
                    item for item in recent_grids
                    if isinstance(item, dict)
                ][:8]
            if hasattr(self, "injection_grid_recent_combo"):
                self._refresh_injection_grid_recent_combo()
            self.update_coordinate_mode_buttons()
            self.top_view.set_coordinate_mode_bregma(self.coordinate_mode == "bregma")
            self.injection_sites_view.set_coordinate_mode_bregma(self.coordinate_mode == "bregma")
        except (OSError, ValueError, TypeError, json.JSONDecodeError):
            self.set_status(f"Could not read settings from {path}")

    def restore_saved_window_geometry(self) -> None:
        geometry = getattr(self, "window_geometry", None)
        if not isinstance(geometry, dict):
            return
        try:
            self.setGeometry(int(geometry["x"]), int(geometry["y"]), int(geometry["width"]), int(geometry["height"]))
        except (KeyError, TypeError, ValueError):
            return
    def reset_movement_key_bindings(self) -> None:
        for name, key in {
            "speed_decrease": Qt.Key.Key_Shift,
            "speed_increase": Qt.Key.Key_Ccedilla,
            "ml_left": Qt.Key.Key_Left,
            "ml_right": Qt.Key.Key_Right,
            "ap_anterior": Qt.Key.Key_Up,
            "ap_posterior": Qt.Key.Key_Down,
            "dv_up": Qt.Key.Key_PageUp,
            "dv_down": Qt.Key.Key_PageDown,
        }.items():
            self.movement_key_bindings[name] = int(key)
            self.movement_key_edits[name].setKeySequence(QKeySequence(key))
        for name, key in {
            "volume_down": Qt.Key.Key_F1,
            "volume_up": Qt.Key.Key_F2,
            "syringe_up": Qt.Key.Key_F3,
            "syringe_down": Qt.Key.Key_F4,
            "stop_injection": Qt.Key.Key_Escape,
        }.items():
            self.syringe_key_bindings[name] = int(key)
            self.movement_key_edits[name].setKeySequence(QKeySequence(key))
        self.set_status("Movement keys reset to defaults")

    def _craniotomy_config(self) -> CraniotomyConfig:
        return CraniotomyConfig(
            diameter_mm=float(self.diameter.value()),
            seed_count=int(self.seed_count.value()),
            trajectory_points=int(self.trajectory_points.value()),
            cut_offset_dv_mm=float(self.cut_offset.value()),
            max_depth_mm=float(self.drill_depth.value()),
            depth_per_round_mm=float(self.depth_per_round.value()),
            skull_thickness_mm=float(self.skull_thickness_mm.value()),
            round_time_seconds=float(self.round_time_seconds.value()),
            drill_rate_mm_per_s=float(self.drill_rate_mm_per_s.value()),
            auto_start_rounds=bool(self.auto_start_rounds.isChecked()),
        )

    def _apply_craniotomy_config(self, config: CraniotomyConfig) -> None:
        self.diameter.setValue(config.diameter_mm)
        self.seed_count.setValue(config.seed_count)
        self.trajectory_points.setValue(config.trajectory_points)
        self.cut_offset.setValue(config.cut_offset_dv_mm)
        self.drill_depth.setValue(config.max_depth_mm)
        self.depth_per_round.setValue(config.depth_per_round_mm)
        self.skull_thickness_mm.setValue(config.skull_thickness_mm)
        self.round_time_seconds.setValue(config.round_time_seconds)
        self.drill_rate_mm_per_s.setValue(config.drill_rate_mm_per_s)
        self.auto_start_rounds.setChecked(config.auto_start_rounds)
        self.depth_legend.set_skull_thickness_mm(self.skull_thickness_mm.value())
        self.update_current_target_depth_label()
        self.redraw_views()

    def _injection_config_dict(self) -> dict[str, object]:
        settings = self._injection_protocol_settings()
        return {
            "main_volume_nl": settings.main_volume_nl,
            "insertion_rate_nl_min": settings.insertion_rate_nl_min,
            "main_rate_nl_min": settings.main_rate_nl_min,
            "injection_depth_mm": settings.injection_depth_mm,
            "insert_retract_speed_um_s": settings.insert_retract_speed_um_s,
            "overshoot_mm": settings.overshoot_mm,
            "post_inject_pause_s": settings.post_inject_pause_s,
            "block_test_volume_nl": self._rounded_test_volume(),
            "block_check_enabled": bool(self.block_check.isChecked()),
        }

    def _apply_injection_config_dict(self, config: dict[str, object]) -> None:
        self._set_number_edit(self.single_injection_volume_nl, int(round(float(config.get("main_volume_nl", 100)))))
        self._set_number_edit(self.insertion_injection_rate_nl_min, float(config.get("insertion_rate_nl_min", 100.0)))
        self._set_number_edit(self.main_injection_rate_nl_min, float(config.get("main_rate_nl_min", 100.0)))
        self._set_number_edit(self.injection_depth_mm, float(config.get("injection_depth_mm", 0.2)))
        self._set_number_edit(self.insert_retract_speed_um_s, float(config.get("insert_retract_speed_um_s", 20.0)))
        self._set_number_edit(self.movement_overshoot_mm, float(config.get("overshoot_mm", 0.05)))
        self._set_number_edit(self.post_inject_pause_s, float(config.get("post_inject_pause_s", 5.0)))
        self._set_number_edit(self.block_test_volume_nl, int(round(float(config.get("block_test_volume_nl", 50)))))
        self.block_check.setChecked(bool(config.get("block_check_enabled", True)))
        self.round_single_injection_volume_up()
        self.round_test_volume_to_supported()
        self.update_injection_rate_label()
        self.refresh_injection_sequence_summary()

    def _write_config_file(self, path: Path, payload: dict[str, object]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def _read_config_file(self, path: Path) -> dict[str, object]:
        return json.loads(path.read_text(encoding="utf-8"))

    def _save_last_used_configs(self) -> None:
        self._write_config_file(
            self._last_used_config_path("injection"),
            self._injection_config_dict(),
        )
        self._write_config_file(
            self._last_used_config_path("craniotomy"),
            {
                "diameter_mm": self._craniotomy_config().diameter_mm,
                "seed_count": self._craniotomy_config().seed_count,
                "trajectory_points": self._craniotomy_config().trajectory_points,
                "cut_offset_dv_mm": self._craniotomy_config().cut_offset_dv_mm,
                "max_depth_mm": self._craniotomy_config().max_depth_mm,
                "depth_per_round_mm": self._craniotomy_config().depth_per_round_mm,
                "skull_thickness_mm": self._craniotomy_config().skull_thickness_mm,
                "round_time_seconds": self._craniotomy_config().round_time_seconds,
                "drill_rate_mm_per_s": self._craniotomy_config().drill_rate_mm_per_s,
                "auto_start_rounds": self._craniotomy_config().auto_start_rounds,
            },
        )

    def _load_last_used_configs(self) -> None:
        injection_path = self._last_used_config_path("injection")
        if injection_path.exists():
            try:
                self._apply_injection_config_dict(self._read_config_file(injection_path))
            except Exception:
                pass
        craniotomy_path = self._last_used_config_path("craniotomy")
        if craniotomy_path.exists():
            try:
                self._apply_craniotomy_config(
                    CraniotomyConfig(
                        diameter_mm=float(self._read_config_file(craniotomy_path).get("diameter_mm", 3.2)),
                        seed_count=int(self._read_config_file(craniotomy_path).get("seed_count", 6)),
                        trajectory_points=int(self._read_config_file(craniotomy_path).get("trajectory_points", 60)),
                        cut_offset_dv_mm=float(self._read_config_file(craniotomy_path).get("cut_offset_dv_mm", 0.0)),
                        max_depth_mm=float(self._read_config_file(craniotomy_path).get("max_depth_mm", 0.2)),
                        depth_per_round_mm=float(self._read_config_file(craniotomy_path).get("depth_per_round_mm", 0.05)),
                        skull_thickness_mm=float(self._read_config_file(craniotomy_path).get("skull_thickness_mm", 0.25)),
                        round_time_seconds=float(self._read_config_file(craniotomy_path).get("round_time_seconds", 60.0)),
                        drill_rate_mm_per_s=float(self._read_config_file(craniotomy_path).get("drill_rate_mm_per_s", 0.01)),
                        auto_start_rounds=bool(self._read_config_file(craniotomy_path).get("auto_start_rounds", True)),
                    )
                )
            except Exception:
                pass

    def save_injection_config(self) -> None:
        directory = self._config_dir("injection")
        path_str, _selected = QFileDialog.getSaveFileName(
            self,
            "Save Injection Config",
            str(directory / "injection_config.json"),
            "JSON Files (*.json)",
        )
        if not path_str:
            return
        path = Path(path_str)
        if path.suffix.lower() != ".json":
            path = path.with_suffix(".json")
        payload = self._injection_config_dict()
        self._write_config_file(path, payload)
        self._write_config_file(self._last_used_config_path("injection"), payload)
        self.set_status(f"Saved injection config to {path}")

    def load_injection_config(self) -> None:
        directory = self._config_dir("injection")
        path_str, _selected = QFileDialog.getOpenFileName(
            self,
            "Load Injection Config",
            str(directory),
            "JSON Files (*.json)",
        )
        if not path_str:
            return
        path = Path(path_str)
        payload = self._read_config_file(path)
        self._apply_injection_config_dict(payload)
        self._write_config_file(self._last_used_config_path("injection"), self._injection_config_dict())
        self.set_status(f"Loaded injection config from {path}")

    def save_craniotomy_config(self) -> None:
        directory = self._config_dir("craniotomy")
        path_str, _selected = QFileDialog.getSaveFileName(
            self,
            "Save Craniotomy Config",
            str(directory / "craniotomy_config.json"),
            "JSON Files (*.json)",
        )
        if not path_str:
            return
        path = Path(path_str)
        if path.suffix.lower() != ".json":
            path = path.with_suffix(".json")
        config = self._craniotomy_config()
        payload = {
            "diameter_mm": config.diameter_mm,
            "seed_count": config.seed_count,
            "trajectory_points": config.trajectory_points,
            "cut_offset_dv_mm": config.cut_offset_dv_mm,
            "max_depth_mm": config.max_depth_mm,
            "depth_per_round_mm": config.depth_per_round_mm,
            "skull_thickness_mm": config.skull_thickness_mm,
            "round_time_seconds": config.round_time_seconds,
            "drill_rate_mm_per_s": config.drill_rate_mm_per_s,
            "auto_start_rounds": config.auto_start_rounds,
        }
        self._write_config_file(path, payload)
        self._write_config_file(self._last_used_config_path("craniotomy"), payload)
        self.set_status(f"Saved craniotomy config to {path}")

    def load_craniotomy_config(self) -> None:
        if not self._require_idle("Load Craniotomy Settings"):
            return
        directory = self._config_dir("craniotomy")
        path_str, _selected = QFileDialog.getOpenFileName(
            self,
            "Load Craniotomy Config",
            str(directory),
            "JSON Files (*.json)",
        )
        if not path_str:
            return
        path = Path(path_str)
        payload = self._read_config_file(path)
        config = CraniotomyConfig(
            diameter_mm=float(payload.get("diameter_mm", 3.2)),
            seed_count=int(payload.get("seed_count", 6)),
            trajectory_points=int(payload.get("trajectory_points", 60)),
            cut_offset_dv_mm=float(payload.get("cut_offset_dv_mm", 0.0)),
            max_depth_mm=float(payload.get("max_depth_mm", 0.2)),
            depth_per_round_mm=float(payload.get("depth_per_round_mm", 0.05)),
            skull_thickness_mm=float(payload.get("skull_thickness_mm", 0.25)),
            round_time_seconds=float(payload.get("round_time_seconds", 60.0)),
            drill_rate_mm_per_s=float(payload.get("drill_rate_mm_per_s", 0.01)),
            auto_start_rounds=bool(payload.get("auto_start_rounds", True)),
        )
        self._apply_craniotomy_config(config)
        self._write_config_file(
            self._last_used_config_path("craniotomy"),
            {
                "diameter_mm": config.diameter_mm,
                "seed_count": config.seed_count,
                "trajectory_points": config.trajectory_points,
                "cut_offset_dv_mm": config.cut_offset_dv_mm,
                "max_depth_mm": config.max_depth_mm,
                "depth_per_round_mm": config.depth_per_round_mm,
                "skull_thickness_mm": config.skull_thickness_mm,
                "round_time_seconds": config.round_time_seconds,
                "drill_rate_mm_per_s": config.drill_rate_mm_per_s,
                "auto_start_rounds": config.auto_start_rounds,
            },
        )
        self.set_status(f"Loaded craniotomy config from {path}")

    def set_status(self, message: str) -> None:
        self.current_action = message or "Trajectory"
        if hasattr(self, "action_status_label"):
            self.action_status_label.setText(self.current_action)

    def update_move_speed_label(self) -> None:
        try:
            speed_index = MOVE_SPEED_OPTIONS_MM.index(self.move_speed_step_mm)
        except ValueError:
            speed_index = min(
                range(len(MOVE_SPEED_OPTIONS_MM)),
                key=lambda index: abs(MOVE_SPEED_OPTIONS_MM[index] - self.move_speed_step_mm),
            )
        ratio = speed_index / max(1, len(MOVE_SPEED_OPTIONS_MM) - 1)
        red = int(37 + ratio * 183)
        green = int(99 - ratio * 61)
        blue = int(235 - ratio * 197)
        color = f"#{red:02x}{green:02x}{blue:02x}"
        self.move_speed_label.setText(f"Move speed = {self.move_speed_step_mm:g} mm")
        self.move_speed_label.setStyleSheet(f"color: {color}; font-weight: 700;")

    def eventFilter(self, watched, event) -> bool:  # noqa: N802
        if event.type() != QEvent.Type.KeyPress:
            return super().eventFilter(watched, event)
        key = event.key()
        if self.validation_move_active and key == Qt.Key.Key_Escape:
            if self.validation_move_cancel_callback is not None:
                self.validation_move_cancel_callback()
            return True
        if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter) and self._focus_is_editable():
            focus_widget = QApplication.focusWidget()
            if focus_widget is not None:
                focus_widget.clearFocus()
            self.setFocus(Qt.OtherFocusReason)
            return True
        if self._focus_is_editable():
            return super().eventFilter(watched, event)
        combined_key = int(key) | event.modifiers().value
        validation_active = self.validation_modal_active
        if not validation_active and self._key_matches_binding(self.syringe_key_bindings["volume_down"], key, combined_key):
            self.adjust_manual_injection_volume(-1)
            return True
        if not validation_active and self._key_matches_binding(self.syringe_key_bindings["volume_up"], key, combined_key):
            self.adjust_manual_injection_volume(1)
            return True
        if not validation_active and self._key_matches_binding(self.syringe_key_bindings["syringe_up"], key, combined_key):
            self.manual_syringe_step(up=True)
            return True
        if not validation_active and self._key_matches_binding(self.syringe_key_bindings["syringe_down"], key, combined_key):
            self.manual_syringe_step(up=False)
            return True
        if not validation_active and self._key_matches_binding(self.syringe_key_bindings["stop_injection"], key, combined_key):
            self.stop_injection()
            return True
        if self._key_matches_binding(self.movement_key_bindings["speed_decrease"], key, combined_key):
            self.adjust_move_speed(-1)
            return True
        if self._key_matches_binding(self.movement_key_bindings["speed_increase"], key, combined_key):
            self.adjust_move_speed(1)
            return True
        for binding, axis, positive, label in (
            (self.movement_key_bindings["ml_left"], "ML", False, "ML left"),
            (self.movement_key_bindings["ml_right"], "ML", True, "ML right"),
            (self.movement_key_bindings["ap_anterior"], "AP", True, "AP anterior"),
            (self.movement_key_bindings["ap_posterior"], "AP", False, "AP posterior"),
            (self.movement_key_bindings["dv_up"], "DV", False, "DV up"),
            (self.movement_key_bindings["dv_down"], "DV", True, "DV down"),
        ):
            if self._key_matches_binding(binding, key, combined_key):
                if self.nudge_all_sites_active and axis in {"AP", "ML"}:
                    self.nudge_all_injection_sites(axis, positive, label)
                    return True
                self.keyboard_nudge(axis, positive, label)
                return True
        return super().eventFilter(watched, event)

    @staticmethod
    def _key_matches_binding(binding: int, key: int, combined_key: int) -> bool:
        """Match normal shortcuts and modifier-only keys such as Shift."""
        return combined_key == binding or int(key) == binding

    def _focus_is_editable(self) -> bool:
        focus_widget = QApplication.focusWidget()
        return isinstance(focus_widget, (QLineEdit, QComboBox, QPlainTextEdit, QKeySequenceEdit))

    def adjust_move_speed(self, direction: int) -> None:
        current_index = min(
            range(len(MOVE_SPEED_OPTIONS_MM)),
            key=lambda index: abs(MOVE_SPEED_OPTIONS_MM[index] - self.move_speed_step_mm),
        )
        next_index = max(0, min(len(MOVE_SPEED_OPTIONS_MM) - 1, current_index + direction))
        self.move_speed_step_mm = MOVE_SPEED_OPTIONS_MM[next_index]
        self.update_move_speed_label()
        self.set_status(f"Move speed set to {self.move_speed_step_mm:g} mm")

    def keyboard_nudge(self, axis: str, positive: bool, label: str) -> None:
        if self._motion_is_active() and not self.validation_modal_active:
            return
        if self._focus_is_editable():
            return
        try:
            self.controller.prepare_motion()
            self.update_move_speed_label()
            self.controller.set_nudge_step(axis, self.move_speed_step_mm)
            self.controller.nudge_axis(axis, positive)
            self.set_status(f"Keyboard nudge: {label} {self.move_speed_step_mm:g} mm")
            try:
                self.refresh_live_position()
            except Exception:
                pass
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def set_nudge_all_sites_active(self, active: bool) -> None:
        """Toggle AP/ML shortcuts between manipulator movement and site translation."""
        if active and not self.injection_sites:
            QMessageBox.information(self, "Nudge All Sites", "Add or load injection sites before enabling this mode.")
            self.nudge_all_sites_btn.blockSignals(True)
            self.nudge_all_sites_btn.setChecked(False)
            self.nudge_all_sites_btn.blockSignals(False)
            return
        if active and self._motion_is_active():
            QMessageBox.warning(self, "Nudge All Sites", "Wait for the current operation to finish first.")
            self.nudge_all_sites_btn.blockSignals(True)
            self.nudge_all_sites_btn.setChecked(False)
            self.nudge_all_sites_btn.blockSignals(False)
            return
        self.nudge_all_sites_active = active
        self.nudge_all_sites_btn.setProperty("variant", "quick-green" if active else None)
        self.nudge_all_sites_btn.style().unpolish(self.nudge_all_sites_btn)
        self.nudge_all_sites_btn.style().polish(self.nudge_all_sites_btn)
        if active:
            self.set_status(
                "Nudge All Sites enabled: AP/ML arrows shift all sites; the manipulator will not move."
            )
        else:
            self.set_status("Nudge All Sites disabled: AP/ML arrows move the manipulator.")

    def nudge_all_injection_sites(self, axis: str, positive: bool, label: str) -> None:
        if self._motion_is_active():
            return
        if not self.injection_sites:
            self.set_status("No injection sites to nudge.")
            return
        shift_mm = self.move_speed_step_mm if positive else -self.move_speed_step_mm
        for site in self.injection_sites:
            if axis == "AP":
                site.ap += shift_mm
            elif axis == "ML":
                site.ml += shift_mm
            # A translated site no longer has a confirmed surface coordinate.
            site.dv = None
            site.generated = True
        self.refresh_injection_sites_list()
        self.set_status(
            f"Nudged all {len(self.injection_sites)} injection sites {label} by "
            f"{self.move_speed_step_mm:g} mm; all now require validation."
        )

    def update_manual_volume_label(self) -> None:
        try:
            volume_index = INJECTION_VOLUME_OPTIONS_NL.index(self.manual_injection_volume_nl)
        except ValueError:
            volume_index = 0
        ratio = volume_index / max(1, len(INJECTION_VOLUME_OPTIONS_NL) - 1)
        red = int(37 + ratio * 183)
        green = int(99 - ratio * 61)
        blue = int(235 - ratio * 197)
        color = f"#{red:02x}{green:02x}{blue:02x}"
        self.manual_volume_label.setText(f"{self.manual_injection_volume_nl} nl")
        self.manual_volume_label.setStyleSheet(f"color: {color}; font-weight: 700;")
        if self.manual_volume_combo.currentData() != self.manual_injection_volume_nl:
            index = self.manual_volume_combo.findData(self.manual_injection_volume_nl)
            if index >= 0:
                self.manual_volume_combo.blockSignals(True)
                self.manual_volume_combo.setCurrentIndex(index)
                self.manual_volume_combo.blockSignals(False)

    def update_injection_rate_label(self) -> None:
        volume_nl = self._rounded_single_injection_volume()
        settings = self._injection_protocol_settings()
        duration_s = self._main_injection_duration_s(settings)
        self.injection_rate_label.setText(
            f"Main volume = {volume_nl} nl; estimated delivery duration = {duration_s:.1f} s"
        )
        self.refresh_injection_sequence_summary()

    def round_single_injection_volume_up(self) -> None:
        self._set_number_edit(self.single_injection_volume_nl, self._rounded_single_injection_volume())

    def round_test_volume_to_supported(self) -> None:
        self._set_number_edit(
            self.block_test_volume_nl,
            self._nearest_supported_injection_volume(self._line_int(self.block_test_volume_nl, 50, 10, 2000)),
        )

    def on_manual_volume_combo_changed(self) -> None:
        value = self.manual_volume_combo.currentData()
        if value is not None:
            self.manual_injection_volume_nl = int(value)
            self.update_manual_volume_label()

    def adjust_manual_injection_volume(self, direction: int) -> None:
        current_index = INJECTION_VOLUME_OPTIONS_NL.index(self.manual_injection_volume_nl)
        next_index = max(0, min(len(INJECTION_VOLUME_OPTIONS_NL) - 1, current_index + direction))
        self.manual_injection_volume_nl = INJECTION_VOLUME_OPTIONS_NL[next_index]
        self.update_manual_volume_label()
        self.set_status(f"Manual injection volume set to {self.manual_injection_volume_nl} nl")

    def manual_syringe_step(self, up: bool) -> None:
        if not self._require_idle("Syringe"):
            return
        try:
            self.ensure_syringe_move_allowed(self.manual_injection_volume_nl, up)
            self.controller.syringe_step(f"{self.manual_injection_volume_nl} nl", up=up)
            self.adjust_tracked_syringe_position(self.manual_injection_volume_nl if up else -self.manual_injection_volume_nl)
            direction = "up" if up else "down"
            self.set_status(f"Syringe step {direction}: {self.manual_injection_volume_nl} nl")
        except Exception as exc:
            QMessageBox.warning(self, "Injectomate", str(exc))

    def set_syringe_position(self, value_nl: object) -> None:
        with self.syringe_position_lock:
            if value_nl is None:
                self.syringe_position_nl = None
            else:
                self.syringe_position_nl = max(0.0, float(value_nl))
            position_nl = self.syringe_position_nl
        if not hasattr(self, "syringe_position_label") or not hasattr(self, "plunger_gauge"):
            return
        if position_nl is None:
            self.syringe_position_label.setText("Syringe position = -- nl")
            self.plunger_gauge.set_position(None)
        else:
            self.syringe_position_label.setText(f"Syringe position = {position_nl:.1f} nl")
            self.plunger_gauge.set_position(position_nl)

    def adjust_tracked_syringe_position(self, delta_nl: float) -> None:
        with self.syringe_position_lock:
            if self.syringe_position_nl is None:
                return
            self.syringe_position_nl = max(SYRINGE_MIN_NL, min(SYRINGE_MAX_NL, self.syringe_position_nl + delta_nl))
            position_nl = self.syringe_position_nl
        self.syringe_position_signal.emit(position_nl)

    def current_syringe_position(self) -> float | None:
        with self.syringe_position_lock:
            return self.syringe_position_nl

    def syringe_limit_message(self, requested_nl: float, up: bool, position_nl: float) -> str | None:
        max_possible_nl = (SYRINGE_MAX_NL - position_nl) if up else (position_nl - SYRINGE_MIN_NL)
        if requested_nl <= max_possible_nl + 1e-6:
            return None
        direction = "up" if up else "down"
        limit = SYRINGE_MAX_NL if up else SYRINGE_MIN_NL
        return (
            f"Requested syringe step {direction} of {requested_nl:.0f} nl would exceed the "
            f"{limit:.0f} nl limit from current position {position_nl:.1f} nl.\n\n"
            f"Maximum movement {direction}: {max(0.0, max_possible_nl):.1f} nl."
        )

    def ensure_syringe_move_allowed(self, requested_nl: float, up: bool) -> None:
        position_nl = self.current_syringe_position()
        if position_nl is None:
            self.sync_syringe_position_before_injection()
            position_nl = self.current_syringe_position()
        if position_nl is None:
            raise StereoDriveError("Syringe position is unknown. Click Update Syringe Position and try again.")
        message = self.syringe_limit_message(requested_nl, up, position_nl)
        if message:
            raise StereoDriveError(message)

    def ensure_total_syringe_capacity(self, requested_total_nl: float) -> None:
        position_nl = self.current_syringe_position()
        if position_nl is None:
            self.sync_syringe_position_before_injection()
            position_nl = self.current_syringe_position()
        if position_nl is None:
            raise StereoDriveError("Syringe position is unknown. Click Update Syringe Position and try again.")
        remaining_capacity_nl = position_nl - SYRINGE_MIN_NL
        if requested_total_nl > remaining_capacity_nl + 1e-6:
            raise StereoDriveError(
                f"Requested injection sequence volume is {requested_total_nl:.0f} nl, "
                f"but current syringe position is {position_nl:.1f} nl and the 0 nl limit leaves "
                f"only {max(0.0, remaining_capacity_nl):.1f} nl available.\n\n"
                f"Maximum sequence volume possible: {max(0.0, remaining_capacity_nl):.1f} nl."
            )

    def show_syringe_limit_warning(self, message: str) -> None:
        QMessageBox.warning(self, "Syringe Limit", message)

    def update_syringe_position_from_scale(self) -> None:
        if self._motion_is_active():
            return
        self._start_syringe_position_scale_read()

    def _start_syringe_position_scale_read(self, wait_for_injection_thread: bool = False) -> None:
        def worker() -> None:
            if wait_for_injection_thread:
                deadline = time.monotonic() + 15.0
                while (
                    self.injection_thread is not None
                    and self.injection_thread.is_alive()
                    and time.monotonic() < deadline
                ):
                    time.sleep(0.05)
            try:
                value_nl = self.controller.read_injectomate_calibrate_scale_nl()
                self.syringe_position_signal.emit(value_nl)
                self.status_signal.emit(f"Syringe position updated from scale: {value_nl:.3f} nl")
            except Exception as exc:
                self.status_signal.emit(f"Syringe position update failed: {exc}")

        threading.Thread(target=worker, daemon=True).start()

    def read_injectomate_scale(self) -> None:
        self.update_syringe_position_from_scale()

    def sync_syringe_position_before_injection(self) -> None:
        value_nl = self.controller.read_injectomate_calibrate_scale_nl()
        self.set_syringe_position(value_nl)
        self.set_status(f"Syringe position checked: {value_nl:.3f} nl")

    def track_injection_delivery(self, volume_nl: int) -> None:
        self.adjust_tracked_syringe_position(-volume_nl)

    def track_syringe_empty(self) -> None:
        self.set_syringe_position(0.0)

    def test_for_blockage(self) -> None:
        if not self._require_idle("Test Volume"):
            return
        try:
            volume_nl = self._nearest_supported_injection_volume(self._line_int(self.block_test_volume_nl, 50, 10, 2000))
            self._set_number_edit(self.block_test_volume_nl, volume_nl)
            self.ensure_syringe_move_allowed(volume_nl, False)
            self.controller.syringe_step(f"{volume_nl} nl", up=False)
            self.track_injection_delivery(volume_nl)
            self.set_status(f"Verifying no blockage (test volume = {volume_nl} nl)")
        except Exception as exc:
            QMessageBox.warning(self, "Injectomate", str(exc))

    def empty_syringe(self) -> None:
        if not self._require_idle("Empty Syringe"):
            return
        try:
            self.controller.empty_syringe()
            self.track_syringe_empty()
            self.set_status("Emptying syringe to 0")
        except Exception as exc:
            QMessageBox.critical(self, "Injectomate", str(exc))

    def update_coordinate_mode_buttons(self) -> None:
        active = "background: #108a54; color: white; font-weight: 700;"
        inactive = ""
        self.axis_mode_btn.setStyleSheet(active if self.coordinate_mode == "axis" else inactive)
        self.bregma_mode_btn.setStyleSheet(active if self.coordinate_mode == "bregma" else inactive)

    def set_coordinate_mode(self, mode: str) -> None:
        if not self._require_idle("Coordinate Mode"):
            return
        if mode == "bregma" and self.bregma_axis is None:
            QMessageBox.information(self, "Bregma Coordinates", "Set Bregma in this GUI before using Bregma coordinates.")
            return
        self.coordinate_mode = mode
        self.update_coordinate_mode_buttons()
        self.top_view.set_coordinate_mode_bregma(mode == "bregma")
        self.injection_sites_view.set_coordinate_mode_bregma(mode == "bregma")
        self.refresh_live_position()
        self.redraw_views()
        self.set_status(f"Using {mode.title()} coordinates.")

    def _axis_to_bregma(self, axis_position: tuple[float, float, float]) -> tuple[float, float, float]:
        if self.bregma_axis is None:
            raise StereoDriveError("GUI Bregma has not been set.")
        return tuple(axis_position[index] - self.bregma_axis[index] for index in range(3))

    def _bregma_to_axis(self, bregma_position: tuple[float, float, float]) -> tuple[float, float, float]:
        if self.bregma_axis is None:
            raise StereoDriveError("GUI Bregma has not been set.")
        position = tuple(self.bregma_axis[index] + bregma_position[index] for index in range(3))
        if not all(math.isfinite(value) for value in position):
            raise StereoDriveError("Movement coordinates must be finite numbers.")
        return position

    def get_gui_position(self) -> tuple[float, float, float]:
        axis_position = self.controller.get_current_axis_position()
        if self.coordinate_mode == "bregma" and self.bregma_axis is not None:
            return self._axis_to_bregma(axis_position)
        return axis_position

    def goto_gui_position(
        self,
        ap: float,
        ml: float,
        dv: float,
        delay_seconds: float = 0.75,
        *,
        title: str = "Moving",
        message: str = "Moving to the requested position. Waiting for StereoDrive to report arrival…",
        show_progress: bool = False,
    ) -> bool:
        axis_position = self._bregma_to_axis((ap, ml, dv)) if self.coordinate_mode == "bregma" else (ap, ml, dv)
        if show_progress:
            if not self._move_to_axis_position_with_progress(axis_position, title=title, message=message, delay_seconds=delay_seconds):
                return False
        else:
            self.controller.goto_axis_position(*axis_position, delay_seconds=delay_seconds)
        return True

    def wait_for_gui_position(self, ap: float, ml: float, dv: float, **kwargs) -> None:
        axis_position = self._bregma_to_axis((ap, ml, dv)) if self.coordinate_mode == "bregma" else (ap, ml, dv)
        self.controller.wait_for_axis_position(*axis_position, **kwargs)

    def set_local_bregma(self) -> None:
        if not self._require_idle("Set Bregma"):
            return
        try:
            # GUI Bregma is deliberately independent of StereoDrive's native
            # Bregma reference. Preserve the mechanical Axis values and use
            # the current position only as this application's local origin.
            self.bregma_axis = self.controller.get_current_axis_position()
            self.anchor_axis = None
            self.anchor_bregma = None
            self.coordinate_mode = "bregma"
            self.update_coordinate_mode_buttons()
            self._save_general_settings()
            self.refresh_live_position()
            self.redraw_views()
            self.set_status("GUI Bregma set at the current Axis position; any previous anchor was cleared.")
        except Exception as exc:
            QMessageBox.critical(self, "Bregma Coordinates", str(exc))

    def set_anchor(self) -> None:
        if not self._require_idle("Set Anchor"):
            return
        try:
            if self.bregma_axis is None:
                raise StereoDriveError("Set GUI Bregma before setting an anchor.")
            self.anchor_axis = self.controller.get_current_axis_position()
            self.anchor_bregma = self._axis_to_bregma(self.anchor_axis)
            self._save_general_settings()
            self.redraw_views()
            self.set_status(
                f"Anchor set at Bregma [{self.anchor_bregma[0]:.2f}, {self.anchor_bregma[1]:.2f}, {self.anchor_bregma[2]:.2f}]."
            )
        except Exception as exc:
            QMessageBox.critical(self, "Anchor", str(exc))

    def at_anchor(self) -> None:
        if not self._require_idle("At Anchor"):
            return
        try:
            if self.bregma_axis is None or self.anchor_bregma is None:
                raise StereoDriveError("Set Bregma and Set Anchor before using At Anchor.")
            response = QMessageBox.warning(
                self,
                "Recalibrate Bregma from Anchor?",
                "This recalibrates the GUI Bregma origin using the current position and the stored anchor. "
                "The craniotomy, injection-site, and position displays will update to the recalibrated Bregma coordinates.",
                QMessageBox.Yes | QMessageBox.Cancel,
                QMessageBox.Cancel,
            )
            if response != QMessageBox.Yes:
                return
            current_anchor_axis = self.controller.get_current_axis_position()
            self.bregma_axis = tuple(
                current_anchor_axis[index] - self.anchor_bregma[index] for index in range(3)
            )
            self.anchor_axis = current_anchor_axis
            self.coordinate_mode = "bregma"
            self.update_coordinate_mode_buttons()
            self.top_view.set_coordinate_mode_bregma(True)
            self.injection_sites_view.set_coordinate_mode_bregma(True)
            self._save_general_settings()
            current_ap, current_ml, current_dv = self._axis_to_bregma(current_anchor_axis)
            self.current_ap_label.setText(f"{current_ap:.2f}")
            self.current_ml_label.setText(f"{current_ml:.2f}")
            self.current_dv_label.setText(f"{current_dv:.2f}")
            self.refresh_injection_sites_list()
            self.redraw_views(current_point=(current_ml, current_ap))
            self.set_status("Bregma recalibrated from the current anchor position.")
        except Exception as exc:
            QMessageBox.critical(self, "At Anchor", str(exc))

    def set_current_location_to_bregma(self) -> None:
        self.set_local_bregma()

    def activate_stereodrive_drill(self) -> None:
        answer = QMessageBox.warning(
            self,
            "Activate Drill",
            "Open StereoDrive's Drill panel and toggle the drill control?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        try:
            self.controller.activate_drill_toggle()
            self.set_status("StereoDrive drill toggle clicked.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive Drill", str(exc))

    def goto_home(self) -> None:
        try:
            if self._run_named_motion_with_progress(
                self.controller.goto_home,
                title="Moving to Home",
                message="Moving to StereoDrive Home. Waiting for StereoDrive to report arrival…",
            ):
                self.set_status("Home command movement has settled. Verify the native Home position in StereoDrive.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def goto_work(self) -> None:
        try:
            if self._run_named_motion_with_progress(
                self.controller.goto_work,
                title="Moving to Work",
                message="Moving to StereoDrive Work. Waiting for StereoDrive to report arrival…",
            ):
                self.set_status("Work command movement has settled. Verify the native Work position in StereoDrive.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def goto_bregma(self) -> None:
        try:
            if not self._move_to_axis_position_with_progress(
                self._bregma_to_axis((0.0, 0.0, 0.0)),
                title="Moving to Bregma",
                message="Moving to Bregma. Waiting for StereoDrive to report arrival…",
            ):
                return
            self.set_status("Reached GUI Bregma: AP 0.00, ML 0.00, DV 0.00.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def open_goto_dialog(self) -> None:
        if not self._require_idle("Go to Position"):
            return
        try:
            axis_position = self.controller.get_current_axis_position()
            using_bregma = self.bregma_axis is not None
            current_ap, current_ml, current_dv = (
                self._axis_to_bregma(axis_position) if using_bregma else axis_position
            )
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Go to Position")
        layout = QGridLayout(dialog)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setHorizontalSpacing(8)
        layout.setVerticalSpacing(6)

        ap_box = self._position_spinbox(current_ap)
        ml_box = self._position_spinbox(current_ml)
        dv_box = self._position_spinbox(current_dv)

        saved_combo = QComboBox()
        name_edit = QLineEdit()
        save_location_btn = QPushButton("Save Position")
        delete_location_btn = QPushButton("Delete Saved Position")

        def populate_saved_positions(selected_name: str | None = None) -> None:
            saved_combo.blockSignals(True)
            saved_combo.clear()
            saved_combo.addItem("Saved positions…", None)
            for name in sorted(self.named_locations, key=str.casefold):
                saved_combo.addItem(name, name)
            if selected_name is not None:
                selected_index = saved_combo.findData(selected_name)
                if selected_index >= 0:
                    saved_combo.setCurrentIndex(selected_index)
            saved_combo.blockSignals(False)

        def load_saved_position(index: int) -> None:
            name = saved_combo.itemData(index)
            if not isinstance(name, str):
                return
            location = self.named_locations.get(name)
            if location is None:
                return
            name_edit.setText(name)
            ap_box.setValue(location.ap)
            ml_box.setValue(location.ml)
            dv_box.setValue(location.dv)

        def save_named_position() -> None:
            if not using_bregma:
                QMessageBox.information(
                    dialog,
                    "Save Position",
                    "Set Bregma before saving named positions. Named positions are always stored in Bregma coordinates.",
                )
                return
            name = name_edit.text().strip()
            if not name:
                QMessageBox.information(dialog, "Save Position", "Enter a name for this position.")
                return
            if name in self.named_locations:
                response = QMessageBox.question(
                    dialog, "Replace Saved Position?",
                    f"Replace the saved Bregma position '{name}'?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
                )
                if response != QMessageBox.Yes:
                    return
            self.named_locations[name] = StoredLocation(ap_box.value(), ml_box.value(), dv_box.value())
            populate_saved_positions(name)
            self._autosave_project_session()
            self.set_status(f"Saved Bregma position '{name}'.")

        def delete_named_position() -> None:
            name = saved_combo.currentData()
            if not isinstance(name, str) or name not in self.named_locations:
                QMessageBox.information(dialog, "Delete Saved Position", "Select a saved position to delete.")
                return
            del self.named_locations[name]
            name_edit.clear()
            populate_saved_positions()
            self._autosave_project_session()
            self.set_status(f"Deleted saved Bregma position '{name}'.")

        populate_saved_positions()
        saved_combo.currentIndexChanged.connect(load_saved_position)
        save_location_btn.clicked.connect(save_named_position)
        delete_location_btn.clicked.connect(delete_named_position)

        coordinate_label = "Bregma" if using_bregma else "Axis (set Bregma to save named positions)"
        layout.addWidget(QLabel("Saved position"), 0, 0)
        layout.addWidget(saved_combo, 0, 1)
        layout.addWidget(QLabel("Name"), 1, 0)
        layout.addWidget(name_edit, 1, 1)
        location_buttons = QHBoxLayout()
        location_buttons.addWidget(save_location_btn)
        location_buttons.addWidget(delete_location_btn)
        layout.addLayout(location_buttons, 2, 0, 1, 2)
        layout.addWidget(QLabel(f"{coordinate_label} AP"), 3, 0)
        layout.addWidget(ap_box, 3, 1)
        layout.addWidget(QLabel(f"{coordinate_label} ML"), 4, 0)
        layout.addWidget(ml_box, 4, 1)
        layout.addWidget(QLabel(f"{coordinate_label} DV"), 5, 0)
        layout.addWidget(dv_box, 5, 1)
        if not using_bregma:
            save_location_btn.setEnabled(False)
            delete_location_btn.setEnabled(False)
            saved_combo.setEnabled(False)
            name_edit.setEnabled(False)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons, 6, 0, 1, 2)

        if dialog.exec() != QDialog.Accepted:
            return

        ap = ap_box.value()
        ml = ml_box.value()
        dv = dv_box.value()
        try:
            if using_bregma:
                target_axis = self._bregma_to_axis((ap, ml, dv))
                if not self._move_to_axis_position_with_progress(
                    target_axis,
                    title="Moving to Position",
                    message="Moving to the requested Bregma position. Waiting for StereoDrive to report arrival…",
                ):
                    return
                self.set_status(f"Moved to Bregma AP {ap:.2f}, ML {ml:.2f}, DV {dv:.2f}.")
            else:
                if not self.goto_gui_position(
                    ap, ml, dv,
                    title="Moving to Position",
                    message="Moving to the requested position. Waiting for StereoDrive to report arrival…",
                    show_progress=True,
                ):
                    return
                self.set_status(f"Moved to Axis AP {ap:.2f}, ML {ml:.2f}, DV {dv:.2f}.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def _position_spinbox(self, value: float) -> NumericLineEdit:
        return NumericLineEdit(value=value, minimum=-100.0, maximum=100.0)

    def start_axis_benchmark(self) -> None:
        if not self._require_idle("Benchmark"):
            return
        if self.benchmark_thread is not None and self.benchmark_thread.is_alive():
            QMessageBox.information(self, "Benchmark", "Benchmark is already running.")
            return
        reply = QMessageBox.question(
            self,
            "Benchmark Axis Moves",
            "This will move AP, ML, and DV out-and-back at several distances. Continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        self.set_status("Running axis movement benchmark...")
        self.controller.prepare_motion()
        self.benchmark_thread = threading.Thread(target=self._run_axis_benchmark, daemon=True)
        self.benchmark_thread.start()

    def _run_axis_benchmark(self) -> None:
        try:
            rows = self.controller.benchmark_axis_moves(
                axes=["AP", "ML", "DV"],
                distances_mm=[0.02, 0.05, 0.1, 0.2, 0.5, 1.0],
                repeats=3,
                tolerance=0.003,
            )
            lines = [
                "axis,distance_mm,direction,repeat,start,target,end,achieved_mm,elapsed_s,mm_per_s,error_mm"
            ]
            for row in rows:
                mm_per_s = row["mm_per_s"]
                lines.append(
                    ",".join(
                        [
                            str(row["axis"]),
                            f"{float(row['distance_mm']):.3f}",
                            str(row["direction"]),
                            str(row["repeat"]),
                            f"{float(row['start']):.4f}",
                            f"{float(row['target']):.4f}",
                            f"{float(row['end']):.4f}",
                            f"{float(row['achieved_mm']):.4f}",
                            f"{float(row['elapsed_s']):.4f}",
                            "" if mm_per_s is None else f"{float(mm_per_s):.5f}",
                            f"{float(row['error_mm']):.4f}",
                        ]
                    )
                )
            self.benchmark_finished_signal.emit("\n".join(lines))
        except Exception as exc:
            try:
                self.controller.stop()
                self.controller.wait_until_stopped()
            except Exception as stop_exc:
                self.status_signal.emit(f"Stop confirmation failed: {stop_exc}")
            self.benchmark_finished_signal.emit(f"Benchmark failed:\n{exc}")
        finally:
            self.benchmark_thread = None

    def show_benchmark_results(self, text: str) -> None:
        self.set_status("Axis movement benchmark complete.")
        dialog = QDialog(self)
        dialog.setWindowTitle("Axis Movement Benchmark")
        layout = QVBoxLayout(dialog)
        result_box = QPlainTextEdit()
        result_box.setPlainText(text)
        result_box.setReadOnly(True)
        result_box.setMinimumWidth(900)
        result_box.setMinimumHeight(500)
        layout.addWidget(result_box)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def set_quick_location(self, slot: str) -> None:
        if not self._require_idle("Store Position"):
            return
        try:
            ap, ml, dv = self.get_bregma_position()
            self.quick_locations[slot] = StoredLocation(ap=ap, ml=ml, dv=dv)
            self.set_status(f"Stored location {slot}: AP {ap:.2f}, ML {ml:.2f}, DV {dv:.2f}.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def goto_quick_location(self, slot: str) -> None:
        location = self.quick_locations.get(slot)
        if location is None:
            QMessageBox.information(self, "Stored Location", f"Location {slot} has not been set.")
            return
        try:
            if location.coordinate_system != "bregma":
                raise StereoDriveError("Re-save this older position: its coordinate reference was not recorded.")
            if not self._move_to_axis_position_with_progress(
                self._bregma_to_axis((location.ap, location.ml, location.dv)),
                delay_seconds=0.5,
                title=f"Moving to Location {slot}",
                message=f"Moving to stored location {slot}. Waiting for StereoDrive to report arrival…",
            ):
                return
            self.set_status(
                f"Reached location {slot}: AP {location.ap:.2f}, ML {location.ml:.2f}, DV {location.dv:.2f}."
            )
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def _rounded_single_injection_volume(self) -> int:
        value = self._line_int(self.single_injection_volume_nl, 100, 10, 100000)
        return max(10, int(math.ceil(value / 10.0) * 10))

    def _nearest_supported_injection_volume(self, volume_nl: int) -> int:
        return min(INJECTION_VOLUME_OPTIONS_NL, key=lambda option: (abs(option - volume_nl), option))

    def _injection_step_plan(self, volume_nl: int) -> list[int]:
        remaining = volume_nl
        plan: list[int] = []
        for option in sorted(INJECTION_VOLUME_OPTIONS_NL, reverse=True):
            while remaining >= option:
                plan.append(option)
                remaining -= option
        if remaining > 0:
            plan.append(10)
        return plan

    def _main_injection_step_plan(self, volume_nl: int) -> list[int]:
        return [10] * max(1, int(math.ceil(volume_nl / 10.0)))

    def _injection_protocol_settings(self) -> InjectionProtocolSettings:
        volume_nl = self._rounded_single_injection_volume()
        insertion_rate_nl_min = max(self._line_float(self.insertion_injection_rate_nl_min, 100.0, 0.1, 100000.0), 0.1)
        main_rate_nl_min = max(self._line_float(self.main_injection_rate_nl_min, 100.0, 0.1, 100000.0), 0.1)
        injection_depth_mm = max(0.0, self._line_float(self.injection_depth_mm, 0.2, 0.0, 20.0))
        speed_um_s = max(self._line_float(self.insert_retract_speed_um_s, 20.0, 0.1, 10000.0), 0.1)
        overshoot_mm = max(0.0, self._line_float(self.movement_overshoot_mm, 0.05, 0.0, 10.0))
        pause_s = max(0.0, self._line_float(self.post_inject_pause_s, 5.0, 0.0, 3600.0))
        return InjectionProtocolSettings(
            main_volume_nl=volume_nl,
            insertion_rate_nl_min=insertion_rate_nl_min,
            main_rate_nl_min=main_rate_nl_min,
            injection_depth_mm=injection_depth_mm,
            insert_retract_speed_um_s=speed_um_s,
            overshoot_mm=overshoot_mm,
            post_inject_pause_s=pause_s,
        )

    def refresh_injection_sequence_summary(self) -> None:
        if not hasattr(self, "sequence_steps_list"):
            return
        settings = self._injection_protocol_settings()
        insertion_time_s, retract_time_s = self._insertion_retraction_times(settings)
        injection_duration_s = self._main_injection_duration_s(settings)
        site_count = len(self.injection_sites) or 1
        total_volume_nl, insertion_volume_nl, test_volume_total_nl = self._estimated_total_syringe_volume_nl(
            settings,
            site_count,
        )
        expected_duration_s = self._estimated_total_protocol_seconds(settings, site_count)
        self.sequence_steps_list.clear()
        steps = [
            (
                f"Total syringe volume required: {total_volume_nl:g} nl for {site_count} site(s) "
                f"(main {settings.main_volume_nl:g} nl/site + insertion {insertion_volume_nl:g} nl/site"
                + (f" + two blockage tests {test_volume_total_nl / site_count:g} nl/site)." if test_volume_total_nl else ").")
                + f" Expected timed protocol duration: {self._format_duration(expected_duration_s)} "
                "(excluding travel and blockage-confirmation time)."
            ),
            "Move to 1.000 mm above the stored surface, then move normally to the surface.",
            (
                f"Insert from surface to {settings.injection_depth_mm + settings.overshoot_mm:.3f} mm below surface at "
                f"{settings.insert_retract_speed_um_s:.1f} um/sec while injecting at "
                f"{settings.insertion_rate_nl_min:.1f} nl/min."
            ),
            f"Retract overshoot back to the target at {settings.insert_retract_speed_um_s:.1f} um/sec.",
        ]
        if injection_duration_s > insertion_time_s + retract_time_s + 0.05:
            steps.append(
                f"Continue injecting at target at {settings.main_rate_nl_min:.1f} nl/min "
                f"until {settings.main_volume_nl} nl total is delivered."
            )
        if settings.post_inject_pause_s > 0:
            steps.append(f"Pause at target for {settings.post_inject_pause_s:.1f} s.")
        steps.append(
            f"Retract to the stored surface at {settings.insert_retract_speed_um_s:.1f} um/sec, "
            "then move normally to 1.000 mm above the surface."
        )
        if self.block_check.isChecked():
            steps.append(
                "Run the blockage test from 1.000 mm above the stored surface; only continue if confirmed not blocked, "
                "otherwise offer repeated test injections until declined."
            )
        for index, text in enumerate(steps, start=1):
            item = QListWidgetItem(f"{index}. {text}")
            item.setSizeHint(QSize(0, 15))
            self.sequence_steps_list.addItem(item)

    def _estimated_total_syringe_volume_nl(
        self,
        settings: InjectionProtocolSettings,
        site_count: int,
    ) -> tuple[float, float, float]:
        """Conservative preparation volume: main, insertion delivery, and two tests/site."""
        insertion_time_s, retract_time_s = self._insertion_retraction_times(settings)
        insertion_volume_nl = settings.insertion_rate_nl_min * (insertion_time_s + retract_time_s) / 60.0
        test_volume_total_nl = 0.0
        if self.block_check.isChecked():
            test_volume_total_nl = 2.0 * self._rounded_test_volume() * site_count
        total_volume_nl = site_count * (settings.main_volume_nl + insertion_volume_nl) + test_volume_total_nl
        return total_volume_nl, insertion_volume_nl, test_volume_total_nl

    def _estimated_total_protocol_seconds(self, settings: InjectionProtocolSettings, site_count: int) -> float:
        """Timed portion only; positioning and user blockage decisions are not predictable."""
        speed_mm_s = max(settings.insert_retract_speed_um_s / 1000.0, 0.0001)
        return_time_s = settings.injection_depth_mm / speed_mm_s
        per_site_s = self._main_injection_duration_s(settings) + settings.post_inject_pause_s + return_time_s
        return max(0.0, site_count * per_site_s)

    def _sequence_step_indexes(self, settings: InjectionProtocolSettings, check_blocked: bool) -> dict[str, int]:
        indexes = {
            "approach": 0,
            "advance": 1,
            "retract": 2,
        }
        next_index = 3
        insertion_time_s, retract_time_s = self._insertion_retraction_times(settings)
        if self._main_injection_duration_s(settings) > insertion_time_s + retract_time_s + 0.05:
            indexes["main_injection"] = next_index
            next_index += 1
        if settings.post_inject_pause_s > 0:
            indexes["pause"] = next_index
            next_index += 1
        indexes["return"] = next_index
        next_index += 1
        if check_blocked:
            indexes["block"] = next_index
        return indexes

    def add_injection_site(self) -> None:
        if not self._require_idle("Add Injection Site"):
            return
        try:
            self._require_project_coordinates("injection_sites")
            ap, ml, dv = self.get_bregma_position()
            self.injection_sites.append(InjectionSite(ap=ap, ml=ml, dv=dv, generated=False))
            self.refresh_injection_sites_list()
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    @staticmethod
    def _grid_config_label(config: dict[str, object]) -> str:
        return (
            f"{int(config['ap_count'])} AP × {int(config['ml_count'])} ML; "
            f"AP {float(config['ap_spacing_mm']):g} mm, ML {float(config['ml_spacing_mm']):g} mm"
        )

    def _refresh_injection_grid_recent_combo(self) -> None:
        if not hasattr(self, "injection_grid_recent_combo"):
            return
        combo = self.injection_grid_recent_combo
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("Recently used…", None)
        for config in self.recent_injection_grid_configs:
            try:
                combo.addItem(self._grid_config_label(config), config)
            except (KeyError, TypeError, ValueError):
                continue
        combo.blockSignals(False)

    def _remember_injection_grid_config(self, config: dict[str, object]) -> None:
        keys = ("ap_count", "ml_count", "ap_spacing_mm", "ml_spacing_mm")
        self.recent_injection_grid_configs = [
            item for item in self.recent_injection_grid_configs
            if any(item.get(key) != config[key] for key in keys)
        ]
        self.recent_injection_grid_configs.insert(0, config)
        self.recent_injection_grid_configs = self.recent_injection_grid_configs[:8]
        self._save_general_settings()

    def add_injection_site_grid(self) -> None:
        if not self._require_idle("Add Injection Grid"):
            return
        if self.injection_sites_coordinate_system != "bregma":
            QMessageBox.warning(self, "Add Injection Grid", "Clear the older site list before adding a grid.")
            return
        if self.coordinate_mode != "bregma" or self.bregma_axis is None:
            QMessageBox.information(
                self,
                "Add Injection Grid",
                "Set Bregma and select Bregma coordinates before creating a grid.",
            )
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Injection Grid")
        layout = QGridLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setHorizontalSpacing(8)
        layout.setVerticalSpacing(6)
        layout.addWidget(QLabel("The grid is centred on the current Bregma AP/ML position. New sites remain unvalidated."), 0, 0, 1, 2)
        self.injection_grid_recent_combo = QComboBox()
        self._refresh_injection_grid_recent_combo()
        layout.addWidget(QLabel("Recently used"), 1, 0)
        layout.addWidget(self.injection_grid_recent_combo, 1, 1)
        ap_count = NumericLineEdit(2, minimum=1, maximum=20, integer=True)
        ml_count = NumericLineEdit(2, minimum=1, maximum=20, integer=True)
        ap_spacing = QLineEdit()
        ml_spacing = QLineEdit()
        ap_spacing.setPlaceholderText("mm")
        ml_spacing.setPlaceholderText("mm")
        layout.addWidget(QLabel("AP sites"), 2, 0)
        layout.addWidget(ap_count, 2, 1)
        layout.addWidget(QLabel("ML sites"), 3, 0)
        layout.addWidget(ml_count, 3, 1)
        layout.addWidget(QLabel("AP spacing (mm)"), 4, 0)
        layout.addWidget(ap_spacing, 4, 1)
        layout.addWidget(QLabel("ML spacing (mm)"), 5, 0)
        layout.addWidget(ml_spacing, 5, 1)

        def copy_spacing(source: QLineEdit, destination: QLineEdit) -> None:
            if source.text().strip() and not destination.text().strip():
                destination.setText(source.text().strip())

        ap_spacing.editingFinished.connect(lambda: copy_spacing(ap_spacing, ml_spacing))
        ml_spacing.editingFinished.connect(lambda: copy_spacing(ml_spacing, ap_spacing))

        def use_recent(index: int) -> None:
            config = self.injection_grid_recent_combo.itemData(index)
            if not isinstance(config, dict):
                return
            ap_count.setValue(int(config["ap_count"]))
            ml_count.setValue(int(config["ml_count"]))
            ap_spacing.setText(f"{float(config['ap_spacing_mm']):g}")
            ml_spacing.setText(f"{float(config['ml_spacing_mm']):g}")

        self.injection_grid_recent_combo.currentIndexChanged.connect(use_recent)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons, 6, 0, 1, 2)
        if dialog.exec() != QDialog.Accepted:
            return
        try:
            config = {
                "ap_count": int(ap_count.value()),
                "ml_count": int(ml_count.value()),
                "ap_spacing_mm": float(ap_spacing.text().strip().replace(",", ".")),
                "ml_spacing_mm": float(ml_spacing.text().strip().replace(",", ".")),
            }
            if config["ap_spacing_mm"] <= 0 or config["ml_spacing_mm"] <= 0:
                raise ValueError("Spacing must be greater than zero.")
            center_ap, center_ml, _center_dv = self.get_bregma_position()
            for ap_index in range(int(config["ap_count"])):
                ap = center_ap + (ap_index - (int(config["ap_count"]) - 1) / 2.0) * float(config["ap_spacing_mm"])
                for ml_index in range(int(config["ml_count"])):
                    ml = center_ml + (ml_index - (int(config["ml_count"]) - 1) / 2.0) * float(config["ml_spacing_mm"])
                    self.injection_sites.append(InjectionSite(ap=ap, ml=ml, dv=None, generated=True))
            self._remember_injection_grid_config(config)
            self.refresh_injection_sites_list()
            self.set_status(f"Added {int(config['ap_count']) * int(config['ml_count'])} unvalidated injection sites.")
        except (TypeError, ValueError) as exc:
            QMessageBox.warning(self, "Add Injection Grid", f"Enter valid grid settings. {exc}")
        finally:
            if hasattr(self, "injection_grid_recent_combo"):
                del self.injection_grid_recent_combo

    def _require_bregma_site_set_context(self, title: str) -> bool:
        if self.coordinate_mode == "bregma" and self.bregma_axis is not None:
            return True
        QMessageBox.information(
            self,
            title,
            "Set Bregma and select Bregma coordinates before saving or loading an injection site set.",
        )
        return False

    def save_injection_site_set(self) -> None:
        """Save Bregma AP/ML targets, deliberately excluding their surface validation."""
        try:
            self._require_project_coordinates("injection_sites")
        except Exception as exc:
            QMessageBox.warning(self, "Save Injection Sites", str(exc))
            return
        if not self._require_bregma_site_set_context("Save Injection Site Set"):
            return
        if not self.injection_sites:
            QMessageBox.information(self, "Save Injection Site Set", "Add one or more injection sites first.")
            return
        directory = self._config_dir("injection_sites")
        path_str, _selected = QFileDialog.getSaveFileName(
            self,
            "Save Injection Site Set",
            str(directory / "injection_site_set.json"),
            "JSON Files (*.json)",
        )
        if not path_str:
            return
        path = Path(path_str)
        if path.suffix.lower() != ".json":
            path = path.with_suffix(".json")
        payload: dict[str, object] = {
            "format": "neurostar-injection-site-set-v1",
            "coordinate_system": "bregma",
            "sites": [
                {"ap": site.ap, "ml": site.ml}
                for site in self.injection_sites
            ],
        }
        self._write_config_file(path, payload)
        self.set_status(f"Saved {len(self.injection_sites)} injection sites to {path}")

    def load_injection_site_set(self) -> None:
        """Load targets as unvalidated so their surfaces are always rechecked."""
        if not self._require_idle("Load Injection Sites"):
            return
        if not self._require_bregma_site_set_context("Load Injection Site Set"):
            return
        directory = self._config_dir("injection_sites")
        path_str, _selected = QFileDialog.getOpenFileName(
            self,
            "Load Injection Site Set",
            str(directory),
            "JSON Files (*.json)",
        )
        if not path_str:
            return
        try:
            payload = self._read_config_file(Path(path_str))
            if not isinstance(payload, dict):
                raise ValueError("The site-set file must contain a JSON object.")
            if payload.get("coordinate_system") not in {None, "bregma"}:
                raise ValueError("This site set is not stored in Bregma coordinates.")
            raw_sites = payload.get("sites")
            if not isinstance(raw_sites, list) or not raw_sites:
                raise ValueError("The file does not contain any injection sites.")
            sites: list[InjectionSite] = []
            for raw_site in raw_sites:
                if not isinstance(raw_site, dict):
                    raise ValueError("Each injection site must contain AP and ML coordinates.")
                ap = float(raw_site["ap"])
                ml = float(raw_site["ml"])
                if not math.isfinite(ap) or not math.isfinite(ml):
                    raise ValueError("Injection-site coordinates must be finite numbers.")
                sites.append(InjectionSite(ap=ap, ml=ml, dv=None, generated=True))
        except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
            QMessageBox.warning(self, "Load Injection Site Set", f"Could not load site set. {exc}")
            return
        if self.injection_sites:
            response = QMessageBox.question(
                self,
                "Replace Injection Sites?",
                f"Loading will replace the current {len(self.injection_sites)} injection site(s). Continue?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if response != QMessageBox.Yes:
                return
        if self.nudge_all_sites_active:
            self.nudge_all_sites_btn.setChecked(False)
        self.injection_sites = sites
        self.injection_sites_coordinate_system = "bregma"
        self.refresh_injection_sites_list()
        self.set_status(
            f"Loaded {len(sites)} injection sites. All are unvalidated; run Validate Sites before injection."
        )

    def remove_selected_injection_site(self) -> None:
        if not self._require_idle("Remove Injection Site"):
            return
        row = self.injection_sites_list.currentRow()
        if 0 <= row < len(self.injection_sites):
            del self.injection_sites[row]
            self.refresh_injection_sites_list()

    def clear_injection_sites(self) -> None:
        if not self._require_idle("Clear Injection Sites"):
            return
        self.injection_sites_coordinate_system = "bregma"
        self.injection_sites.clear()
        if self.nudge_all_sites_active:
            self.nudge_all_sites_btn.setChecked(False)
        if self.injection_sites_zoom_combo.currentIndex() == 1:
            self.injection_sites_zoom_combo.setCurrentIndex(2)
        self.refresh_injection_sites_list()

    def start_injection_site_validation(self) -> None:
        if not self._require_idle("Validate Sites"):
            return
        try:
            self._require_project_coordinates("injection_sites")
        except Exception as exc:
            QMessageBox.warning(self, "Validate Sites", str(exc))
            return
        if self._motion_is_active():
            QMessageBox.warning(self, "Validate Sites", "Wait for the current movement, drill, injection, benchmark, or probe to finish first.")
            return
        if not self.injection_sites:
            QMessageBox.information(self, "Validate Sites", "Add one or more injection sites first.")
            return
        if self.coordinate_mode != "bregma" or self.bregma_axis is None:
            QMessageBox.information(
                self,
                "Validate Sites",
                "Set Bregma and select Bregma coordinates before validating sites.",
            )
            return
        if self.nudge_all_sites_active:
            self.nudge_all_sites_btn.setChecked(False)
        choice = QMessageBox(self)
        choice.setWindowTitle("Validate Injection Sites")
        choice.setText("Choose which sites to step through.")
        validate_all = choice.addButton("Validate all", QMessageBox.AcceptRole)
        validate_unvalidated = choice.addButton("Validate unvalidated", QMessageBox.ActionRole)
        choice.addButton(QMessageBox.Cancel)
        choice.exec()
        if choice.clickedButton() == validate_all:
            self._run_injection_site_validation(0, validate_all_sites=True)
        elif choice.clickedButton() == validate_unvalidated:
            start_index = next((index for index, site in enumerate(self.injection_sites) if site.dv is None), None)
            if start_index is None:
                QMessageBox.information(self, "Validate Sites", "All injection sites are already validated.")
                return
            self._run_injection_site_validation(start_index, validate_all_sites=False)

    def _move_to_injection_site_for_validation(self, site: InjectionSite) -> bool:
        """Approach a site at Bregma DV -0.5 mm with a cancellable wait dialog."""
        target_axis = self._bregma_to_axis((site.ap, site.ml, -0.5))
        return self._move_through_axis_positions_with_progress(
            self._axis_clearance_path(target_axis, target_axis[2]),
            title="Moving to Injection Site",
            message="Moving to the site 0.5 mm above Bregma DV zero. Waiting for StereoDrive to report arrival…",
        )

    def _move_to_axis_position_with_progress(
        self,
        target_axis: tuple[float, float, float],
        *,
        title: str,
        message: str,
        delay_seconds: float = 0.75,
        stage_messages: list[str] | None = None,
    ) -> bool:
        return self._move_through_axis_positions_with_progress(
            self._axis_clearance_path(target_axis, self.bregma_axis[2] - 0.5)
            if self.bregma_axis is not None else [target_axis],
            title=title, message=message, delay_seconds=delay_seconds,
            stage_messages=stage_messages,
        )

    def _move_through_axis_positions_with_progress(
        self,
        target_axes: list[tuple[float, float, float]],
        *,
        title: str,
        message: str,
        delay_seconds: float = 0.75,
        stage_messages: list[str] | None = None,
    ) -> bool:
        """Move through one or more Axis targets with one cancellable progress dialog."""
        if not target_axes:
            return True
        if not self._require_idle(title):
            return False
        self.controller.prepare_motion()
        cancelled = threading.Event()
        result: dict[str, object] = {}
        progress_state = {"message": message}

        def move_worker() -> None:
            try:
                for target_index, target_axis in enumerate(target_axes):
                    if cancelled.is_set():
                        return
                    description = stage_messages[target_index] if stage_messages and target_index < len(stage_messages) else message
                    progress_state["message"] = (
                        f"{description}\n\nStage {target_index + 1}/{len(target_axes)}: "
                        f"waiting for Axis AP {target_axis[0]:.2f}, ML {target_axis[1]:.2f}, DV {target_axis[2]:.2f} mm."
                    )
                    self.controller.goto_axis_position(*target_axis, delay_seconds=delay_seconds,
                                                       stop_requested=cancelled.is_set)
                    self.controller.wait_for_axis_position(*target_axis, stop_requested=cancelled.is_set,
                        position_callback=self.validation_move_position_signal.emit)
                result["completed"] = True
            except Exception as exc:
                result["error"] = exc
                try:
                    self.controller.stop()
                    self.controller.wait_until_stopped()
                except Exception as stop_exc:
                    result["error"] = stop_exc

        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setModal(True)
        dialog.setWindowFlag(Qt.WindowCloseButtonHint, False)
        layout = QVBoxLayout(dialog)
        message_label = QLabel(message)
        message_label.setWordWrap(True)
        layout.addWidget(message_label)
        progress = QProgressBar()
        progress.setRange(0, 0)
        layout.addWidget(progress)
        cancel_button = QPushButton("Cancel Movement (Esc)")
        layout.addWidget(cancel_button)

        def request_cancel() -> None:
            if cancelled.is_set():
                return
            cancelled.set()
            cancel_button.setEnabled(False)
            progress_state["message"] = "Stopping StereoDrive movement and waiting for confirmation…"
            message_label.setText(progress_state["message"])
            try:
                self.controller.stop()
            except Exception:
                pass

        cancel_button.clicked.connect(request_cancel)
        worker = threading.Thread(target=move_worker, daemon=True)
        timer = QTimer(dialog)

        def check_worker() -> None:
            if message_label.text() != progress_state["message"]:
                message_label.setText(progress_state["message"])
            if not worker.is_alive():
                timer.stop()
                dialog.accept()

        timer.timeout.connect(check_worker)
        self.validation_move_active = True
        self.validation_move_cancel_callback = request_cancel
        worker.start()
        timer.start(50)
        try:
            dialog.exec()
        finally:
            timer.stop()
            self.validation_move_active = False
            self.validation_move_cancel_callback = None
        error = result.get("error")
        if isinstance(error, Exception) and (not cancelled.is_set() or "cancelled" not in str(error).lower()):
            raise error
        if cancelled.is_set():
            self.controller.wait_until_stopped()
            self.set_status("Movement cancelled; stopped position confirmed.")
            return False
        if not result.get("completed"):
            raise StereoDriveError("StereoDrive move ended without confirming arrival.")
        return True

    def _run_named_motion_with_progress(self, command, *, title: str, message: str) -> bool:
        """Run a Home/Work command when StereoDrive does not expose its target coordinates."""
        if not self._require_idle(title):
            return False
        self.controller.prepare_motion()
        initial_axis = self.controller.get_current_axis_position()
        cancelled = threading.Event()
        result: dict[str, object] = {}

        def move_worker() -> None:
            try:
                command()
                self.controller.wait_for_named_motion(initial_axis, stop_requested=cancelled.is_set,
                    position_callback=self.validation_move_position_signal.emit)
                result["completed"] = True
            except Exception as exc:
                result["error"] = exc
                try:
                    self.controller.stop()
                    self.controller.wait_until_stopped()
                except Exception as stop_exc:
                    result["error"] = stop_exc

        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.setModal(True)
        dialog.setWindowFlag(Qt.WindowCloseButtonHint, False)
        layout = QVBoxLayout(dialog)
        message_label = QLabel(message)
        message_label.setWordWrap(True)
        layout.addWidget(message_label)
        progress = QProgressBar()
        progress.setRange(0, 0)
        layout.addWidget(progress)
        cancel_button = QPushButton("Cancel Movement (Esc)")
        layout.addWidget(cancel_button)

        def request_cancel() -> None:
            if cancelled.is_set():
                return
            cancelled.set()
            cancel_button.setEnabled(False)
            message_label.setText("Stopping StereoDrive movement and waiting for confirmation…")
            try:
                self.controller.stop()
            except Exception:
                pass

        cancel_button.clicked.connect(request_cancel)
        worker = threading.Thread(target=move_worker, daemon=True)
        timer = QTimer(dialog)

        def check_worker() -> None:
            if not worker.is_alive():
                timer.stop()
                dialog.accept()

        timer.timeout.connect(check_worker)
        self.validation_move_active = True
        self.validation_move_cancel_callback = request_cancel
        worker.start()
        timer.start(50)
        try:
            dialog.exec()
        finally:
            timer.stop()
            self.validation_move_active = False
            self.validation_move_cancel_callback = None
        error = result.get("error")
        if isinstance(error, Exception) and (not cancelled.is_set() or "cancelled" not in str(error).lower()):
            raise error
        if cancelled.is_set():
            self.controller.wait_until_stopped()
            self.set_status("Movement cancelled; stopped position confirmed.")
            return False
        if not result.get("completed"):
            raise StereoDriveError("StereoDrive move ended without confirming arrival.")
        return True

    def _validation_dialog(self, index: int, total: int, site: InjectionSite) -> tuple[str, QDialog]:
        dialog = QDialog(self)
        dialog.setWindowTitle("Validate Injection Site")
        dialog.setModal(True)
        layout = QVBoxLayout(dialog)
        state = "unvalidated" if site.dv is None else f"surface DV {site.dv:.2f} mm"
        layout.addWidget(QLabel(
            f"Site {index + 1} of {total}: AP {site.ap:.2f}, ML {site.ml:.2f} ({state})\n\n"
            "The manipulator is at this site, 0.5 mm above Bregma DV zero. Use the normal keyboard nudges "
            "to refine AP/ML and lower DV until the tool touches surface."
        ))
        buttons = QDialogButtonBox()
        validate_button = buttons.addButton("Validate and Next", QDialogButtonBox.AcceptRole)
        next_button = buttons.addButton("Next Without Validating", QDialogButtonBox.ActionRole)
        delete_button = buttons.addButton("Delete Point", QDialogButtonBox.DestructiveRole)
        cancel_button = buttons.addButton(QDialogButtonBox.Cancel)
        layout.addWidget(buttons)
        result = {"action": "cancel"}

        def choose(action: str) -> None:
            result["action"] = action
            dialog.accept()

        validate_button.clicked.connect(lambda: choose("validate"))
        next_button.clicked.connect(lambda: choose("next"))
        delete_button.clicked.connect(lambda: choose("delete"))
        cancel_button.clicked.connect(dialog.reject)
        self.validation_modal_active = True
        try:
            dialog.exec()
        finally:
            self.validation_modal_active = False
        return str(result["action"]), dialog

    def _next_validation_index(self, index: int, validate_all_sites: bool) -> int | None:
        if validate_all_sites:
            return index + 1 if index + 1 < len(self.injection_sites) else None
        for candidate in range(index + 1, len(self.injection_sites)):
            if self.injection_sites[candidate].dv is None:
                return candidate
        return None

    def _run_injection_site_validation(self, start_index: int, validate_all_sites: bool) -> None:
        index: int | None = start_index
        try:
            while index is not None and index < len(self.injection_sites):
                site = self.injection_sites[index]
                self.refresh_injection_sites_list(active_index=index)
                self.set_status(f"Moving to injection site {index + 1}/{len(self.injection_sites)} for validation.")
                if not self._move_to_injection_site_for_validation(site):
                    break
                action, _dialog = self._validation_dialog(index, len(self.injection_sites), site)
                if action == "cancel":
                    self.set_status("Injection-site validation stopped.")
                    break
                if action == "validate":
                    ap, ml, dv = self.get_bregma_position()
                    self.injection_sites[index] = InjectionSite(ap=ap, ml=ml, dv=dv, generated=False)
                    self.set_status(f"Validated injection site {index + 1}.")
                elif action == "delete":
                    del self.injection_sites[index]
                    self.set_status(f"Deleted injection site {index + 1}.")
                    if validate_all_sites:
                        index = index if index < len(self.injection_sites) else None
                    else:
                        index = next(
                            (candidate for candidate in range(index, len(self.injection_sites))
                            if self.injection_sites[candidate].dv is None),
                            None,
                        )
                    continue
                index = self._next_validation_index(index, validate_all_sites)
            else:
                self.set_status("Injection-site validation complete.")
                QMessageBox.information(self, "Validate Sites", "No more sites remain in this validation pass.")
        except Exception as exc:
            QMessageBox.critical(self, "Validate Sites", str(exc))
        finally:
            self.refresh_injection_sites_list()

    def refresh_injection_sites_list(self, active_index: int | None = None) -> None:
        self.injection_sites_list.clear()
        for index, site in enumerate(self.injection_sites, start=1):
            if self.injection_sites_coordinate_system != "bregma":
                item = QListWidgetItem(f"{index}. AP {site.ap:.2f}, ML {site.ml:.2f} — reference unknown; recreate/load sites")
                item.setForeground(QColor("#a3a3a3"))
            elif site.dv is None:
                item = QListWidgetItem(f"{index}. AP {site.ap:.2f}, ML {site.ml:.2f} — unvalidated")
                item.setForeground(QColor("#a3a3a3"))
            else:
                item = QListWidgetItem(f"{index}. AP {site.ap:.2f}, ML {site.ml:.2f}, surface DV {site.dv:.2f}")
                item.setForeground(QColor("#111827"))
            self.injection_sites_list.addItem(item)
        if active_index is not None and 0 <= active_index < self.injection_sites_list.count():
            item = self.injection_sites_list.item(active_index)
            item.setBackground(QColor("#2563eb"))
            item.setForeground(QColor("#ffffff"))
            self.injection_sites_list.setCurrentRow(active_index)
            self.injection_sites_list.scrollToItem(item)
        self.refresh_injection_sequence_summary()
        if hasattr(self, "injection_sites_view"):
            self.redraw_views()

    def _active_injection_sites(self) -> list[InjectionSite]:
        self._require_project_coordinates("injection_sites")
        if self.injection_sites:
            unvalidated = [index + 1 for index, site in enumerate(self.injection_sites) if site.dv is None]
            if unvalidated:
                numbers = ", ".join(str(index) for index in unvalidated)
                raise StereoDriveError(f"Validate injection site(s) {numbers} before starting an injection.")
            return list(self.injection_sites)
        ap, ml, dv = self.get_bregma_position()
        return [InjectionSite(ap=ap, ml=ml, dv=dv)]

    def start_single_injection(self) -> None:
        if not self._require_idle("Injection"):
            return
        if self.injection_thread is not None and self.injection_thread.is_alive():
            return
        try:
            sites = self._active_injection_sites()
            settings = self._injection_protocol_settings()
            self._set_number_edit(self.single_injection_volume_nl, settings.main_volume_nl)
            self._set_number_edit(self.insertion_injection_rate_nl_min, settings.insertion_rate_nl_min)
            self._set_number_edit(self.main_injection_rate_nl_min, settings.main_rate_nl_min)
            self._set_number_edit(self.injection_depth_mm, settings.injection_depth_mm)
            self._set_number_edit(self.insert_retract_speed_um_s, settings.insert_retract_speed_um_s)
            self._set_number_edit(self.movement_overshoot_mm, settings.overshoot_mm)
            self._set_number_edit(self.post_inject_pause_s, settings.post_inject_pause_s)
            injection_plan = self._main_injection_step_plan(settings.main_volume_nl)
            if not injection_plan:
                return
            self.sync_syringe_position_before_injection()
            required_volume_nl, _insertion_volume_nl, _test_volume_total_nl = self._estimated_total_syringe_volume_nl(
                settings,
                len(sites),
            )
            test_volume_nl = self._rounded_test_volume()
            self.ensure_total_syringe_capacity(required_volume_nl)
            self._start_injection_sequence(
                sites=sites,
                settings=settings,
                injection_plan=injection_plan,
                check_blocked=self.block_check.isChecked(),
                test_volume_nl=test_volume_nl,
                start_site_offset=0,
                total_site_count=len(sites),
                initial_status=(
                    f"Ready: {len(sites)} site(s), {settings.main_volume_nl} nl; insertion "
                    f"{settings.insertion_rate_nl_min:.1f}, main {settings.main_rate_nl_min:.1f} nl/min"
                ),
            )
        except Exception as exc:
            QMessageBox.warning(self, "Injection", str(exc))

    def resume_injection_from_selected(self) -> None:
        if not self._require_idle("Injection"):
            return
        if self.injection_thread is not None and self.injection_thread.is_alive():
            return
        if not self.injection_sites:
            QMessageBox.information(self, "Injection", "Resume from selected requires stored injection sites.")
            return
        row = self.injection_sites_list.currentRow()
        if not (0 <= row < len(self.injection_sites)):
            QMessageBox.information(self, "Injection", "Select the site to resume from in the injection site list.")
            return
        unvalidated = [index + 1 for index, site in enumerate(self.injection_sites[row:], start=row) if site.dv is None]
        if unvalidated:
            QMessageBox.warning(
                self,
                "Injection",
                f"Validate injection site(s) {', '.join(str(index) for index in unvalidated)} before resuming.",
            )
            return
        try:
            self._require_project_coordinates("injection_sites")
            settings = self._injection_protocol_settings()
            self._set_number_edit(self.single_injection_volume_nl, settings.main_volume_nl)
            self._set_number_edit(self.insertion_injection_rate_nl_min, settings.insertion_rate_nl_min)
            self._set_number_edit(self.main_injection_rate_nl_min, settings.main_rate_nl_min)
            self._set_number_edit(self.injection_depth_mm, settings.injection_depth_mm)
            self._set_number_edit(self.insert_retract_speed_um_s, settings.insert_retract_speed_um_s)
            self._set_number_edit(self.movement_overshoot_mm, settings.overshoot_mm)
            self._set_number_edit(self.post_inject_pause_s, settings.post_inject_pause_s)
            remaining_sites = list(self.injection_sites[row:])
            injection_plan = self._main_injection_step_plan(settings.main_volume_nl)
            if not injection_plan:
                return
            self.sync_syringe_position_before_injection()
            required_volume_nl, _insertion_volume_nl, _test_volume_total_nl = self._estimated_total_syringe_volume_nl(
                settings,
                len(remaining_sites),
            )
            test_volume_nl = self._rounded_test_volume()
            self.ensure_total_syringe_capacity(required_volume_nl)
            self._start_injection_sequence(
                sites=remaining_sites,
                settings=settings,
                injection_plan=injection_plan,
                check_blocked=self.block_check.isChecked(),
                test_volume_nl=test_volume_nl,
                start_site_offset=row,
                total_site_count=len(self.injection_sites),
                initial_status=f"Resuming from selected site {row + 1}/{len(self.injection_sites)}",
            )
        except Exception as exc:
            QMessageBox.warning(self, "Injection", str(exc))

    def _start_injection_sequence(
        self,
        sites: list[InjectionSite],
        settings: InjectionProtocolSettings,
        injection_plan: list[int],
        check_blocked: bool,
        test_volume_nl: int,
        start_site_offset: int,
        total_site_count: int,
        initial_status: str,
    ) -> None:
        self._require_project_coordinates("injection_sites")
        # Saved sites stay GUI-Bregma-relative; workers receive frozen Axis targets.
        sites = [InjectionSite(*self._bregma_to_axis((site.ap, site.ml, site.dv)), generated=site.generated)
                 for site in sites]
        self.injection_clearance_axis_dv = self.bregma_axis[2] - 0.5
        self.controller.prepare_motion()
        if self.nudge_all_sites_active:
            self.nudge_all_sites_btn.setChecked(False)
        self.injection_pause_requested.clear()
        self.injection_stop_requested.clear()
        self.injection_progress.setValue(int((start_site_offset / max(1, total_site_count)) * 100))
        self.injection_site_progress.setValue(0)
        self.set_status(initial_status)
        self.start_injection_btn.setEnabled(False)
        self.pause_injection_btn.setText("Pause")
        self.injection_thread = threading.Thread(
            target=self._run_injection_protocol,
            args=(
                sites,
                settings,
                injection_plan,
                check_blocked,
                test_volume_nl,
                start_site_offset,
                total_site_count,
            ),
            daemon=True,
        )
        self.injection_thread.start()

    def pause_resume_injection(self) -> None:
        if self.injection_thread is None or not self.injection_thread.is_alive():
            return
        if self.injection_pause_requested.is_set():
            self.injection_pause_requested.clear()
            self.pause_injection_btn.setText("Pause")
            self.set_status("Injection resumed")
        else:
            self.injection_pause_requested.set()
            self.pause_injection_btn.setText("Resume")
            self.set_status("Injection paused")

    def stop_injection(self) -> None:
        self.injection_stop_requested.set()
        self.set_status("Stopping injection")
        try:
            self.controller.stop()
        except Exception as exc:
            self.set_status(f"Could not stop manipulator: {exc}")
        try:
            self.controller.stop_injectomate_motion()
        except Exception:
            pass
        self._start_syringe_position_scale_read(wait_for_injection_thread=True)

    def _rounded_test_volume(self) -> int:
        volume_nl = self._nearest_supported_injection_volume(self._line_int(self.block_test_volume_nl, 50, 10, 2000))
        self._set_number_edit(self.block_test_volume_nl, volume_nl)
        return volume_nl

    def _run_injection_protocol(
        self,
        sites: list[InjectionSite],
        settings: InjectionProtocolSettings,
        injection_plan: list[int],
        check_blocked: bool,
        test_volume_nl: int,
        start_site_offset: int = 0,
        total_site_count: int | None = None,
    ) -> None:
        try:
            total_units = max(1, total_site_count if total_site_count is not None else len(sites))
            step_indexes = self._sequence_step_indexes(settings, check_blocked)
            for relative_site_index, site in enumerate(sites, start=1):
                site_index = start_site_offset + relative_site_index
                if self.injection_stop_requested.is_set():
                    break
                self.active_injection_site_signal.emit(site_index - 1)
                self.sequence_step_signal.emit(step_indexes["approach"])
                self.injection_progress_signal.emit(
                    int(((site_index - 1) / total_units) * 100),
                    f"Moving to 1 mm above surface for site {site_index}/{total_units}",
                )
                above_dv = self._above_surface_dv(site)
                self._approach_axis_position((site.ap, site.ml, above_dv),
                                             self.injection_clearance_axis_dv,
                                             self.injection_stop_requested.is_set)
                self.controller.wait_for_axis_position(
                    site.ap,
                    site.ml,
                    above_dv,
                    tolerance_mm=0.02,
                    timeout_seconds=60.0,
                    poll_seconds=0.1,
                    stop_requested=self.injection_stop_requested.is_set,
                )
                if self.injection_stop_requested.is_set():
                    break
                self._run_protocol_at_site(
                    site,
                    settings,
                    injection_plan,
                    step_indexes,
                    site_index,
                    total_units,
                )
                if check_blocked and not self.injection_stop_requested.is_set():
                    self.sequence_step_signal.emit(step_indexes["block"])
                    self._run_block_test(site, settings, test_volume_nl)
            if self.injection_stop_requested.is_set():
                self.controller.stop()
                self.controller.wait_until_stopped()
                self.sequence_step_signal.emit(-1)
                self.active_injection_site_signal.emit(-1)
                self.injection_finished_signal.emit("Injection stopped")
            else:
                self.sequence_step_signal.emit(-1)
                self.active_injection_site_signal.emit(-1)
                self.injection_finished_signal.emit("Injection protocol complete")
        except Exception as exc:
            try:
                self.controller.stop()
                self.controller.wait_until_stopped()
            except Exception as stop_exc:
                self.sequence_step_signal.emit(-1)
                self.active_injection_site_signal.emit(-1)
                self.injection_finished_signal.emit(f"Stop confirmation failed: {stop_exc}. Original error: {exc}")
                return
            if self.injection_stop_requested.is_set():
                self.sequence_step_signal.emit(-1)
                self.active_injection_site_signal.emit(-1)
                self.injection_finished_signal.emit("Injection stopped")
                return
            message = str(exc)
            if "Maximum movement" in message or "limit" in message:
                self.syringe_limit_warning_signal.emit(message)
            self.sequence_step_signal.emit(-1)
            self.active_injection_site_signal.emit(-1)
            self.injection_finished_signal.emit(str(exc))

    def _run_protocol_at_site(
        self,
        site: InjectionSite,
        settings: InjectionProtocolSettings,
        injection_plan: list[int],
        step_indexes: dict[str, int],
        site_index: int,
        site_count: int,
    ) -> None:
        self.sequence_step_signal.emit(step_indexes["approach"])
        self.injection_progress_signal.emit(
            int(((site_index - 1) / max(1, site_count)) * 100),
            f"Moving to surface for injection site {site_index}/{site_count}",
        )
        self.controller.goto_axis_position(site.ap, site.ml, site.dv, delay_seconds=0.5,
                                           stop_requested=self.injection_stop_requested.is_set)
        self.controller.wait_for_axis_position(
            site.ap,
            site.ml,
            site.dv,
            tolerance_mm=0.02,
            timeout_seconds=60.0,
            poll_seconds=0.1,
            stop_requested=self.injection_stop_requested.is_set,
        )
        if self.injection_stop_requested.is_set():
            return

        insertion_time_s, retract_time_s = self._insertion_retraction_times(settings)
        movement_total_s = insertion_time_s + retract_time_s
        injection_duration_s = self._main_injection_duration_s(settings)
        active_work_s = max(injection_duration_s, movement_total_s)
        protocol_duration_s = max(active_work_s, 0.1)
        injection_events = self._scheduled_main_injection_events(injection_plan, settings)
        movement_targets = self._protocol_movement_targets(site, settings)
        next_movement_endpoint = 1
        endpoint_step_mm, endpoint_dwell_s = self._slow_axis_step_and_dwell(settings)
        delivered = 0
        event_index = 0
        start_time = time.monotonic()
        last_move_at = 0.0
        while not self.injection_stop_requested.is_set():
            start_time += self._wait_while_injection_paused()
            elapsed = time.monotonic() - start_time
            # Do not skip overshoot/depth endpoints when a UI/syringe call
            # takes longer than a sampling interval.
            while next_movement_endpoint < len(movement_targets) and elapsed >= movement_targets[next_movement_endpoint][0]:
                self.controller.move_axis_to_target(
                    "DV", movement_targets[next_movement_endpoint][1], step_mm=endpoint_step_mm,
                    stop_requested=self.injection_stop_requested.is_set, dwell_seconds=endpoint_dwell_s,
                )
                next_movement_endpoint += 1
            if elapsed >= protocol_duration_s and event_index >= len(injection_events):
                break
            while event_index < len(injection_events) and elapsed >= injection_events[event_index][0]:
                step_nl = injection_events[event_index][1]
                self.ensure_syringe_move_allowed(step_nl, False)
                self.controller.syringe_step(
                    f"{step_nl} nl",
                    up=False,
                    stop_requested=self.injection_stop_requested.is_set,
                )
                self.track_injection_delivery(step_nl)
                delivered += step_nl
                event_index += 1
            if movement_targets and elapsed - last_move_at >= 0.05:
                target_axis_dv = self._interpolated_movement_dv(movement_targets, elapsed)
                self.controller.move_axis_to_target(
                    "DV",
                    target_axis_dv,
                    step_mm=5.0,
                    stop_requested=self.injection_stop_requested.is_set,
                    status_callback=None,
                    dwell_seconds=0.002,
                )
                last_move_at = elapsed
            site_fraction = min(1.0, elapsed / protocol_duration_s)
            total_fraction = ((site_index - 1) + site_fraction) / max(1, site_count)
            current_volume = min(delivered, settings.main_volume_nl)
            in_insertion_phase = movement_targets and elapsed < insertion_time_s
            in_retraction_phase = movement_targets and insertion_time_s <= elapsed < movement_total_s
            if in_insertion_phase and current_volume < settings.main_volume_nl:
                self.sequence_step_signal.emit(step_indexes["advance"])
                message = f"Inserting pipette while injecting (current volume = {current_volume} nl)"
            elif in_retraction_phase and current_volume < settings.main_volume_nl:
                self.sequence_step_signal.emit(step_indexes["retract"])
                message = f"Retracting overshoot while injecting (current volume = {current_volume} nl)"
            elif current_volume < settings.main_volume_nl:
                self.sequence_step_signal.emit(step_indexes.get("main_injection", step_indexes["retract"]))
                message = f"Injecting (current volume = {current_volume} nl)"
            elif in_insertion_phase:
                self.sequence_step_signal.emit(step_indexes["advance"])
                message = "Inserting pipette"
            elif in_retraction_phase:
                self.sequence_step_signal.emit(step_indexes["retract"])
                message = "Retracting overshoot"
            else:
                message = f"Injection/movement complete at site {site_index}/{site_count}"
            self.injection_progress_signal.emit(
                int(total_fraction * 100),
                message,
            )
            self.injection_site_progress_signal.emit(int(site_fraction * 100))
            time.sleep(0.02)
        if self.injection_stop_requested.is_set():
            return
        # The timed loop may finish between samples. Confirm its final depth
        # before the post-injection hold or surface retraction starts.
        if movement_targets:
            final_step_mm, final_dwell_s = self._slow_axis_step_and_dwell(settings)
            self.controller.move_axis_to_target(
                "DV", movement_targets[-1][1], step_mm=final_step_mm,
                stop_requested=self.injection_stop_requested.is_set, dwell_seconds=final_dwell_s,
            )
        if settings.post_inject_pause_s > 0:
            self.sequence_step_signal.emit(step_indexes["pause"])
            pause_started = time.monotonic()
            while not self.injection_stop_requested.is_set():
                pause_elapsed = time.monotonic() - pause_started
                remaining_s = settings.post_inject_pause_s - pause_elapsed
                if remaining_s <= 0:
                    break
                site_fraction = min(1.0, (active_work_s + pause_elapsed) / (active_work_s + settings.post_inject_pause_s))
                total_fraction = ((site_index - 1) + site_fraction) / max(1, site_count)
                self.injection_progress_signal.emit(
                    int(total_fraction * 100),
                    f"Post-injection pause at site {site_index}/{site_count}: {remaining_s:.1f}s remaining",
                )
                self.injection_site_progress_signal.emit(int(site_fraction * 100))
                time.sleep(0.05)
        if self.injection_stop_requested.is_set():
            return
        self.sequence_step_signal.emit(step_indexes["return"])
        self.injection_progress_signal.emit(
            int((site_index / max(1, site_count)) * 100),
            f"Retracting to surface at {settings.insert_retract_speed_um_s:.1f} um/sec for site {site_index}/{site_count}",
        )
        retract_step_mm, retract_dwell_s = self._slow_axis_step_and_dwell(settings)
        self.controller.move_axis_to_target(
            "DV",
            site.dv,
            step_mm=retract_step_mm,
            tolerance=0.003,
            stop_requested=self.injection_stop_requested.is_set,
            dwell_seconds=retract_dwell_s,
        )
        if self.injection_stop_requested.is_set():
            return
        above_dv = self._above_surface_dv(site)
        self.injection_progress_signal.emit(
            int((site_index / max(1, site_count)) * 100),
            f"Moving normally to 1 mm above surface for site {site_index}/{site_count}",
        )
        self.controller.goto_axis_position(site.ap, site.ml, above_dv, delay_seconds=0.5,
                                           stop_requested=self.injection_stop_requested.is_set)
        self.controller.wait_for_axis_position(
            site.ap,
            site.ml,
            above_dv,
            tolerance_mm=0.02,
            timeout_seconds=60.0,
            poll_seconds=0.1,
            stop_requested=self.injection_stop_requested.is_set,
        )

    def _scheduled_main_injection_events(
        self,
        plan: list[int],
        settings: InjectionProtocolSettings,
    ) -> list[tuple[float, int]]:
        insertion_time_s, retract_time_s = self._insertion_retraction_times(settings)
        insertion_window_s = insertion_time_s + retract_time_s
        insertion_volume_nl = (settings.insertion_rate_nl_min / 60.0) * insertion_window_s
        delivered = 0
        events: list[tuple[float, int]] = []
        for step_nl in plan:
            delivered += step_nl
            if delivered <= insertion_volume_nl:
                event_time_s = (delivered / max(settings.insertion_rate_nl_min, 0.1)) * 60.0
            else:
                remaining_after_insert_nl = delivered - insertion_volume_nl
                event_time_s = insertion_window_s + (remaining_after_insert_nl / max(settings.main_rate_nl_min, 0.1)) * 60.0
            events.append((event_time_s, step_nl))
        return events

    def _target_injection_site_dv(self, site: InjectionSite, settings: InjectionProtocolSettings) -> float:
        return site.dv + settings.injection_depth_mm

    def _above_surface_dv(self, site: InjectionSite) -> float:
        return site.dv - 1.0

    def _bregma_dv_to_axis(self, bregma_dv: float) -> float:
        """Translate a GUI-Bregma DV value to mechanical Axis DV."""
        return self._bregma_to_axis((0.0, 0.0, bregma_dv))[2]

    def _main_injection_duration_s(self, settings: InjectionProtocolSettings) -> float:
        insertion_time_s, retract_time_s = self._insertion_retraction_times(settings)
        insertion_window_s = insertion_time_s + retract_time_s
        insertion_volume_nl = (settings.insertion_rate_nl_min / 60.0) * insertion_window_s
        remaining_volume_nl = max(0.0, settings.main_volume_nl - insertion_volume_nl)
        main_delivery_s = (remaining_volume_nl / max(settings.main_rate_nl_min, 0.1)) * 60.0
        return insertion_window_s + main_delivery_s

    def _insertion_retraction_times(self, settings: InjectionProtocolSettings) -> tuple[float, float]:
        speed_mm_s = max(settings.insert_retract_speed_um_s / 1000.0, 0.0001)
        insertion_time_s = (settings.injection_depth_mm + settings.overshoot_mm) / speed_mm_s
        retract_time_s = settings.overshoot_mm / speed_mm_s
        return insertion_time_s, retract_time_s

    def _slow_axis_step_and_dwell(self, settings: InjectionProtocolSettings) -> tuple[float, float]:
        step_mm = 0.005
        speed_mm_s = max(settings.insert_retract_speed_um_s / 1000.0, 0.0001)
        return step_mm, step_mm / speed_mm_s

    def _protocol_movement_targets(
        self,
        site: InjectionSite,
        settings: InjectionProtocolSettings,
    ) -> list[tuple[float, float]]:
        insertion_time_s, retract_time_s = self._insertion_retraction_times(settings)
        target_dv = self._target_injection_site_dv(site, settings)
        overshoot_dv = target_dv + settings.overshoot_mm
        targets = [
            (0.0, site.dv),
            (insertion_time_s, overshoot_dv),
            (insertion_time_s + retract_time_s, target_dv),
        ]
        return targets

    def _interpolated_movement_dv(self, targets: list[tuple[float, float]], elapsed_s: float) -> float:
        if elapsed_s <= targets[0][0]:
            return targets[0][1]
        for (start_t, start_dv), (end_t, end_dv) in zip(targets[:-1], targets[1:]):
            if start_t <= elapsed_s <= end_t:
                fraction = 1.0 if math.isclose(start_t, end_t) else (elapsed_s - start_t) / (end_t - start_t)
                return start_dv + fraction * (end_dv - start_dv)
        return targets[-1][1]

    def _wait_while_injection_paused(self) -> float:
        if not self.injection_pause_requested.is_set():
            return 0.0
        paused_at = time.monotonic()
        while self.injection_pause_requested.is_set() and not self.injection_stop_requested.is_set():
            time.sleep(0.05)
        return time.monotonic() - paused_at

    def _run_block_test(
        self,
        site: InjectionSite,
        settings: InjectionProtocolSettings,
        test_volume_nl: int,
    ) -> None:
        self.injection_progress_signal.emit(100, "Retracting pipette")
        above_dv = self._above_surface_dv(site)
        self.controller.goto_axis_position(site.ap, site.ml, above_dv, delay_seconds=0.5,
                                           stop_requested=self.injection_stop_requested.is_set)
        self.controller.wait_for_axis_position(
            site.ap,
            site.ml,
            above_dv,
            tolerance_mm=0.02,
            timeout_seconds=60.0,
            poll_seconds=0.1,
            stop_requested=self.injection_stop_requested.is_set,
        )
        while not self.injection_stop_requested.is_set():
            QApplication.beep()
            for remaining in range(5, 0, -1):
                if self.injection_stop_requested.is_set():
                    return
                self.injection_progress_signal.emit(100, f"Verifying no blockage in {remaining}s")
                time.sleep(1.0)
            for step_nl in self._injection_step_plan(test_volume_nl):
                if self.injection_stop_requested.is_set():
                    return
                self.injection_progress_signal.emit(100, f"Verifying no blockage (test volume = {step_nl} nl)")
                self.ensure_syringe_move_allowed(step_nl, False)
                self.controller.syringe_step(
                    f"{step_nl} nl",
                    up=False,
                    stop_requested=self.injection_stop_requested.is_set,
                )
                self.track_injection_delivery(step_nl)
            self.block_prompt_event = threading.Event()
            self.block_prompt_result = "clear"
            self.block_prompt_signal.emit()
            while not self.injection_stop_requested.is_set():
                if self.block_prompt_event.wait(timeout=0.1):
                    break
            if self.injection_stop_requested.is_set():
                return
            if self.block_prompt_result == "clear":
                self.injection_progress_signal.emit(100, "Blockage test confirmed clear")
                return
            if self.block_prompt_result == "retest":
                self.injection_progress_signal.emit(100, "Repeating blockage test injection")
                continue
            self.injection_stop_requested.set()
            self.injection_progress_signal.emit(100, "Sequence stopped after blockage test")
            return

    def set_injection_progress(self, percent: int, message: str) -> None:
        self.injection_progress.setValue(max(0, min(100, percent)))
        self.set_status(message)

    def set_injection_site_progress(self, percent: int) -> None:
        self.injection_site_progress.setValue(max(0, min(100, percent)))

    def set_active_sequence_step(self, row: int) -> None:
        if not hasattr(self, "sequence_steps_list"):
            return
        for index in range(self.sequence_steps_list.count()):
            item = self.sequence_steps_list.item(index)
            font = item.font()
            if font.pointSize() <= 0:
                font.setPointSize(8)
            font.setBold(index == row)
            item.setFont(font)
        if 0 <= row < self.sequence_steps_list.count():
            self.sequence_steps_list.setCurrentRow(row)
            self.sequence_steps_list.scrollToItem(self.sequence_steps_list.item(row))
        else:
            self.sequence_steps_list.clearSelection()
            self.sequence_steps_list.setCurrentRow(-1)

    def set_active_injection_site(self, row: int) -> None:
        if not hasattr(self, "injection_sites_list"):
            return
        for index in range(self.injection_sites_list.count()):
            item = self.injection_sites_list.item(index)
            font = item.font()
            if font.pointSize() <= 0:
                font.setPointSize(9)
            font.setBold(index == row)
            item.setFont(font)
        if 0 <= row < self.injection_sites_list.count():
            self.injection_sites_list.setCurrentRow(row)
            self.injection_sites_list.scrollToItem(self.injection_sites_list.item(row))
        else:
            self.injection_sites_list.clearSelection()
            self.injection_sites_list.setCurrentRow(-1)

    def finish_injection(self, message: str) -> None:
        self.set_active_sequence_step(-1)
        self.set_active_injection_site(-1)
        display_message = "Sequence complete" if message == "Injection protocol complete" else message
        self.set_status(display_message)
        self.start_injection_btn.setEnabled(True)
        self.pause_injection_btn.setText("Pause")
        if message in ("Injection complete", "Injection protocol complete"):
            self.injection_progress.setValue(100)
            self.injection_site_progress.setValue(100)
        self.injection_pause_requested.clear()
        self.injection_stop_requested.clear()
        self.injection_thread = None

    def show_block_prompt(self) -> None:
        def exec_with_alert(box: QMessageBox) -> None:
            alert_timer = QTimer(box)
            alert_timer.timeout.connect(QApplication.beep)
            QApplication.beep()
            alert_timer.start(900)
            try:
                box.exec()
            finally:
                alert_timer.stop()

        result = "clear"
        box = QMessageBox(self)
        box.setWindowTitle("Blockage Check")
        box.setText("Is the test injection confirmed not blocked?")
        clear_button = box.addButton("Not blocked", QMessageBox.AcceptRole)
        blocked_button = box.addButton("Blocked", QMessageBox.RejectRole)
        box.setDefaultButton(clear_button)
        exec_with_alert(box)
        if box.clickedButton() == blocked_button:
            retry_box = QMessageBox(self)
            retry_box.setWindowTitle("Blockage Check")
            retry_box.setText("Blocked reported. Do another test injection?")
            retry_button = retry_box.addButton("Another test", QMessageBox.AcceptRole)
            stop_button = retry_box.addButton("No", QMessageBox.RejectRole)
            retry_box.setDefaultButton(retry_button)
            exec_with_alert(retry_box)
            result = "retest" if retry_box.clickedButton() == retry_button else "stop"
        self.block_prompt_result = result
        if self.block_prompt_event is not None:
            self.block_prompt_event.set()

    def _is_blue_plunger_pixel(self, image: QImage, x: int, y: int) -> bool:
        color = image.pixelColor(x, y)
        red = color.red()
        green = color.green()
        blue = color.blue()
        return blue > 120 and blue > red * 1.5 and blue > green * 1.2

    def _normalized_blue_mask(
        self, image: QImage, x_start: int, x_end: int, y_start: int, y_end: int
    ) -> set[tuple[int, int]]:
        source_width = max(1, x_end - x_start + 1)
        source_height = max(1, y_end - y_start + 1)
        mask: set[tuple[int, int]] = set()
        for y in range(y_start, y_end + 1):
            for x in range(x_start, x_end + 1):
                if self._is_blue_plunger_pixel(image, x, y):
                    nx = min(11, int((x - x_start) * 12 / source_width))
                    ny = min(17, int((y - y_start) * 18 / source_height))
                    mask.add((nx, ny))
        return mask

    def _plunger_digit_templates(self) -> list[tuple[str, set[tuple[int, int]]]]:
        templates = getattr(self, "_cached_plunger_digit_templates", None)
        if templates is not None:
            return templates

        templates = []
        self._cached_plunger_digit_templates = templates
        return templates

    def _recognize_plunger_digit(self, image: QImage, x_start: int, x_end: int, y_start: int, y_end: int) -> str | None:
        mask = self._normalized_blue_mask(image, x_start, x_end, y_start, y_end)
        if not mask:
            return None

        best_digit = None
        best_score = 0.0
        for digit, template in self._plunger_digit_templates():
            overlap = len(mask & template)
            total = len(mask | template)
            if total <= 0:
                continue
            score = overlap / total
            if score > best_score:
                best_score = score
                best_digit = digit
        if best_score < 0.22:
            return None
        return best_digit

    def _blue_digit_groups(self, image: QImage) -> list[tuple[int, int, int, int]]:
        image_width = image.width()
        image_height = image.height()
        y_start = max(0, int(image_height * 0.80))
        y_end = image_height - 1

        blue_columns: list[int] = []
        for x in range(image_width):
            hits = 0
            for y in range(y_start, y_end + 1):
                if self._is_blue_plunger_pixel(image, x, y):
                    hits += 1
            if hits >= 2:
                blue_columns.append(x)
        if not blue_columns:
            return []

        groups: list[tuple[int, int]] = []
        group_start = blue_columns[0]
        previous = blue_columns[0]
        for x in blue_columns[1:]:
            if x - previous > 2:
                groups.append((group_start, previous))
                group_start = x
            previous = x
        groups.append((group_start, previous))

        digit_groups: list[tuple[int, int, int, int]] = []
        for x0, x1 in groups:
            if x1 - x0 < 2:
                continue
            ys: list[int] = []
            for y in range(y_start, y_end + 1):
                for x in range(x0, x1 + 1):
                    if self._is_blue_plunger_pixel(image, x, y):
                        ys.append(y)
            if not ys:
                continue
            if max(ys) - min(ys) < 5:
                continue
            digit_groups.append((x0, x1, min(ys), max(ys)))
        return digit_groups

    def _read_plunger_text_from_image(self, image: QImage) -> float | None:
        digits: list[str] = []
        for x0, x1, y0, y1 in self._blue_digit_groups(image):
            digit = self._recognize_plunger_digit(image, x0, x1, y0, y1)
            if digit is None:
                return None
            digits.append(digit)

        if not digits:
            return None
        value = int("".join(digits))
        if 0 <= value <= 5000:
            return float(value)
        return None

    def _plunger_gauge_capture_context(self) -> tuple[QImage, dict[str, object]] | None:
        rect = self.controller.get_mmc_depth_gauge_rect()
        if rect is None:
            return None
        left, top, width, height = rect
        center = QPoint(left + width // 2, top + height // 2)
        screen = QApplication.screenAt(center) or QApplication.primaryScreen()
        if screen is None:
            return None
        dpr = max(1.0, float(screen.devicePixelRatio()))
        logical_left = int(round(left / dpr))
        logical_top = int(round(top / dpr))
        logical_width = max(1, int(round(width / dpr)))
        logical_height = max(1, int(round(height / dpr)))
        image = None
        capture_method = "screen"
        pixmap = screen.grabWindow(0, logical_left, logical_top, logical_width, logical_height)
        if not pixmap.isNull():
            screen_image = pixmap.toImage()
            if screen_image.width() > 1 and screen_image.height() > 1 and self._blue_digit_groups(screen_image):
                image = screen_image

        if image is None:
            hwnd = self.controller.get_mmc_depth_gauge_handle()
            if hwnd is not None:
                image = self._capture_window_with_print_window(hwnd, width, height)
                capture_method = "print_window"

        if image is None or image.width() <= 1 or image.height() <= 1:
            return None
        return (
            image,
            {
                "gauge_rect_physical": rect,
                "capture_rect_logical": (logical_left, logical_top, logical_width, logical_height),
                "screen_name": screen.name(),
                "screen_geometry": (
                    screen.geometry().x(),
                    screen.geometry().y(),
                    screen.geometry().width(),
                    screen.geometry().height(),
                ),
                "device_pixel_ratio": dpr,
                "capture_method": capture_method,
            },
        )

    def _capture_plunger_gauge_image(self) -> QImage | None:
        context = self._plunger_gauge_capture_context()
        if context is None:
            return None
        image, _metadata = context
        return image

    def _capture_window_with_print_window(self, hwnd: int, width: int, height: int) -> QImage | None:
        window_dc = user32.GetDC(hwnd)
        if not window_dc:
            return None
        memory_dc = gdi32.CreateCompatibleDC(window_dc)
        bitmap = gdi32.CreateCompatibleBitmap(window_dc, width, height)
        if not memory_dc or not bitmap:
            if bitmap:
                gdi32.DeleteObject(bitmap)
            if memory_dc:
                gdi32.DeleteDC(memory_dc)
            user32.ReleaseDC(hwnd, window_dc)
            return None
        old_object = gdi32.SelectObject(memory_dc, bitmap)
        try:
            if not user32.PrintWindow(hwnd, memory_dc, PW_RENDERFULLCONTENT):
                return None
            bitmap_info = BITMAPINFO()
            bitmap_info.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            bitmap_info.bmiHeader.biWidth = width
            bitmap_info.bmiHeader.biHeight = -height
            bitmap_info.bmiHeader.biPlanes = 1
            bitmap_info.bmiHeader.biBitCount = 32
            bitmap_info.bmiHeader.biCompression = BI_RGB
            buffer_size = width * height * 4
            buffer = (ctypes.c_ubyte * buffer_size)()
            rows = gdi32.GetDIBits(
                memory_dc,
                bitmap,
                0,
                height,
                ctypes.cast(buffer, ctypes.c_void_p),
                ctypes.byref(bitmap_info),
                DIB_RGB_COLORS,
            )
            if rows == 0:
                return None
            image = QImage(bytes(buffer), width, height, QImage.Format_BGRA8888)
            return image.copy()
        finally:
            if old_object:
                gdi32.SelectObject(memory_dc, old_object)
            gdi32.DeleteObject(bitmap)
            gdi32.DeleteDC(memory_dc)
            user32.ReleaseDC(hwnd, window_dc)

    def _blue_filtered_plunger_image(self, image: QImage) -> QImage:
        filtered = QImage(image.size(), QImage.Format_ARGB32)
        filtered.fill(QColor("#ffffff"))
        for y in range(image.height()):
            for x in range(image.width()):
                if self._is_blue_plunger_pixel(image, x, y):
                    filtered.setPixelColor(x, y, image.pixelColor(x, y))
        return filtered

    def save_plunger_debug_capture(self) -> None:
        try:
            context = self._plunger_gauge_capture_context()
            if context is None:
                raise StereoDriveError("Could not capture MMCDepth plunger gauge.")
            image, metadata = context
            filtered = self._blue_filtered_plunger_image(image)
            value = self._read_plunger_text_from_image(image)
            output_dir = Path(__file__).resolve().parent
            raw_path = output_dir / "plunger_debug_raw.png"
            filtered_path = output_dir / "plunger_debug_filtered.png"
            value_path = output_dir / "plunger_debug_value.txt"
            image.save(str(raw_path))
            filtered.save(str(filtered_path))
            value_text = "--" if value is None else f"{value:.0f}"
            value_path.write_text(
                "\n".join(
                    [
                        f"detected_value_nl={value_text}",
                        f"gauge_rect_physical={metadata['gauge_rect_physical']}",
                        f"capture_rect_logical={metadata['capture_rect_logical']}",
                        f"screen_name={metadata['screen_name']}",
                        f"screen_geometry={metadata['screen_geometry']}",
                        f"device_pixel_ratio={metadata['device_pixel_ratio']}",
                        f"capture_method={metadata['capture_method']}",
                        f"digit_groups={self._blue_digit_groups(image)}",
                        f"capture_size={image.width()}x{image.height()}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            self.set_status(
                f"Saved plunger debug capture: {raw_path.name}, {filtered_path.name}; detected {value_text} nl."
            )
        except Exception as exc:
            QMessageBox.critical(self, "Plunger Debug", str(exc))

    def save_calibrate_dialog_debug(self) -> None:
        try:
            snapshot = self.controller.get_injectomate_calibrate_snapshot()
            output_dir = Path(__file__).resolve().parent
            json_path = output_dir / "injectomate_calibrate_debug.json"
            text_path = output_dir / "injectomate_calibrate_values.txt"
            json_path.write_text(json.dumps(snapshot, indent=2), encoding="utf-8")
            candidates = snapshot.get("numeric_candidates", [])
            lines = ["numeric_candidates:"]
            for candidate in candidates:
                lines.append(
                    "value={value} control_id={control_id} class={class_name} text={text!r} rect=({left},{top},{right},{bottom})".format(
                        **candidate
                    )
                )
            text_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
            self.set_status(f"Saved calibrate dialog debug: {json_path.name}, {text_path.name}.")
        except Exception as exc:
            QMessageBox.critical(self, "Calibrate Debug", str(exc))

    def read_plunger_gauge_from_screen(self) -> float | None:
        try:
            image = self._capture_plunger_gauge_image()
            if image is None:
                return None
            return self._read_plunger_text_from_image(image)
        except Exception:
            return None

    def refresh_live_position(self) -> None:
        if self.validation_move_active:
            return
        try:
            axis_position = self.controller.get_current_axis_position()
            if self.coordinate_mode == "bregma" and self.bregma_axis is not None:
                ap, ml, dv = self._axis_to_bregma(axis_position)
            else:
                ap, ml, dv = axis_position
            self.current_ap_label.setText(f"{ap:.2f}")
            self.current_ml_label.setText(f"{ml:.2f}")
            self.current_dv_label.setText(f"{dv:.2f}")
            if self.seeds or self.top_view.overlay_image is not None:
                self.redraw_views(current_point=(ml, ap))
        except Exception as exc:
            self.set_status(str(exc))

    def _set_validation_move_position(self, axis_position: object) -> None:
        if not isinstance(axis_position, tuple) or len(axis_position) != 3:
            return
        ap, ml, dv = (float(axis_position[0]), float(axis_position[1]), float(axis_position[2]))
        if self.coordinate_mode == "bregma" and self.bregma_axis is not None:
            ap, ml, dv = self._axis_to_bregma((ap, ml, dv))
        self.current_ap_label.setText(f"{ap:.2f}")
        self.current_ml_label.setText(f"{ml:.2f}")
        self.current_dv_label.setText(f"{dv:.2f}")
        self.redraw_views(current_point=(ml, ap))

    def select_overlay(self) -> None:
        path = self.overlay_combo.currentData()
        image = QImage(path) if path else None
        calibration = None
        if path:
            metadata_path = Path(path).with_suffix(".json")
            try:
                calibration = json.loads(metadata_path.read_text(encoding="utf-8")).get("calibration", {})
            except (OSError, ValueError, json.JSONDecodeError):
                calibration = None
        self.top_view.set_overlay_image(image, calibration)
        if hasattr(self, "injection_sites_view"):
            self.injection_sites_view.set_overlay_image(image, calibration)
        self.set_status("Overlay cleared." if not path else f"Overlay selected: {Path(path).name}")
        self.refresh_live_position()

    def generate_seeds(self) -> None:
        if not self._require_idle("Craniotomy"):
            return
        try:
            ap, ml, _dv = self.get_bregma_position()
            self.mid_ap.setValue(ap)
            self.mid_ml.setValue(ml)
            diameter = self.diameter.value()
            seed_count = self.seed_count.value()
            if diameter <= 0:
                raise StereoDriveError("Diameter must be greater than zero.")
            radius = diameter / 2.0
            mid_ap = self.mid_ap.value()
            mid_ml = self.mid_ml.value()
            self.seeds = []
            self.craniotomy_coordinate_system = "bregma"
            for index in range(seed_count):
                theta = 2.0 * math.pi * index / seed_count
                self.seeds.append(
                    SeedPoint(
                        index=index,
                        angle_deg=math.degrees(theta),
                        ap=mid_ap + radius * math.cos(theta),
                        ml=mid_ml + radius * math.sin(theta),
                    )
                )
            self.current_seed_index = 0
            self.trajectory = self._flat_circle_trajectory(mid_ap, mid_ml, radius)
            self.current_seed_spin.blockSignals(True)
            self.current_seed_spin.setRange(1, len(self.seeds))
            self.current_seed_spin.setValue(1)
            self.current_seed_spin.blockSignals(False)
            self.update_seed_selector_label()
            self.drill_completed_points = 0
            self.drilled_depths = [0.0] * len(self.trajectory)
            self.frozen_points = [False] * len(self.trajectory)
            self.current_target_depth_mm = self._initial_target_depth()
            self.active_surface_dv = None
            self.active_depth_ratio = None
            self.active_drill_depth_mm = None
            self.drilling_paused = False
            self.drill_pause_requested.clear()
            self.drill_stop_requested.clear()
            self.start_round_btn.setText("Start Drilling")
            if self.freeze_draw_btn.isChecked():
                self.freeze_draw_btn.setChecked(False)
            if self.unfreeze_draw_btn.isChecked():
                self.unfreeze_draw_btn.setChecked(False)
            self.update_current_target_depth_label()
            self.redraw_views()
            self.set_status(f"Generated {seed_count} seed points from current AP/ML midpoint.")
            reply = QMessageBox.question(
                self,
                "Craniotomy",
                "Move to first seed now?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if reply == QMessageBox.Yes:
                self.move_to_current_seed()
        except Exception as exc:
            QMessageBox.critical(self, "Craniotomy", str(exc))

    def set_craniotomy_center(self) -> None:
        if not self._require_idle("Craniotomy Center"):
            return
        try:
            position = self.get_bregma_position()
            self.mid_ap.setValue(position[0])
            self.mid_ml.setValue(position[1])
            self.set_status(f"Craniotomy center set to AP {position[0]:.2f}, ML {position[1]:.2f}.")
        except Exception as exc:
            QMessageBox.critical(self, "Craniotomy", str(exc))

    def set_zoom_mode(self, index: int) -> None:
        self.top_view.reset_navigation()
        self.top_view.set_zoom_level({0: 1.0, 1: 0.5, 2: 0.0}.get(index, 1.0))
        self.redraw_views()

    def set_injection_sites_zoom_mode(self, index: int) -> None:
        self.injection_sites_view.reset_navigation()
        self.injection_sites_view.set_zoom_level(0.0 if index == 2 else 1.0)
        self.redraw_views()

    def move_to_map_location(self, ml: float, ap: float) -> None:
        if not self._require_idle("Map Move"):
            return
        answer = QMessageBox.question(
            self, "Move to Map Location",
            f"Move to AP {ap:.2f}, ML {ml:.2f}? The drill will first move 0.5 mm above the reference DV.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        try:
            axis_position = self.controller.get_current_axis_position()
            if self.coordinate_mode == "bregma" and self.bregma_axis is not None:
                target = self._bregma_to_axis((ap, ml, 0.0))
                safe = self._bregma_to_axis((ap, ml, -0.5))
                safe_message = "Moving to the site 0.5 mm above Bregma DV zero. Waiting for StereoDrive to report arrival…"
            else:
                target = (ap, ml, axis_position[2])
                safe = (target[0], target[1], target[2] - 0.5)
                safe_message = "Moving to the site 0.5 mm above the current DV position. Waiting for StereoDrive to report arrival…"
            def stage_message(description: str, position: tuple[float, float, float]) -> str:
                return (
                    f"{description}\n\nWaiting for Axis AP {position[0]:.2f} mm, "
                    f"ML {position[1]:.2f} mm, DV {position[2]:.2f} mm."
                )

            path = self._axis_clearance_path(target, safe[2])
            if not self._move_through_axis_positions_with_progress(
                path,
                title="Moving to Map Location",
                message=stage_message(safe_message, safe),
                stage_messages=[
                    "Retracting vertically before lateral travel.",
                    stage_message(safe_message, safe),
                    stage_message("Moving to the selected map location.", target),
                ],
            ):
                return
            self.set_status(f"Moved to map location AP {ap:.2f}, ML {ml:.2f}.")
        except Exception as exc:
            QMessageBox.critical(self, "Map Move", str(exc))

    def current_ap_label_value(self) -> float:
        return float(self.current_ap_label.text())

    def current_ml_label_value(self) -> float:
        return float(self.current_ml_label.text())

    def _flat_circle_trajectory(self, mid_ap: float, mid_ml: float, radius: float) -> list[tuple[float, float, float]]:
        point_count = self.trajectory_points.value()
        return [
            (
                mid_ap + radius * math.cos(2.0 * math.pi * idx / point_count),
                mid_ml + radius * math.sin(2.0 * math.pi * idx / point_count),
                0.0,
            )
            for idx in range(point_count + 1)
        ]

    def _initial_target_depth(self) -> float:
        return min(max(self.depth_per_round.value(), 0.0), self.drill_depth.value())

    def update_current_target_depth_label(self) -> None:
        self.current_target_depth_label.setText(f"Current Target Depth: {self.current_target_depth_mm:.3f} mm")

    def change_current_target_depth(self) -> None:
        max_depth = max(0.0, self.drill_depth.value())
        dialog = QDialog(self)
        dialog.setWindowTitle("Current Target Depth")
        layout = QGridLayout(dialog)
        layout.addWidget(QLabel("Current target depth (mm)"), 0, 0)
        depth_box = QLineEdit(f"{self.current_target_depth_mm:.3f}")
        layout.addWidget(depth_box, 0, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons, 1, 0, 1, 2)
        if dialog.exec() != QDialog.Accepted:
            return
        try:
            value = float(depth_box.text())
        except ValueError:
            QMessageBox.warning(self, "Current Target Depth", "Enter a numeric depth in mm.")
            return
        self.current_target_depth_mm = min(max(value, 0.0), max_depth)
        self.update_current_target_depth_label()
        self.redraw_views()

    def update_seed_selector_label(self) -> None:
        if self.current_seed_index is None or not self.seeds:
            self.current_seed_coords.setText("Seed: -")
            return
        seed = self.seeds[self.current_seed_index]
        state = "surface set" if seed.dv is not None else "surface pending"
        self.current_seed_coords.setText(
            f"Seed {seed.index + 1} [{seed.ap:.2f}, {seed.ml:.2f}] {state}"
        )

    def on_seed_spin_changed(self, value: int) -> None:
        if not self.seeds:
            self.current_seed_index = None
            self.update_seed_selector_label()
            return
        self.current_seed_index = max(0, min(len(self.seeds) - 1, value - 1))
        self.update_seed_selector_label()

    def move_to_current_seed(self) -> None:
        if self.current_seed_index is None or not self.seeds:
            QMessageBox.information(self, "Craniotomy", "Generate seed points first.")
            return
        try:
            self._require_project_coordinates("craniotomy")
            seed = self.seeds[self.current_seed_index]
            if not self._move_to_axis_position_with_progress(
                self._bregma_to_axis((seed.ap, seed.ml, -1.0)),
                delay_seconds=1.0,
                title=f"Moving to Seed {seed.index + 1}",
                message="Moving to the seed position. Waiting for StereoDrive to report arrival…",
            ):
                return
            self.set_status(
                f"Moved to seed {seed.index + 1} target [{seed.ap:.2f}, {seed.ml:.2f}, -1.00]. Lower manually to the skull surface, then click 'Set Surface'."
            )
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def capture_surface(self) -> None:
        if not self._require_idle("Capture Surface"):
            return
        if self.current_seed_index is None or not self.seeds:
            QMessageBox.information(self, "Craniotomy", "Generate seed points first.")
            return
        try:
            self._require_project_coordinates("craniotomy")
            if QApplication.keyboardModifiers() & Qt.KeyboardModifier.ControlModifier:
                for seed in self.seeds:
                    seed.dv = 0.0
                    seed.sampled_ap = seed.ap
                    seed.sampled_ml = seed.ml
                self.current_seed_index = None
                self.compute_trajectory()
                self.update_seed_selector_label()
                self.redraw_views()
                self.set_status("Debug: set all seed surfaces to DV 0.00 and updated the trajectory.")
                return
            ap, ml, dv = self.get_bregma_position()
            seed = self.seeds[self.current_seed_index]
            seed.dv = dv
            seed.sampled_ap = ap
            seed.sampled_ml = ml
            self.compute_trajectory()
            next_pending = next((s.index for s in self.seeds if s.dv is None), None)
            self.current_seed_index = next_pending
            if next_pending is not None:
                self.current_seed_spin.blockSignals(True)
                self.current_seed_spin.setValue(next_pending + 1)
                self.current_seed_spin.blockSignals(False)
            self.update_seed_selector_label()
            self.redraw_views()
            if next_pending is None:
                self.set_status("Captured all seed points and updated the trajectory.")
            else:
                self.move_to_current_seed()
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def clear_surface_measurements(self) -> None:
        if not self._require_idle("Clear Surfaces"):
            return
        for seed in self.seeds:
            seed.dv = None
            seed.sampled_ap = None
            seed.sampled_ml = None
        self.trajectory = []
        self.drilled_depths = []
        self.frozen_points = []
        self.drill_completed_points = 0
        self.drill_round_started_at = None
        self.drill_round_target_seconds = 0.0
        self.active_surface_dv = None
        self.active_depth_ratio = None
        self.active_drill_depth_mm = None
        self.current_target_depth_mm = self._initial_target_depth()
        self.drilling_paused = False
        self.start_round_btn.setText("Start Drilling")
        if self.freeze_draw_btn.isChecked():
            self.freeze_draw_btn.setChecked(False)
        if self.unfreeze_draw_btn.isChecked():
            self.unfreeze_draw_btn.setChecked(False)
        self.current_seed_index = 0 if self.seeds else None
        if self.seeds:
            self.current_seed_spin.blockSignals(True)
            self.current_seed_spin.setValue(1)
            self.current_seed_spin.blockSignals(False)
        self.update_seed_selector_label()
        self.update_current_target_depth_label()
        self.redraw_views()
        self.set_status("Cleared all captured surface measurements.")

    def clear_craniotomy(self) -> None:
        """Discard the active craniotomy plan without changing injection or calibration data."""
        if not self._require_idle("Clear Craniotomy"):
            return
        if not self.seeds and not self.trajectory:
            self.set_status("There is no active craniotomy to clear.")
            return
        response = QMessageBox.warning(
            self,
            "Clear Craniotomy?",
            "This removes the current craniotomy seeds, surface measurements, trajectory, drill progress, and frozen points. "
            "It does not change injection sites, Bregma/anchor calibration, or the craniotomy setup values.",
            QMessageBox.Yes | QMessageBox.Cancel,
            QMessageBox.Cancel,
        )
        if response != QMessageBox.Yes:
            return
        self.seeds.clear()
        self.craniotomy_coordinate_system = "bregma"
        self.trajectory.clear()
        self.drilled_depths.clear()
        self.frozen_points.clear()
        self.current_seed_index = None
        self.current_seed_spin.blockSignals(True)
        self.current_seed_spin.setRange(1, 1)
        self.current_seed_spin.setValue(1)
        self.current_seed_spin.blockSignals(False)
        self.drill_completed_points = 0
        self.drill_round_started_at = None
        self.drill_round_target_seconds = 0.0
        self.active_surface_dv = None
        self.active_depth_ratio = None
        self.active_drill_depth_mm = None
        self.current_target_depth_mm = self._initial_target_depth()
        self.drilling_paused = False
        self.drill_pause_requested.clear()
        self.drill_stop_requested.clear()
        self.start_round_btn.setText("Start Drilling")
        if self.freeze_draw_btn.isChecked():
            self.freeze_draw_btn.setChecked(False)
        if self.unfreeze_draw_btn.isChecked():
            self.unfreeze_draw_btn.setChecked(False)
        self.update_seed_selector_label()
        self.update_current_target_depth_label()
        self.redraw_views()
        self.set_status("Cleared the current craniotomy.")

    def clear_project(self) -> None:
        """Discard all recoverable project state while keeping application preferences."""
        if self._motion_is_active():
            QMessageBox.warning(self, "Clear Project", "Stop and finish the current operation before clearing the project.")
            return
        response = QMessageBox.warning(
            self,
            "Clear Project?",
            "This clears the current craniotomy, injection sites, Bregma/anchor calibration, and stored A/B/C locations. "
            "Keyboard shortcuts and reusable injection/craniotomy settings are kept.",
            QMessageBox.Yes | QMessageBox.Cancel,
            QMessageBox.Cancel,
        )
        if response != QMessageBox.Yes:
            return
        self.seeds.clear()
        self.trajectory.clear()
        self.drilled_depths.clear()
        self.frozen_points.clear()
        self.current_seed_index = None
        self.current_seed_spin.blockSignals(True)
        self.current_seed_spin.setRange(1, 1)
        self.current_seed_spin.setValue(1)
        self.current_seed_spin.blockSignals(False)
        self.drill_completed_points = 0
        self.drill_round_started_at = None
        self.drill_round_target_seconds = 0.0
        self.active_surface_dv = None
        self.active_depth_ratio = None
        self.active_drill_depth_mm = None
        self.current_target_depth_mm = self._initial_target_depth()
        self.drilling_paused = False
        self.drill_pause_requested.clear()
        self.drill_stop_requested.clear()
        self.start_round_btn.setText("Start Drilling")
        self.injection_sites.clear()
        if self.nudge_all_sites_active:
            self.nudge_all_sites_btn.setChecked(False)
        self.bregma_axis = None
        self.anchor_axis = None
        self.anchor_bregma = None
        self.coordinate_mode = "axis"
        self.craniotomy_coordinate_system = "bregma"
        self.injection_sites_coordinate_system = "bregma"
        self.quick_locations.clear()
        if self.freeze_draw_btn.isChecked():
            self.freeze_draw_btn.setChecked(False)
        if self.unfreeze_draw_btn.isChecked():
            self.unfreeze_draw_btn.setChecked(False)
        self.update_coordinate_mode_buttons()
        self.top_view.set_coordinate_mode_bregma(False)
        self.injection_sites_view.set_coordinate_mode_bregma(False)
        self.update_seed_selector_label()
        self.update_current_target_depth_label()
        self.refresh_injection_sites_list()
        self.redraw_views()
        try:
            self._project_session_path().unlink(missing_ok=True)
        except OSError:
            pass
        self._save_general_settings()
        self.set_status("Cleared the current project and its recovery session.")

    def compute_trajectory(self) -> None:
        if not self._require_idle("Compute Trajectory"):
            return
        captured = [seed for seed in self.seeds if seed.dv is not None]
        if len(captured) < 2:
            radius = self.diameter.value() / 2.0
            self.trajectory = self._flat_circle_trajectory(self.mid_ap.value(), self.mid_ml.value(), radius)
            if len(self.drilled_depths) != len(self.trajectory):
                self.drilled_depths = [0.0] * len(self.trajectory)
            if len(self.frozen_points) != len(self.trajectory):
                self.frozen_points = [False] * len(self.trajectory)
            self.redraw_views()
            return
        known = sorted(
            ((2.0 * math.pi * seed.index / len(self.seeds), seed.dv) for seed in self.seeds if seed.dv is not None),
            key=lambda item: item[0],
        )
        angles = [item[0] for item in known]
        values = [item[1] for item in known]
        radius = self.diameter.value() / 2.0
        self.trajectory = []
        for idx in range(self.trajectory_points.value() + 1):
            theta = 2.0 * math.pi * idx / self.trajectory_points.value()
            ap = self.mid_ap.value() + radius * math.cos(theta)
            ml = self.mid_ml.value() + radius * math.sin(theta)
            dv = self.interpolate_periodic(theta, angles, values) + self.cut_offset.value()
            self.trajectory.append((ap, ml, dv))
        self.drilled_depths = [0.0] * len(self.trajectory)
        self.frozen_points = [False] * len(self.trajectory)
        self.redraw_views()

    def toggle_freeze_mode(self, enabled: bool) -> None:
        if enabled and self.unfreeze_draw_btn.isChecked():
            self.unfreeze_draw_btn.setChecked(False)
        self.top_view.set_freeze_mode(enabled)
        if enabled:
            self.set_status("Draw on the circle to freeze points from deeper drilling.")
        elif self.trajectory:
            self.set_status("Freeze drawing off.")

    def toggle_unfreeze_mode(self, enabled: bool) -> None:
        if enabled and self.freeze_draw_btn.isChecked():
            self.freeze_draw_btn.setChecked(False)
        self.top_view.set_unfreeze_mode(enabled)
        if enabled:
            self.set_status("Draw on the circle to remove frozen points.")
        elif self.trajectory:
            self.set_status("Unfreeze drawing off.")

    def clear_frozen_points(self) -> None:
        if not self.frozen_points:
            return
        self.frozen_points = [False] * len(self.frozen_points)
        if self.freeze_draw_btn.isChecked():
            self.freeze_draw_btn.setChecked(False)
        self.redraw_views()
        self.set_status("Cleared all frozen trajectory points.")

    def mark_frozen_point(self, index: int) -> None:
        if not self.trajectory or index < 0 or index >= len(self.trajectory):
            return
        if not self.frozen_points:
            self.frozen_points = [False] * len(self.trajectory)
        if not self.frozen_points[index]:
            self.frozen_points[index] = True
            self.redraw_views()

    def unmark_frozen_point(self, index: int) -> None:
        if not self.frozen_points or index < 0 or index >= len(self.frozen_points):
            return
        if self.frozen_points[index]:
            self.frozen_points[index] = False
            self.redraw_views()

    def interpolate_periodic(self, theta: float, angles: list[float], values: list[float]) -> float:
        theta = theta % (2.0 * math.pi)
        extended_angles = angles + [angles[0] + 2.0 * math.pi]
        extended_values = values + [values[0]]
        for idx in range(len(extended_angles) - 1):
            a0 = extended_angles[idx]
            a1 = extended_angles[idx + 1]
            if a0 <= theta <= a1:
                t = 0.0 if math.isclose(a0, a1) else (theta - a0) / (a1 - a0)
                return extended_values[idx] + t * (extended_values[idx + 1] - extended_values[idx])
        return values[-1]

    def stop_motion(self) -> None:
        self.drill_stop_requested.set()
        self.injection_stop_requested.set()
        if self.validation_move_cancel_callback is not None:
            self.validation_move_cancel_callback()
        try:
            self.controller.stop()
            self.set_status("Sent Stop command.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))
        finally:
            if self.injection_thread is not None and self.injection_thread.is_alive():
                try:
                    self.controller.stop_injectomate_motion()
                except Exception as exc:
                    self.set_status(f"Syringe stop could not be confirmed: {exc}")

    def set_drill_completed_points(self, completed_points: int) -> None:
        self.drill_completed_points = completed_points
        self.redraw_views()

    def start_drilling_round(self) -> None:
        if self.drill_thread is not None and self.drill_thread.is_alive():
            self.pause_drilling_round()
            return
        if not self._require_idle("Drilling"):
            return
        try:
            self._require_project_coordinates("craniotomy")
        except Exception as exc:
            QMessageBox.warning(self, "Drilling", str(exc))
            return
        if not self.trajectory or len(self.seeds) < 2 or any(seed.dv is None for seed in self.seeds):
            QMessageBox.information(self, "Craniotomy", "Capture all seed surfaces before drilling.")
            return
        max_depth = max(0.0, self.drill_depth.value())
        if self.current_target_depth_mm <= 0.0:
            self.current_target_depth_mm = self._initial_target_depth()
        depth = min(self.current_target_depth_mm, max_depth)
        self.current_target_depth_mm = depth
        self.update_current_target_depth_label()
        if depth <= 0.0:
            QMessageBox.information(self, "Craniotomy", "Current target depth must be greater than zero.")
            return
        self.controller.prepare_motion()
        if self.nudge_all_sites_active:
            self.nudge_all_sites_btn.setChecked(False)
        self.drill_pause_requested.clear()
        self.drill_stop_requested.clear()
        self.drilling_paused = False
        self.drill_completed_points = 0
        current_depths = list(self.drilled_depths or [0.0] * len(self.trajectory))
        frozen_points = list(self.frozen_points or [False] * len(self.trajectory))
        if not any((not frozen) and current_depth + 0.0005 < depth for current_depth, frozen in zip(current_depths, frozen_points, strict=False)):
            self.on_drill_round_finished("completed")
            return
        round_time_seconds = self.round_time_seconds.value()
        self.drill_round_started_at = time.monotonic()
        self.drill_round_target_seconds = round_time_seconds
        self.active_drill_depth_mm = depth
        self.active_depth_ratio = 0.0
        self.start_round_btn.setText("Pause")
        surface_targets = [self._bregma_to_axis(point) for point in self.trajectory]
        self.drill_clearance_axis_dv = self.bregma_axis[2] - 0.5
        target_depths = [
            current_depth if frozen else max(current_depth, depth)
            for current_depth, frozen in zip(current_depths, frozen_points, strict=False)
        ]
        if surface_targets:
            self.active_surface_dv = surface_targets[0][2]
        center_above_position = self._bregma_to_axis((
            self.mid_ap.value(),
            self.mid_ml.value(),
            self._craniotomy_center_surface_dv() - 2.0,
        ))
        self.drill_thread = threading.Thread(
            target=self._run_drilling_round,
            args=(
                surface_targets,
                current_depths,
                target_depths,
                frozen_points,
                round_time_seconds,
                depth,
                center_above_position,
            ),
            daemon=True,
        )
        self.drill_thread.start()

    def pause_drilling_round(self) -> None:
        self.drill_pause_requested.set()
        self.set_status("Pausing drilling and retracting to 2 mm above surface.")

    def on_drill_round_finished(self, outcome: str) -> None:
        self.drill_round_started_at = None
        self.drill_round_target_seconds = 0.0
        self.drill_completed_points = 0
        if outcome == "paused":
            self.drilling_paused = True
            self.start_round_btn.setText("Continue")
            self.redraw_views()
            return
        if outcome in ("stopped", "error"):
            self.drilling_paused = False
            self.start_round_btn.setText("Start Drilling")
            self.redraw_views()
            return
        max_depth = max(0.0, self.drill_depth.value())
        completed_depth = self.current_target_depth_mm
        if completed_depth >= max_depth - 0.0005:
            self.current_target_depth_mm = max_depth
            self.drilling_paused = False
            self.start_round_btn.setText("Start Drilling")
            self.update_current_target_depth_label()
            self.set_status("Drilling complete to max depth.")
            self.redraw_views()
            return
        self.current_target_depth_mm = min(max_depth, completed_depth + self.depth_per_round.value())
        self.update_current_target_depth_label()
        if self.auto_start_rounds.isChecked():
            should_continue = self.show_next_round_countdown(self.current_target_depth_mm)
            if should_continue:
                self.drilling_paused = False
                self.start_round_btn.setText("Pause")
                self.set_status(f"Starting next drilling round to {self.current_target_depth_mm:.3f} mm.")
                QTimer.singleShot(100, self.start_drilling_round)
            else:
                self.drilling_paused = True
                self.start_round_btn.setText("Continue")
                self.set_status(f"Paused before next round to {self.current_target_depth_mm:.3f} mm.")
        else:
            self.drilling_paused = True
            self.start_round_btn.setText("Continue")
            self.set_status(f"Round complete. Next target depth is {self.current_target_depth_mm:.3f} mm.")
        self.redraw_views()

    def _craniotomy_center_surface_dv(self) -> float:
        captured_dvs = [seed.dv for seed in self.seeds if seed.dv is not None]
        if captured_dvs:
            return (sum(captured_dvs) / len(captured_dvs)) + self.cut_offset.value()
        if self.active_surface_dv is not None:
            return self.active_surface_dv - self.bregma_axis[2]
        return 0.0

    def show_next_round_countdown(self, target_depth_mm: float) -> bool:
        QApplication.beep()
        dialog = QDialog(self)
        dialog.setWindowTitle("Next Drilling Round")
        dialog.setModal(True)

        layout = QVBoxLayout(dialog)
        label = QLabel()
        progress = QProgressBar()
        progress.setRange(0, 1000)
        progress.setValue(0)
        change_button = QPushButton("Change next target depth")
        pause_button = QPushButton("Pause")
        layout.addWidget(label)
        layout.addWidget(progress)
        layout.addWidget(change_button)
        layout.addWidget(pause_button)

        duration_s = 10.0
        started_at = time.monotonic()
        paused = {"value": False}
        target_depth = {"value": target_depth_mm}

        def pause_countdown() -> None:
            paused["value"] = True
            timer.stop()
            dialog.reject()

        def update_countdown() -> None:
            elapsed_s = max(0.0, time.monotonic() - started_at)
            remaining_s = max(0.0, duration_s - elapsed_s)
            label.setText(
                f"Will start a new round to drill to {target_depth['value']:.3f} mm in {remaining_s:0.1f} s."
            )
            progress.setValue(int((elapsed_s / duration_s) * 1000))
            if remaining_s <= 0.0:
                timer.stop()
                dialog.accept()

        def change_target_depth() -> None:
            edit_dialog = QDialog(dialog)
            edit_dialog.setWindowTitle("Next Target Depth")
            edit_layout = QGridLayout(edit_dialog)
            edit_layout.addWidget(QLabel("Next target depth (mm)"), 0, 0)
            edit_box = QLineEdit(f"{target_depth['value']:.3f}")
            edit_layout.addWidget(edit_box, 0, 1)
            buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            buttons.accepted.connect(edit_dialog.accept)
            buttons.rejected.connect(edit_dialog.reject)
            edit_layout.addWidget(buttons, 1, 0, 1, 2)
            if edit_dialog.exec() != QDialog.Accepted:
                return
            try:
                value = float(edit_box.text())
            except ValueError:
                QMessageBox.warning(dialog, "Next Target Depth", "Enter a numeric depth in mm.")
                return
            value = max(0.0, min(value, max(0.0, self.drill_depth.value())))
            target_depth["value"] = value
            self.current_target_depth_mm = value
            self.update_current_target_depth_label()
            update_countdown()

        timer = QTimer(dialog)
        timer.timeout.connect(update_countdown)
        change_button.clicked.connect(change_target_depth)
        pause_button.clicked.connect(pause_countdown)
        update_countdown()
        timer.start(100)
        accepted = dialog.exec() == QDialog.Accepted
        timer.stop()
        return accepted and not paused["value"]

    def stop_drilling_round(self) -> None:
        self.drill_stop_requested.set()
        try:
            self.controller.stop()
        except Exception:
            pass
        self.set_status("Stopping drilling round.")

    def _should_abort_drilling(self) -> bool:
        return self.drill_pause_requested.is_set() or self.drill_stop_requested.is_set()

    def _continuous_round_substep_count(
        self,
        surface_targets: list[tuple[float, float, float]],
        frozen_points: list[bool],
        traversal_order: list[int],
        path_step_mm: float,
    ) -> int:
        total = 0
        previous_index: int | None = None
        for index in traversal_order:
            if frozen_points[index]:
                previous_index = None
                continue
            if previous_index is not None and index != previous_index:
                start_ap, start_ml, _start_dv = surface_targets[previous_index]
                end_ap, end_ml, _end_dv = surface_targets[index]
                total += max(1, int(math.ceil(math.hypot(end_ap - start_ap, end_ml - start_ml) / path_step_mm)))
            previous_index = index
        return max(1, total)

    def _continuous_round_axis_travel_mm(
        self,
        surface_targets: list[tuple[float, float, float]],
        frozen_points: list[bool],
        traversal_order: list[int],
    ) -> float:
        total = 0.0
        previous_index: int | None = None
        for index in traversal_order:
            if frozen_points[index]:
                previous_index = None
                continue
            if previous_index is not None and index != previous_index:
                start_ap, start_ml, _start_dv = surface_targets[previous_index]
                end_ap, end_ml, _end_dv = surface_targets[index]
                total += abs(end_ap - start_ap) + abs(end_ml - start_ml)
            previous_index = index
        return total

    def _continuous_round_path_step_mm(
        self,
        surface_targets: list[tuple[float, float, float]],
        frozen_points: list[bool],
        traversal_order: list[int],
        round_time_seconds: float,
    ) -> float:
        axis_travel_mm = self._continuous_round_axis_travel_mm(surface_targets, frozen_points, traversal_order)
        # Benchmark-derived typical time for one AP/ML nudge including UI/readback overhead.
        seconds_per_axis_nudge = 0.85
        required_step = axis_travel_mm * seconds_per_axis_nudge / max(round_time_seconds, 1.0)
        for candidate in (0.05, 0.1, 0.2, 0.5, 1.0):
            if required_step <= candidate:
                return candidate
        return 1.0

    def _mark_continuous_round_point(self, index: int, target_depth: float, point_count: int) -> None:
        if index < len(self.drilled_depths):
            self.drilled_depths[index] = target_depth
        if index == 0 and point_count < len(self.drilled_depths):
            self.drilled_depths[point_count] = target_depth
        self.active_depth_ratio = max(
            0.0,
            min(1.0, target_depth / max(self.skull_thickness_mm.value(), 0.001)),
        )
        self.redraw_signal.emit()
        self.drill_progress_signal.emit(index + 1)

    def _follow_continuous_round_segment(
        self,
        surface_targets: list[tuple[float, float, float]],
        start_index: int,
        end_index: int,
        target_depth: float,
        path_step_mm: float,
        planar_tolerance_mm: float,
        dv_tolerance_mm: float,
        completed_substeps: int,
        total_substeps: int,
        round_started_at: float,
        round_time_seconds: float,
        point_count: int,
    ) -> int:
        start_ap, start_ml, start_surface_dv = surface_targets[start_index]
        end_ap, end_ml, end_surface_dv = surface_targets[end_index]
        distance = math.hypot(end_ap - start_ap, end_ml - start_ml)
        substeps = max(1, int(math.ceil(distance / path_step_mm)))
        for substep in range(1, substeps + 1):
            if self._should_abort_drilling():
                break
            fraction = substep / substeps
            ap = start_ap + (end_ap - start_ap) * fraction
            ml = start_ml + (end_ml - start_ml) * fraction
            surface_dv = start_surface_dv + (end_surface_dv - start_surface_dv) * fraction
            target_dv = surface_dv + target_depth
            self.active_surface_dv = surface_dv
            self.status_signal.emit(
                f"Continuous cutting at {int(((end_index % point_count) / max(1, point_count)) * 100)}%"
            )
            self.controller.move_to_position_nudged(
                ap,
                ml,
                target_dv,
                step_mm=path_step_mm,
                planar_tolerance=planar_tolerance_mm,
                dv_tolerance=dv_tolerance_mm,
                stop_requested=self._should_abort_drilling,
                status_callback=None,
                dwell_seconds=0.0,
            )
            completed_substeps += 1
            target_elapsed = round_time_seconds * (completed_substeps / total_substeps)
            while not self._should_abort_drilling():
                remaining = target_elapsed - (time.monotonic() - round_started_at)
                if remaining <= 0.0:
                    break
                time.sleep(min(0.05, remaining))
        return completed_substeps

    def _run_drilling_round(
        self,
        surface_targets: list[tuple[float, float, float]],
        current_depths: list[float],
        target_depths: list[float],
        frozen_points: list[bool],
        round_time_seconds: float,
        depth_mm: float,
        center_above_position: tuple[float, float, float],
    ) -> None:
        point_count = max(1, len(surface_targets) - 1)
        outcome = "completed"
        try:
            if point_count <= 0:
                return
            needs_drilling = [
                (not frozen_points[index]) and current_depths[index] + 0.0005 < target_depths[index]
                for index in range(point_count)
            ]
            if not any(needs_drilling):
                return
            first_index = next(index for index, needed in enumerate(needs_drilling) if needed)
            traversal_order = [(first_index + offset) % point_count for offset in range(point_count + 1)]
            path_step_mm = self._continuous_round_path_step_mm(
                surface_targets,
                frozen_points,
                traversal_order,
                round_time_seconds,
            )
            planar_tolerance_mm = max(0.015, min(0.08, path_step_mm * 0.25))
            dv_tolerance_mm = max(0.02, min(0.05, path_step_mm * 0.25))
            self.status_signal.emit(
                f"Continuous tracing step {path_step_mm:.3f} mm, AP/ML tolerance {planar_tolerance_mm:.3f} mm."
            )
            total_substeps = self._continuous_round_substep_count(
                surface_targets,
                frozen_points,
                traversal_order,
                path_step_mm,
            )
            completed_substeps = 0
            round_started_at = time.monotonic()
            at_cutting_depth = False
            previous_index: int | None = None

            for order_position, index in enumerate(traversal_order):
                if self._should_abort_drilling():
                    break
                ap, ml, surface_dv = surface_targets[index]
                current_depth = current_depths[index]
                target_depth = target_depths[index]
                if frozen_points[index]:
                    if at_cutting_depth and self.active_surface_dv is not None:
                        self.status_signal.emit("Frozen section: retracting before crossing gap.")
                        self.controller.move_axis_to_target(
                            "DV",
                            self.active_surface_dv - 2.0,
                            step_mm=5.0,
                            stop_requested=self._should_abort_drilling,
                            status_callback=None,
                            dwell_seconds=0.005,
                        )
                    self.drill_progress_signal.emit(index + 1)
                    at_cutting_depth = False
                    previous_index = None
                    continue
                if not at_cutting_depth:
                    current_dv_target = surface_dv + current_depth
                    target_dv = surface_dv + target_depth
                    self.active_surface_dv = surface_dv
                    self.active_depth_ratio = max(
                        0.0,
                        min(1.0, current_depth / max(self.skull_thickness_mm.value(), 0.001)),
                    )
                    self.redraw_signal.emit()
                    self.status_signal.emit(
                        f"Moving to start continuous cut at {int(order_position / max(1, point_count) * 100)}%"
                    )
                    self._approach_axis_position((ap, ml, current_dv_target),
                        min(surface_dv - 2.0, self.drill_clearance_axis_dv), self._should_abort_drilling)
                    self.controller.wait_for_axis_position(
                        ap,
                        ml,
                        current_dv_target,
                        tolerance_mm=0.02,
                        timeout_seconds=60.0,
                        poll_seconds=0.1,
                        stop_requested=self._should_abort_drilling,
                    )
                    self.status_signal.emit(f"Entering continuous cut to DV {target_dv:.2f}")
                    dwell_seconds = 0.005 / max(self.drill_rate_mm_per_s.value(), 0.001)
                    self.controller.move_axis_to_target(
                        "DV",
                        target_dv,
                        step_mm=0.005,
                        stop_requested=self._should_abort_drilling,
                        status_callback=None,
                        dwell_seconds=dwell_seconds,
                    )
                    at_cutting_depth = True
                    previous_index = index
                    self._mark_continuous_round_point(index, target_depths[index], point_count)
                    continue
                if previous_index is None:
                    previous_index = index
                    continue
                if index == previous_index:
                    continue
                completed_substeps = self._follow_continuous_round_segment(
                    surface_targets,
                    previous_index,
                    index,
                    target_depth,
                    path_step_mm,
                    planar_tolerance_mm,
                    dv_tolerance_mm,
                    completed_substeps,
                    total_substeps,
                    round_started_at,
                    round_time_seconds,
                    point_count,
                )
                self._mark_continuous_round_point(index, target_depth, point_count)
                previous_index = index
            if not self._should_abort_drilling():
                center_ap, center_ml, center_dv = center_above_position
                self.status_signal.emit("Round complete. Returning above craniotomy center.")
                self._approach_axis_position((center_ap, center_ml, center_dv),
                                             self.drill_clearance_axis_dv, self._should_abort_drilling)
                self.controller.wait_for_axis_position(
                    center_ap,
                    center_ml,
                    center_dv,
                    tolerance_mm=0.02,
                    timeout_seconds=60.0,
                    poll_seconds=0.1,
                    stop_requested=self._should_abort_drilling,
                )
                if not self._should_abort_drilling():
                    self.status_signal.emit("Drilling round complete. Returned above craniotomy center.")
        except Exception as exc:
            if not self._should_abort_drilling():
                outcome = "error"
                self.status_signal.emit(str(exc))
        finally:
            if outcome == "error" or self.drill_stop_requested.is_set():
                try:
                    self.controller.stop()
                    self.controller.wait_until_stopped()
                except Exception as stop_exc:
                    outcome = "error"
                    self.status_signal.emit(f"Stop confirmation failed: {stop_exc}")
            if self.drill_pause_requested.is_set() and not self.drill_stop_requested.is_set():
                outcome = "paused"
                try:
                    retract_dv = (self.active_surface_dv - 2.0) if self.active_surface_dv is not None else -2.0
                    self.controller.move_axis_to_target(
                        "DV",
                        retract_dv,
                        step_mm=5.0,
                        stop_requested=None,
                        status_callback=None,
                        dwell_seconds=0.005,
                    )
                    self.status_signal.emit("Drilling paused 2 mm above surface.")
                except Exception as retract_exc:
                    outcome = "error"
                    try:
                        self.controller.stop()
                        self.controller.wait_until_stopped()
                    except Exception as stop_exc:
                        self.status_signal.emit(f"Stop confirmation failed: {stop_exc}")
                    self.status_signal.emit(str(retract_exc))
                self.drill_pause_requested.clear()
            elif self.drill_stop_requested.is_set():
                if outcome != "error":
                    outcome = "stopped"
                    self.status_signal.emit("Drilling round stopped; stable Axis position confirmed.")
                self.drill_stop_requested.clear()
            self.drill_round_started_at = None
            self.drill_round_target_seconds = 0.0
            self.active_surface_dv = None
            self.active_depth_ratio = None
            self.active_drill_depth_mm = None
            self.drill_thread = None
            self.redraw_signal.emit()
            self.drill_round_finished_signal.emit(outcome)

    def redraw_views(self, current_point: tuple[float, float] | None = None) -> None:
        top_points: list[tuple[float, float, float]] = []
        skull_thickness_mm = max(self.skull_thickness_mm.value(), 0.001)
        current_depth_ratio = self.active_depth_ratio
        map_trajectory = self.trajectory if self.craniotomy_coordinate_system == "bregma" else []
        for index, (ap, ml, _dv) in enumerate(map_trajectory):
            depth_mm = self.drilled_depths[index] if index < len(self.drilled_depths) else 0.0
            point_depth_ratio = max(0.0, min(1.0, depth_mm / skull_thickness_mm))
            display_ap, display_ml = self._project_map_position(ap, ml, self.craniotomy_coordinate_system)
            top_points.append((display_ml, display_ap, point_depth_ratio))
        top_seeds = []
        for seed in self.seeds if self.craniotomy_coordinate_system == "bregma" else []:
            display_ap, display_ml = self._project_map_position(seed.ap, seed.ml, self.craniotomy_coordinate_system)
            top_seeds.append((display_ml, display_ap, seed.dv is not None))
        if current_point is None and (self.seeds or self.injection_sites or self.top_view.overlay_image is not None):
            try:
                current_ap, current_ml, current_dv = self.controller.get_current_axis_position()
                if self.coordinate_mode == "bregma":
                    current_ap, current_ml, current_dv = self._axis_to_bregma((current_ap, current_ml, current_dv))
                current_point = (current_ml, current_ap)
            except Exception:
                current_point = None
        if self.drill_thread is not None and self.drill_thread.is_alive() and self.active_surface_dv is not None:
            try:
                _current_ap, _current_ml, current_dv = self.controller.get_current_axis_position()
                current_depth_mm = max(0.0, current_dv - self.active_surface_dv)
                computed_ratio = max(0.0, min(1.0, current_depth_mm / skull_thickness_mm))
                if self.active_depth_ratio is None or abs(computed_ratio - self.active_depth_ratio) >= 0.0005:
                    self.active_depth_ratio = computed_ratio
                current_depth_ratio = self.active_depth_ratio
            except Exception:
                current_depth_ratio = self.active_depth_ratio
        self.depth_legend.set_skull_thickness_mm(self.skull_thickness_mm.value())
        self.depth_legend.set_current_depth_ratio(current_depth_ratio)
        self._update_round_status_labels()
        self.top_view.set_data(
            top_points,
            top_seeds,
            frozen_points=self.frozen_points,
            current_point=current_point,
            anchor_point=(self.anchor_bregma[1], self.anchor_bregma[0])
            if self.coordinate_mode == "bregma" and self.anchor_bregma is not None else None,
        )
        show_craniotomy = self.show_craniotomy_on_injection_map.isChecked()
        injection_site_points = []
        for site in self.injection_sites if self.injection_sites_coordinate_system == "bregma" else []:
            display_ap, display_ml = self._project_map_position(site.ap, site.ml, self.injection_sites_coordinate_system)
            injection_site_points.append((display_ml, display_ap))
        focus_points = [(point[0], point[1]) for point in top_points]
        if self.injection_sites_zoom_combo.currentIndex() == 1:
            focus_points = injection_site_points
        elif self.injection_sites_zoom_combo.currentIndex() == 2:
            focus_points = []
        self.injection_sites_view.set_data(
            top_points if show_craniotomy else [],
            [],
            frozen_points=self.frozen_points,
            current_point=current_point,
            injection_sites=injection_site_points,
            anchor_point=(self.anchor_bregma[1], self.anchor_bregma[0])
            if self.coordinate_mode == "bregma" and self.anchor_bregma is not None else None,
        )
        self.injection_sites_view.set_view_focus_points(focus_points or None)
        self.update_seed_selector_label()

    def _project_map_position(self, ap: float, ml: float, frame: str) -> tuple[float, float]:
        if self.coordinate_mode == "axis" and self.bregma_axis is not None and frame == "bregma":
            return self.bregma_axis[0] + ap, self.bregma_axis[1] + ml
        return ap, ml

    def _format_duration(self, seconds: float) -> str:
        total_seconds = max(0, int(round(seconds)))
        minutes, secs = divmod(total_seconds, 60)
        hours, minutes = divmod(minutes, 60)
        if hours > 0:
            return f"{hours:d}:{minutes:02d}:{secs:02d}"
        return f"{minutes:02d}:{secs:02d}"

    def _update_round_status_labels(self) -> None:
        if self.drill_round_started_at is None or self.drill_round_target_seconds <= 0:
            self.round_elapsed_label.setText("Elapsed: --:--")
            self.round_remaining_label.setText("Remaining: --:--")
            self.round_percent_label.setText("Complete: --%")
            return
        elapsed = max(0.0, time.monotonic() - self.drill_round_started_at)
        remaining = max(0.0, self.drill_round_target_seconds - elapsed)
        total_points = max(1, len(self.trajectory))
        percent = min(100.0, (self.drill_completed_points / total_points) * 100.0)
        self.round_elapsed_label.setText(f"Elapsed: {self._format_duration(elapsed)}")
        self.round_remaining_label.setText(f"Remaining: {self._format_duration(remaining)}")
        self.round_percent_label.setText(f"Complete: {percent:.0f}%")


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Craniotomy Planner")
    application_font = QFont("Segoe UI")
    application_font.setPointSize(9)
    app.setFont(application_font)
    try:
        window = CraniotomyWindow()
    except StereoDriveError:
        QMessageBox.warning(
            None,
            "StereoDrive Not Found",
            "StereoDrive main window was not found.\n\n"
            "Open StereoDrive first, then start Craniotomy Planner again.",
        )
        return
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
