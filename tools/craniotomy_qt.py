import math
import ctypes
import json
import sys
import threading
import time
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QEvent, QPoint, QPointF, QRectF, QSize, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QFont, QImage, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import (
    QApplication,
    QAbstractSpinBox,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSpinBox,
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
    dv: float


@dataclass
class StoredLocation:
    ap: float
    ml: float
    dv: float


@dataclass
class InjectionProtocolSettings:
    main_volume_nl: int
    insertion_rate_nl_min: float
    main_rate_nl_min: float
    injection_depth_mm: float
    insert_retract_speed_um_s: float
    overshoot_mm: float
    post_inject_pause_s: float


class ProjectionWidget(QWidget):
    freeze_drawn = Signal(int)
    unfreeze_drawn = Signal(int)

    def __init__(self, title: str, x_label: str, y_label: str, invert_y: bool = False, parent: QWidget | None = None):
        super().__init__(parent)
        self.title = title
        self.x_label = x_label
        self.y_label = y_label
        self.invert_y = invert_y
        self.trajectory: list[tuple[float, float, float]] = []
        self.seed_points: list[tuple[float, float, bool]] = []
        self.frozen_points: list[bool] = []
        self.current_point: tuple[float, float] | None = None
        self.freeze_mode = False
        self.unfreeze_mode = False
        self._trajectory_screen_points: list[QPointF] = []
        self._inner_ring_screen_points: list[QPointF] = []
        self.setMinimumHeight(360)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

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
    ) -> None:
        self.trajectory = trajectory
        self.seed_points = seed_points
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
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        if self.freeze_mode and event.buttons() & Qt.LeftButton:
            self._emit_nearest_trajectory_index(event.position(), freeze=True)
            event.accept()
            return
        if self.unfreeze_mode and event.buttons() & Qt.LeftButton:
            self._emit_nearest_trajectory_index(event.position(), freeze=False)
            event.accept()
            return
        super().mouseMoveEvent(event)

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
        available_rect = self.rect().adjusted(pad, pad + 20, -pad, -pad)
        side = min(available_rect.width(), available_rect.height())
        draw_rect = QRectF(
            available_rect.left() + (available_rect.width() - side) / 2.0,
            available_rect.top() + (available_rect.height() - side) / 2.0,
            side,
            side,
        )

        painter.setPen(QPen(QColor("#cad7cb"), 1))
        painter.setBrush(QColor("#ffffff"))
        painter.drawRoundedRect(draw_rect, 16, 16)

        painter.setPen(QColor("#496052"))
        painter.drawText(16, 20, self.title)

        if not self.trajectory and not self.seed_points:
            painter.setPen(QColor("#8b9a8d"))
            painter.drawText(self.rect(), Qt.AlignCenter, "No trajectory yet")
            return

        xs = [p[0] for p in self.trajectory] + [s[0] for s in self.seed_points]
        ys = [p[1] for p in self.trajectory] + [s[1] for s in self.seed_points]
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
        uniform_span = max(span_x, span_y) * 1.3
        cx = (min_x + max_x) / 2.0
        cy = (min_y + max_y) / 2.0
        min_x = cx - uniform_span / 2.0
        max_x = cx + uniform_span / 2.0
        min_y = cy - uniform_span / 2.0
        max_y = cy + uniform_span / 2.0

        def map_point(x: float, y: float) -> QPointF:
            px = draw_rect.left() + (x - min_x) / (max_x - min_x) * draw_rect.width()
            normalized_y = (y - min_y) / (max_y - min_y)
            if self.invert_y:
                py = draw_rect.top() + normalized_y * draw_rect.height()
            else:
                py = draw_rect.bottom() - normalized_y * draw_rect.height()
            return QPointF(px, py)

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

        if self.current_point is not None:
            pt = map_point(self.current_point[0], self.current_point[1])
            marker_pen = QPen(QColor("#1f2937"), 4)
            painter.setPen(marker_pen)
            painter.drawLine(pt + QPointF(-10, -10), pt + QPointF(10, 10))
            painter.drawLine(pt + QPointF(-10, 10), pt + QPointF(10, -10))

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
    injection_finished_signal = Signal(str)
    syringe_position_signal = Signal(object)
    syringe_limit_warning_signal = Signal(str)
    block_prompt_signal = Signal()

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
        self.drill_pause_requested = threading.Event()
        self.drill_stop_requested = threading.Event()
        self.drill_thread: threading.Thread | None = None
        self.drill_completed_points = 0
        self.drill_round_started_at: float | None = None
        self.drill_round_target_seconds: float = 0.0
        self.active_surface_dv: float | None = None
        self.active_depth_ratio: float | None = None
        self.active_drill_depth_mm: float | None = None
        self.manual_injection_volume_nl = DEFAULT_INJECTION_VOLUME_NL
        self.syringe_position_nl: float | None = None
        self.syringe_position_lock = threading.Lock()
        self.injection_sites: list[InjectionSite] = []
        self.quick_locations: dict[str, StoredLocation] = {}
        self.injection_thread: threading.Thread | None = None
        self.injection_pause_requested = threading.Event()
        self.injection_stop_requested = threading.Event()
        self.block_prompt_event: threading.Event | None = None
        self.block_prompt_continue = True
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
        self.injection_finished_signal.connect(self.finish_injection)
        self.syringe_position_signal.connect(self.set_syringe_position)
        self.syringe_limit_warning_signal.connect(self.show_syringe_limit_warning)
        self.block_prompt_signal.connect(self.show_block_prompt)
        self._build_ui()
        QApplication.instance().installEventFilter(self)
        self.refresh_live_position()
        QTimer.singleShot(250, self.update_syringe_position_from_scale)
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.refresh_live_position)
        self.refresh_timer.start(50)
        self._start_warning_auto_confirm_watcher()

    def _start_warning_auto_confirm_watcher(self) -> None:
        def worker() -> None:
            while not self.warning_auto_confirm_stop.is_set():
                try:
                    self.controller.confirm_below_skull_warning(timeout_seconds=0.0)
                except Exception:
                    pass
                self.warning_auto_confirm_stop.wait(0.05)

        threading.Thread(target=worker, daemon=True).start()

    def closeEvent(self, event) -> None:  # noqa: N802
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
            QDoubleSpinBox, QSpinBox {
                border: 1px solid #cfdbcf;
                border-radius: 6px;
                padding: 1px 3px;
                background: white;
                min-height: 14px;
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
        home_btn = QPushButton("Home")
        home_btn.clicked.connect(self.goto_home)
        work_btn = QPushButton("Work")
        work_btn.clicked.connect(self.goto_work)
        header_stop_btn = QPushButton("Stop")
        header_stop_btn.setProperty("variant", "danger")
        header_stop_btn.style().unpolish(header_stop_btn)
        header_stop_btn.style().polish(header_stop_btn)
        header_stop_btn.clicked.connect(self.stop_motion)
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
            goto_btn = QPushButton(name)
            goto_btn.setProperty("variant", variant)
            goto_btn.clicked.connect(lambda _checked=False, slot=name: self.goto_quick_location(slot))
            for button in (set_btn, goto_btn):
                button.style().unpolish(button)
                button.style().polish(button)
                quick_buttons.append(button)
        position_layout.addWidget(QLabel("AP"))
        position_layout.addWidget(self.current_ap_label)
        position_layout.addWidget(QLabel("ML"))
        position_layout.addWidget(self.current_ml_label)
        position_layout.addWidget(QLabel("DV"))
        position_layout.addWidget(self.current_dv_label)
        position_layout.addWidget(header_stop_btn)
        position_layout.addStretch(1)
        header_layout.addWidget(set_bregma_btn)
        header_layout.addWidget(home_btn)
        header_layout.addWidget(work_btn)
        for button in quick_buttons:
            header_layout.addWidget(button)
        header_layout.addStretch(1)
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

        setup_box = QGroupBox("Setup")
        setup_layout = QGridLayout(setup_box)
        setup_layout.setContentsMargins(7, 6, 7, 7)
        setup_layout.setHorizontalSpacing(8)
        setup_layout.setVerticalSpacing(3)
        content.addWidget(setup_box)

        self.mid_ap = self._double_spinbox()
        self.mid_ml = self._double_spinbox()
        self.diameter = self._double_spinbox(value=3.0, minimum=0.1, maximum=20.0)
        self.seed_count = self._spinbox(value=6, minimum=3, maximum=24)
        self.trajectory_points = self._spinbox(value=60, minimum=12, maximum=360)
        self.cut_offset = self._double_spinbox(value=0.0, minimum=-5.0, maximum=5.0)
        self.drill_depth = self._double_spinbox(value=0.10, minimum=0.0, maximum=5.0)
        self.skull_thickness_mm = self._double_spinbox(value=0.25, minimum=0.001, maximum=5.0)
        self.round_time_seconds = self._double_spinbox(value=60.0, minimum=1.0, maximum=3600.0)
        self.drill_rate_mm_per_s = self._double_spinbox(value=0.01, minimum=0.001, maximum=5.0)
        self.current_seed_spin = self._spinbox(value=1, minimum=1, maximum=1)
        self.current_seed_spin.valueChanged.connect(self.on_seed_spin_changed)
        self.current_seed_coords = QLabel("Seed: -")
        self.current_seed_coords.setWordWrap(True)
        self.current_seed_coords.setProperty("role", "muted")

        setup_layout.addWidget(QLabel("Mid AP"), 0, 0)
        setup_layout.addWidget(self.mid_ap, 0, 1)
        setup_layout.addWidget(QLabel("Mid ML"), 0, 2)
        setup_layout.addWidget(self.mid_ml, 0, 3)

        setup_layout.addWidget(QLabel("Diameter (mm)"), 1, 0)
        setup_layout.addWidget(self.diameter, 1, 1)
        setup_layout.addWidget(QLabel("Seed Points"), 1, 2)
        setup_layout.addWidget(self.seed_count, 1, 3)
        setup_layout.addWidget(QLabel("Trajectory Points"), 1, 4)
        setup_layout.addWidget(self.trajectory_points, 1, 5)

        setup_layout.addWidget(QLabel("Cut Offset DV"), 2, 0)
        setup_layout.addWidget(self.cut_offset, 2, 1)
        setup_layout.addWidget(QLabel("Drill Depth"), 2, 2)
        setup_layout.addWidget(self.drill_depth, 2, 3)
        setup_layout.addWidget(QLabel("Skull Thickness (mm)"), 2, 4)
        setup_layout.addWidget(self.skull_thickness_mm, 2, 5)

        setup_layout.addWidget(QLabel("Current Seed"), 3, 0)
        setup_layout.addWidget(self.current_seed_spin, 3, 1)
        setup_layout.addWidget(QLabel("Drill Rate (mm/s)"), 3, 2)
        setup_layout.addWidget(self.drill_rate_mm_per_s, 3, 3)
        setup_layout.addWidget(QLabel("Round Time (s)"), 3, 4)
        setup_layout.addWidget(self.round_time_seconds, 3, 5)

        generate_btn = QPushButton("Generate Seeds")
        generate_btn.clicked.connect(self.generate_seeds)
        self.move_seed_btn = QPushButton("Move To Current Seed")
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
        self.start_round_btn = QPushButton("Start Drilling Round")
        self.start_round_btn.setProperty("variant", "primary")
        self.start_round_btn.style().unpolish(self.start_round_btn)
        self.start_round_btn.style().polish(self.start_round_btn)
        self.start_round_btn.clicked.connect(self.start_drilling_round)
        self.pause_round_btn = QPushButton("Pause Round")
        self.pause_round_btn.clicked.connect(self.pause_drilling_round)
        self.stop_drill_btn = QPushButton("Stop Drilling")
        self.stop_drill_btn.setProperty("variant", "danger")
        self.stop_drill_btn.style().unpolish(self.stop_drill_btn)
        self.stop_drill_btn.style().polish(self.stop_drill_btn)
        self.stop_drill_btn.clicked.connect(self.stop_drilling_round)
        self.freeze_draw_btn = QPushButton("Draw Freeze")
        self.freeze_draw_btn.setCheckable(True)
        self.freeze_draw_btn.toggled.connect(self.toggle_freeze_mode)
        self.unfreeze_draw_btn = QPushButton("Draw Unfreeze")
        self.unfreeze_draw_btn.setCheckable(True)
        self.unfreeze_draw_btn.toggled.connect(self.toggle_unfreeze_mode)
        self.clear_freeze_btn = QPushButton("Clear Freeze")
        self.clear_freeze_btn.clicked.connect(self.clear_frozen_points)
        setup_layout.addWidget(self.current_seed_coords, 4, 0, 1, 6)

        button_layout = QGridLayout()
        button_layout.setHorizontalSpacing(6)
        button_layout.setVerticalSpacing(3)
        button_layout.addWidget(generate_btn, 0, 0)
        button_layout.addWidget(clear_btn, 0, 1)
        button_layout.addWidget(stop_btn, 0, 2)
        button_layout.addWidget(self.move_seed_btn, 1, 0)
        button_layout.addWidget(self.capture_surface_btn, 1, 1)
        button_layout.addWidget(self.start_round_btn, 1, 2)
        button_layout.addWidget(self.freeze_draw_btn, 2, 0)
        button_layout.addWidget(self.clear_freeze_btn, 2, 1)
        button_layout.addWidget(self.unfreeze_draw_btn, 2, 2)
        button_layout.addWidget(self.pause_round_btn, 3, 0)
        button_layout.addWidget(self.stop_drill_btn, 3, 1, 1, 2)
        setup_layout.addLayout(button_layout, 5, 0, 1, 6)

        views_box = QGroupBox()
        views_layout = QGridLayout(views_box)
        views_layout.setContentsMargins(7, 6, 7, 7)
        content.addWidget(views_box, 1)

        views_header = QHBoxLayout()
        views_header.setSpacing(5)
        views_title = QLabel("Trajectory View")
        views_title.setProperty("role", "muted")
        self.move_speed_label = QLabel()
        views_header.addWidget(views_title)
        views_header.addWidget(self.move_speed_label)
        views_header.addStretch(1)
        views_layout.addLayout(views_header, 0, 0, 1, 2)

        self.top_view = ProjectionWidget(self.current_action, "ML", "AP")
        self.top_view.freeze_drawn.connect(self.mark_frozen_point)
        self.top_view.unfreeze_drawn.connect(self.unmark_frozen_point)
        self.top_view.setMinimumSize(420, 420)
        self.top_view.setMaximumWidth(620)
        views_layout.addWidget(self.top_view, 1, 0)
        legend_layout = QVBoxLayout()
        legend_layout.setSpacing(3)
        self.depth_legend = DepthLegendWidget()
        self.depth_legend.set_skull_thickness_mm(self.skull_thickness_mm.value())
        legend_layout.addWidget(self.depth_legend, 0, Qt.AlignTop)
        self.round_elapsed_label = QLabel("Elapsed: --:--")
        self.round_remaining_label = QLabel("Remaining: --:--")
        self.round_percent_label = QLabel("Complete: --%")
        self.round_elapsed_label.setProperty("role", "muted")
        self.round_remaining_label.setProperty("role", "muted")
        self.round_percent_label.setProperty("role", "muted")
        legend_layout.addWidget(self.round_elapsed_label)
        legend_layout.addWidget(self.round_remaining_label)
        legend_layout.addWidget(self.round_percent_label)
        legend_layout.addStretch(1)
        views_layout.addLayout(legend_layout, 1, 1)
        self.update_move_speed_label()
        self._build_injection_tab()

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
        single_layout.addWidget(QLabel("Program sequence"), 3, 0, 1, 6)
        single_layout.addWidget(self.sequence_steps_list, 4, 0, 1, 6)
        single_layout.addWidget(QLabel("Overall sequence progress"), 5, 0)
        single_layout.addWidget(self.injection_progress, 5, 1, 1, 5)
        single_layout.addWidget(QLabel("Current injection/movement"), 6, 0)
        single_layout.addWidget(self.injection_site_progress, 6, 1, 1, 5)
        single_layout.addWidget(self.start_injection_btn, 7, 0)
        single_layout.addWidget(self.pause_injection_btn, 7, 1)
        single_layout.addWidget(self.stop_injection_btn, 7, 2, 1, 4)

        sites_box = QGroupBox("Injection Sites")
        sites_layout = QGridLayout(sites_box)
        sites_layout.setContentsMargins(7, 6, 7, 7)
        sites_layout.setHorizontalSpacing(8)
        sites_layout.setVerticalSpacing(3)
        layout.addWidget(sites_box)

        self.injection_sites_list = QListWidget()
        add_site_btn = QPushButton("Add Injection Site")
        add_site_btn.clicked.connect(self.add_injection_site)
        remove_site_btn = QPushButton("Remove Selected Site")
        remove_site_btn.clicked.connect(self.remove_selected_injection_site)
        clear_sites_btn = QPushButton("Clear Sites")
        clear_sites_btn.clicked.connect(self.clear_injection_sites)
        self.block_check = QCheckBox("Check blockage after each site")
        self.block_check.toggled.connect(self.refresh_injection_sequence_summary)
        sites_layout.addWidget(add_site_btn, 0, 0)
        sites_layout.addWidget(remove_site_btn, 0, 1)
        sites_layout.addWidget(clear_sites_btn, 0, 2)
        sites_layout.addWidget(self.block_check, 0, 3, 1, 2)
        sites_layout.addWidget(self.injection_sites_list, 1, 0, 1, 5)
        layout.addStretch(1)
        self.update_manual_volume_label()
        self.update_injection_rate_label()
        self.refresh_injection_sequence_summary()

    def _double_spinbox(self, value: float = 0.0, minimum: float = -100.0, maximum: float = 100.0) -> QDoubleSpinBox:
        widget = QDoubleSpinBox()
        widget.setDecimals(2)
        widget.setRange(minimum, maximum)
        widget.setValue(value)
        return widget

    def _spinbox(self, value: int, minimum: int, maximum: int) -> QSpinBox:
        widget = QSpinBox()
        widget.setRange(minimum, maximum)
        widget.setValue(value)
        return widget

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

    def set_status(self, message: str) -> None:
        self.current_action = message or "Trajectory"
        if hasattr(self, "action_status_label"):
            self.action_status_label.setText(self.current_action)
        self.top_view.title = self.current_action
        self.top_view.update()

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
        if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter) and self._focus_is_editable():
            focus_widget = QApplication.focusWidget()
            if isinstance(focus_widget, QAbstractSpinBox):
                focus_widget.interpretText()
            if focus_widget is not None:
                focus_widget.clearFocus()
            self.setFocus(Qt.OtherFocusReason)
            return True
        if key == Qt.Key.Key_F1:
            self.adjust_manual_injection_volume(-1)
            return True
        if key == Qt.Key.Key_F2:
            self.adjust_manual_injection_volume(1)
            return True
        if key == Qt.Key.Key_F3:
            self.manual_syringe_step(up=True)
            return True
        if key == Qt.Key.Key_F4:
            self.manual_syringe_step(up=False)
            return True
        if key == Qt.Key.Key_Escape:
            self.stop_injection()
            return True
        if self._focus_is_editable():
            return super().eventFilter(watched, event)
        if key == Qt.Key.Key_Shift:
            self.adjust_move_speed(1)
            return True
        if key == Qt.Key.Key_Control:
            self.adjust_move_speed(-1)
            return True
        key_map = {
            Qt.Key.Key_Left: ("ML", False, "ML left"),
            Qt.Key.Key_Right: ("ML", True, "ML right"),
            Qt.Key.Key_Up: ("AP", True, "AP anterior"),
            Qt.Key.Key_Down: ("AP", False, "AP posterior"),
            Qt.Key.Key_PageUp: ("DV", False, "DV up"),
            Qt.Key.Key_PageDown: ("DV", True, "DV down"),
        }
        if key not in key_map:
            return super().eventFilter(watched, event)
        axis, positive, label = key_map[key]
        self.keyboard_nudge(axis, positive, label)
        return True

    def _focus_is_editable(self) -> bool:
        focus_widget = QApplication.focusWidget()
        return isinstance(focus_widget, (QLineEdit, QAbstractSpinBox, QComboBox))

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
        if self.drill_thread is not None and self.drill_thread.is_alive():
            return
        if self._focus_is_editable():
            return
        try:
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
        if self.injection_thread is not None and self.injection_thread.is_alive():
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
        if self.injection_thread is not None and self.injection_thread.is_alive():
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
        if self.injection_thread is not None and self.injection_thread.is_alive():
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
        if self.injection_thread is not None and self.injection_thread.is_alive():
            return
        try:
            self.controller.empty_syringe()
            self.track_syringe_empty()
            self.set_status("Emptying syringe to 0")
        except Exception as exc:
            QMessageBox.critical(self, "Injectomate", str(exc))

    def set_current_location_to_bregma(self) -> None:
        try:
            self.controller.set_current_location_to_bregma()
            self.refresh_live_position()
            self.set_status("Current location set to Bregma. AP/ML/DV verified at 0.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def goto_home(self) -> None:
        try:
            self.controller.goto_home()
            self.set_status("Sent GoTo 'Home' command.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def goto_work(self) -> None:
        try:
            self.controller.goto_work()
            self.set_status("Sent GoTo 'Work' command.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def set_quick_location(self, slot: str) -> None:
        try:
            ap, ml, dv = self.controller.get_current_position()
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
            self.controller.goto_position(location.ap, location.ml, location.dv, delay_seconds=0.5)
            self.set_status(
                f"Moving to location {slot}: AP {location.ap:.2f}, ML {location.ml:.2f}, DV {location.dv:.2f}."
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
        self.sequence_steps_list.clear()
        steps = [
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
            steps.append("Run the blockage test from 1.000 mm above the stored surface.")
        for index, text in enumerate(steps, start=1):
            item = QListWidgetItem(f"{index}. {text}")
            font = item.font()
            font.setBold(False)
            item.setFont(font)
            item.setSizeHint(QSize(0, 15))
            self.sequence_steps_list.addItem(item)

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
        try:
            ap, ml, dv = self.controller.get_current_position()
            self.injection_sites.append(InjectionSite(ap=ap, ml=ml, dv=dv))
            self.refresh_injection_sites_list()
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def remove_selected_injection_site(self) -> None:
        row = self.injection_sites_list.currentRow()
        if 0 <= row < len(self.injection_sites):
            del self.injection_sites[row]
            self.refresh_injection_sites_list()

    def clear_injection_sites(self) -> None:
        self.injection_sites.clear()
        self.refresh_injection_sites_list()

    def refresh_injection_sites_list(self) -> None:
        self.injection_sites_list.clear()
        for index, site in enumerate(self.injection_sites, start=1):
            self.injection_sites_list.addItem(
                f"{index}. AP {site.ap:.2f}, ML {site.ml:.2f}, surface DV {site.dv:.2f}"
            )

    def _active_injection_sites(self) -> list[InjectionSite]:
        if self.injection_sites:
            return list(self.injection_sites)
        ap, ml, dv = self.controller.get_current_position()
        return [InjectionSite(ap=ap, ml=ml, dv=dv)]

    def start_single_injection(self) -> None:
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
            required_volume_nl = settings.main_volume_nl * len(sites)
            test_volume_nl = self._rounded_test_volume()
            if self.block_check.isChecked():
                required_volume_nl += test_volume_nl * len(sites)
            self.ensure_total_syringe_capacity(required_volume_nl)
            self.injection_pause_requested.clear()
            self.injection_stop_requested.clear()
            self.injection_progress.setValue(0)
            self.injection_site_progress.setValue(0)
            self.set_status(
                f"Ready: {len(sites)} site(s), {settings.main_volume_nl} nl; insertion {settings.insertion_rate_nl_min:.1f}, main {settings.main_rate_nl_min:.1f} nl/min"
            )
            self.start_injection_btn.setEnabled(False)
            self.pause_injection_btn.setText("Pause")
            self.injection_thread = threading.Thread(
                target=self._run_injection_protocol,
                args=(
                    sites,
                    settings,
                    injection_plan,
                    self.block_check.isChecked(),
                    test_volume_nl,
                ),
                daemon=True,
            )
            self.injection_thread.start()
        except Exception as exc:
            QMessageBox.warning(self, "Injection", str(exc))

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
    ) -> None:
        try:
            total_units = max(1, len(sites))
            step_indexes = self._sequence_step_indexes(settings, check_blocked)
            for site_index, site in enumerate(sites, start=1):
                if self.injection_stop_requested.is_set():
                    break
                self.sequence_step_signal.emit(step_indexes["approach"])
                self.injection_progress_signal.emit(
                    int(((site_index - 1) / total_units) * 100),
                    f"Moving to 1 mm above surface for site {site_index}/{len(sites)}",
                )
                above_dv = self._above_surface_dv(site)
                self.controller.goto_position(site.ap, site.ml, above_dv, delay_seconds=0.5)
                self.controller.wait_for_position(
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
                    len(sites),
                )
                if check_blocked and not self.injection_stop_requested.is_set():
                    self.sequence_step_signal.emit(step_indexes["block"])
                    self._run_block_test(site, settings, test_volume_nl)
            if self.injection_stop_requested.is_set():
                self.sequence_step_signal.emit(-1)
                self.injection_finished_signal.emit("Injection stopped")
            else:
                self.sequence_step_signal.emit(-1)
                self.injection_finished_signal.emit("Injection protocol complete")
        except Exception as exc:
            if self.injection_stop_requested.is_set():
                self.sequence_step_signal.emit(-1)
                self.injection_finished_signal.emit("Injection stopped")
                return
            message = str(exc)
            if "Maximum movement" in message or "limit" in message:
                self.syringe_limit_warning_signal.emit(message)
            self.sequence_step_signal.emit(-1)
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
        self.controller.goto_position(site.ap, site.ml, site.dv, delay_seconds=0.5)
        self.controller.wait_for_position(
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
        delivered = 0
        event_index = 0
        start_time = time.monotonic()
        last_move_at = 0.0
        while not self.injection_stop_requested.is_set():
            start_time += self._wait_while_injection_paused()
            elapsed = time.monotonic() - start_time
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
                target_dv = self._interpolated_movement_dv(movement_targets, elapsed)
                self.controller.move_axis_to_target(
                    "DV",
                    target_dv,
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
        self.controller.goto_position(site.ap, site.ml, above_dv, delay_seconds=0.5)
        self.controller.wait_for_position(
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
        self.controller.goto_position(site.ap, site.ml, above_dv, delay_seconds=0.5)
        self.controller.wait_for_position(
            site.ap,
            site.ml,
            above_dv,
            tolerance_mm=0.02,
            timeout_seconds=60.0,
            poll_seconds=0.1,
            stop_requested=self.injection_stop_requested.is_set,
        )
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
        self.block_prompt_continue = True
        self.block_prompt_signal.emit()
        self.block_prompt_event.wait(timeout=6.0)
        if not self.block_prompt_continue:
            self.injection_pause_requested.set()
            self.injection_progress_signal.emit(100, "Paused after blockage test")

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
            font.setBold(index == row)
            item.setFont(font)
        if 0 <= row < self.sequence_steps_list.count():
            self.sequence_steps_list.setCurrentRow(row)
            self.sequence_steps_list.scrollToItem(self.sequence_steps_list.item(row))
        else:
            self.sequence_steps_list.clearSelection()
            self.sequence_steps_list.setCurrentRow(-1)

    def finish_injection(self, message: str) -> None:
        self.set_active_sequence_step(-1)
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
        box = QMessageBox(self)
        box.setWindowTitle("Blockage Check")
        box.setText("Test injection complete. Continue protocol?")
        continue_button = box.addButton("Continue", QMessageBox.AcceptRole)
        pause_button = box.addButton("Pause", QMessageBox.RejectRole)
        box.setDefaultButton(continue_button)
        QTimer.singleShot(5000, continue_button.click)
        box.exec()
        self.block_prompt_continue = box.clickedButton() != pause_button
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
        try:
            ap, ml, dv = self.controller.get_current_position()
            self.current_ap_label.setText(f"{ap:.2f}")
            self.current_ml_label.setText(f"{ml:.2f}")
            self.current_dv_label.setText(f"{dv:.2f}")
            if self.seeds:
                self.redraw_views(current_point=(ml, ap))
        except Exception as exc:
            self.set_status(str(exc))

    def generate_seeds(self) -> None:
        try:
            ap, ml, _dv = self.controller.get_current_position()
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
            self.trajectory = []
            self.current_seed_spin.blockSignals(True)
            self.current_seed_spin.setRange(1, len(self.seeds))
            self.current_seed_spin.setValue(1)
            self.current_seed_spin.blockSignals(False)
            self.update_seed_selector_label()
            self.drill_completed_points = 0
            self.drilled_depths = []
            self.frozen_points = []
            self.active_surface_dv = None
            self.active_depth_ratio = None
            self.active_drill_depth_mm = None
            if self.freeze_draw_btn.isChecked():
                self.freeze_draw_btn.setChecked(False)
            if self.unfreeze_draw_btn.isChecked():
                self.unfreeze_draw_btn.setChecked(False)
            self.redraw_views()
            self.set_status(f"Generated {seed_count} seed points from current AP/ML midpoint.")
        except Exception as exc:
            QMessageBox.critical(self, "Craniotomy", str(exc))

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
            seed = self.seeds[self.current_seed_index]
            self.controller.goto_position(seed.ap, seed.ml, -1.0, delay_seconds=1.0)
            self.set_status(
                f"Moved to seed {seed.index + 1} target [{seed.ap:.2f}, {seed.ml:.2f}, -1.00]. Lower manually to the skull surface, then click 'Set Surface'."
            )
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def capture_surface(self) -> None:
        if self.current_seed_index is None or not self.seeds:
            QMessageBox.information(self, "Craniotomy", "Generate seed points first.")
            return
        try:
            ap, ml, dv = self.controller.get_current_position()
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
        self.redraw_views()
        self.set_status("Cleared all captured surface measurements.")

    def compute_trajectory(self) -> None:
        captured = [seed for seed in self.seeds if seed.dv is not None]
        if len(captured) < 2:
            self.trajectory = []
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
        try:
            self.drill_stop_requested.set()
            self.controller.stop()
            self.set_status("Sent Stop command.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def set_drill_completed_points(self, completed_points: int) -> None:
        self.drill_completed_points = completed_points
        self.redraw_views()

    def start_drilling_round(self) -> None:
        if self.drill_thread is not None and self.drill_thread.is_alive():
            return
        if not self.trajectory or any(seed.dv is None for seed in self.seeds):
            QMessageBox.information(self, "Craniotomy", "Capture all seed surfaces before drilling.")
            return
        self.drill_pause_requested.clear()
        self.drill_stop_requested.clear()
        self.drill_completed_points = 0
        depth = self.drill_depth.value()
        current_depths = list(self.drilled_depths or [0.0] * len(self.trajectory))
        frozen_points = list(self.frozen_points or [False] * len(self.trajectory))
        if not any((not frozen) and current_depth + 0.0005 < depth for current_depth, frozen in zip(current_depths, frozen_points, strict=False)):
            QMessageBox.information(
                self,
                "Craniotomy",
                "Requested depth has already been drilled in all unfrozen section",
            )
            return
        round_time_seconds = self.round_time_seconds.value()
        self.drill_round_started_at = time.monotonic()
        self.drill_round_target_seconds = round_time_seconds
        self.active_drill_depth_mm = depth
        self.active_depth_ratio = 0.0
        surface_targets = list(self.trajectory)
        target_depths = [
            current_depth if frozen else max(current_depth, depth)
            for current_depth, frozen in zip(current_depths, frozen_points, strict=False)
        ]
        if surface_targets:
            self.active_surface_dv = surface_targets[0][2]
        self.drill_thread = threading.Thread(
            target=self._run_drilling_round,
            args=(surface_targets, current_depths, target_depths, frozen_points, round_time_seconds, depth),
            daemon=True,
        )
        self.drill_thread.start()

    def pause_drilling_round(self) -> None:
        self.drill_pause_requested.set()
        self.set_status("Pausing drilling round and retracting to DV -2.00.")

    def stop_drilling_round(self) -> None:
        self.drill_stop_requested.set()
        try:
            self.controller.stop()
        except Exception:
            pass
        self.set_status("Stopping drilling round.")

    def _should_abort_drilling(self) -> bool:
        return self.drill_pause_requested.is_set() or self.drill_stop_requested.is_set()

    def _run_drilling_round(
        self,
        surface_targets: list[tuple[float, float, float]],
        current_depths: list[float],
        target_depths: list[float],
        frozen_points: list[bool],
        round_time_seconds: float,
        depth_mm: float,
    ) -> None:
        point_count = len(surface_targets)
        per_point_budget = round_time_seconds / max(1, point_count)
        in_air = False
        active_drill_section = False
        try:
            for index, (ap, ml, surface_dv) in enumerate(surface_targets):
                if self._should_abort_drilling():
                    break
                current_depth = current_depths[index]
                target_depth = target_depths[index]
                current_dv_target = surface_dv + current_depth
                target_dv = surface_dv + target_depth
                self.active_surface_dv = surface_dv
                self.active_depth_ratio = max(
                    0.0,
                    min(1.0, current_depth / max(self.skull_thickness_mm.value(), 0.001)),
                )
                self.redraw_signal.emit()
                self.status_signal.emit(f"Moving to {int(index / max(1, point_count) * 100)}%")
                started_at = time.monotonic()
                if frozen_points[index]:
                    if not in_air:
                        self.status_signal.emit("Frozen section: retracting to DV 1.00")
                        self.controller.move_axis_to_target(
                            "DV",
                            1.0,
                            step_mm=5.0,
                            stop_requested=self._should_abort_drilling,
                            status_callback=None,
                            dwell_seconds=0.005,
                        )
                        in_air = True
                        active_drill_section = False
                    self.controller.move_to_position_nudged(
                        ap,
                        ml,
                        1.0,
                        step_mm=5.0,
                        stop_requested=self._should_abort_drilling,
                        status_callback=None,
                        dwell_seconds=0.005,
                    )
                else:
                    if index == 0 or in_air or (not active_drill_section):
                        self.controller.goto_position(ap, ml, current_dv_target, delay_seconds=0.5)
                        self.controller.wait_for_position(
                            ap,
                            ml,
                            current_dv_target,
                            tolerance_mm=0.02,
                            timeout_seconds=60.0,
                            poll_seconds=0.1,
                            stop_requested=self._should_abort_drilling,
                        )
                        in_air = False
                        active_drill_section = True
                    else:
                        self.controller.move_to_position_nudged(
                            ap,
                            ml,
                            target_dv,
                            step_mm=5.0,
                            stop_requested=self._should_abort_drilling,
                            status_callback=None,
                            dwell_seconds=0.005,
                        )
                    if target_depth > current_depth + 0.0005:
                        self.status_signal.emit(
                            f"Drilling down at {int(index / max(1, point_count) * 100)}% to DV {target_dv:.2f}"
                        )
                        dwell_seconds = 0.005 / max(self.drill_rate_mm_per_s.value(), 0.001)
                        self.controller.move_axis_to_target(
                            "DV",
                            target_dv,
                            step_mm=0.005,
                            stop_requested=self._should_abort_drilling,
                            status_callback=None,
                            dwell_seconds=dwell_seconds,
                        )
                        current_depths[index] = target_depth
                        self.drilled_depths[index] = target_depth
                        self.active_depth_ratio = max(
                            0.0,
                            min(1.0, target_depth / max(self.skull_thickness_mm.value(), 0.001)),
                        )
                        self.redraw_signal.emit()
                self.drill_progress_signal.emit(index + 1)
                remaining = per_point_budget - (time.monotonic() - started_at)
                while remaining > 0 and not self._should_abort_drilling():
                    sleep_window = min(0.05, remaining)
                    time.sleep(sleep_window)
                    remaining -= sleep_window
            if not self._should_abort_drilling():
                midpoint_ap = self.mid_ap.value()
                midpoint_ml = self.mid_ml.value()
                self.status_signal.emit("Drilling round complete. Returning to midpoint.")
                self.controller.goto_position(midpoint_ap, midpoint_ml, -1.0, delay_seconds=0.5)
                self.controller.wait_for_position(
                    midpoint_ap,
                    midpoint_ml,
                    -1.0,
                    tolerance_mm=0.02,
                    timeout_seconds=60.0,
                    poll_seconds=0.1,
                    stop_requested=self._should_abort_drilling,
                )
                if not self._should_abort_drilling():
                    self.status_signal.emit("Drilling round complete. Returned to midpoint.")
        except Exception as exc:
            if not self._should_abort_drilling():
                self.status_signal.emit(str(exc))
        finally:
            if self.drill_pause_requested.is_set():
                try:
                    self.controller.move_axis_to_target(
                        "DV",
                        -2.0,
                        step_mm=5.0,
                        stop_requested=None,
                        status_callback=None,
                        dwell_seconds=0.005,
                    )
                    self.status_signal.emit("Round paused at DV -2.00.")
                except Exception as retract_exc:
                    self.status_signal.emit(str(retract_exc))
                self.drill_pause_requested.clear()
            elif self.drill_stop_requested.is_set():
                self.status_signal.emit("Drilling round stopped.")
                self.drill_stop_requested.clear()
            self.drill_round_started_at = None
            self.drill_round_target_seconds = 0.0
            self.active_surface_dv = None
            self.active_depth_ratio = None
            self.active_drill_depth_mm = None
            self.drill_thread = None
            self.redraw_signal.emit()

    def redraw_views(self, current_point: tuple[float, float] | None = None) -> None:
        top_points: list[tuple[float, float, float]] = []
        skull_thickness_mm = max(self.skull_thickness_mm.value(), 0.001)
        current_depth_ratio = self.active_depth_ratio
        for index, (ap, ml, _dv) in enumerate(self.trajectory):
            depth_mm = self.drilled_depths[index] if index < len(self.drilled_depths) else 0.0
            point_depth_ratio = max(0.0, min(1.0, depth_mm / skull_thickness_mm))
            top_points.append((ml, ap, point_depth_ratio))
        top_seeds = [(seed.ml, seed.ap, seed.dv is not None) for seed in self.seeds]
        if current_point is None and self.seeds:
            try:
                current_ap, current_ml, current_dv = self.controller.get_current_position()
                current_point = (current_ml, current_ap)
            except Exception:
                current_point = None
        if self.drill_thread is not None and self.drill_thread.is_alive() and self.active_surface_dv is not None:
            try:
                _current_ap, _current_ml, current_dv = self.controller.get_current_position()
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
            current_point=current_point if self.seeds else None,
        )
        self.update_seed_selector_label()

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
    app.setFont(QFont("Segoe UI", 9))
    window = CraniotomyWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
