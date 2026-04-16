import math
import sys
from dataclasses import dataclass

from PySide6.QtCore import QPointF, QRectF, Qt, QTimer
from PySide6.QtGui import QColor, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import (
    QApplication,
    QDoubleSpinBox,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from stereodrive_controller import StereoDriveController, StereoDriveError


@dataclass
class SeedPoint:
    index: int
    angle_deg: float
    ap: float
    ml: float
    dv: float | None = None
    sampled_ap: float | None = None
    sampled_ml: float | None = None


class ProjectionWidget(QWidget):
    def __init__(self, title: str, x_label: str, y_label: str, invert_y: bool = False, parent: QWidget | None = None):
        super().__init__(parent)
        self.title = title
        self.x_label = x_label
        self.y_label = y_label
        self.invert_y = invert_y
        self.trajectory: list[tuple[float, float, float]] = []
        self.seed_points: list[tuple[float, float, bool]] = []
        self.current_point: tuple[float, float] | None = None
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
        current_point: tuple[float, float] | None = None,
    ) -> None:
        self.trajectory = trajectory
        self.seed_points = seed_points
        self.current_point = current_point
        self.update()

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
        painter.drawText(16, 20, f"{self.title} ({self.x_label} vs {self.y_label})")

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
        uniform_span = max(span_x, span_y)
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

        if len(self.trajectory) > 1:
            for start, end in zip(self.trajectory[:-1], self.trajectory[1:]):
                progress = max(0.0, min(1.0, (start[2] + end[2]) / 2.0))
                red = int(130 + progress * 90)
                green = int(170 - progress * 70)
                blue = int(150 - progress * 50)
                painter.setPen(QPen(QColor(red, green, blue), 6))
                painter.drawLine(map_point(start[0], start[1]), map_point(end[0], end[1]))

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


class CraniotomyWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.controller = StereoDriveController()
        self.seeds: list[SeedPoint] = []
        self.trajectory: list[tuple[float, float, float]] = []
        self.current_seed_index: int | None = None
        self.setWindowTitle("Craniotomy Planner")
        self.resize(1120, 760)
        self._build_ui()
        self.refresh_live_position()
        self.refresh_timer = QTimer(self)
        self.refresh_timer.timeout.connect(self.refresh_live_position)
        self.refresh_timer.start(250)

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        root.setStyleSheet(
            """
            QWidget {
                background: #eef3ea;
                color: #173122;
                font-family: 'Segoe UI';
                font-size: 12px;
            }
            QGroupBox {
                border: 1px solid #d4ded3;
                border-radius: 14px;
                margin-top: 10px;
                background: rgba(255,255,255,0.92);
                font-weight: 600;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }
            QPushButton {
                border: none;
                border-radius: 14px;
                padding: 8px 12px;
                background: #dceae0;
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
            QLabel[role="hero"] {
                font-size: 34px;
                font-weight: 700;
            }
            QLabel[role="muted"] {
                color: #5e7064;
            }
            QDoubleSpinBox, QSpinBox {
                border: 1px solid #cfdbcf;
                border-radius: 10px;
                padding: 6px;
                background: white;
                min-height: 18px;
            }
            """
        )

        layout = QVBoxLayout(root)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        content = QVBoxLayout()
        content.setSpacing(12)
        layout.addLayout(content, 1)

        setup_box = QGroupBox("Setup")
        setup_layout = QGridLayout(setup_box)
        setup_layout.setContentsMargins(14, 14, 14, 14)
        setup_layout.setHorizontalSpacing(10)
        setup_layout.setVerticalSpacing(8)
        content.addWidget(setup_box)

        midpoint_btn = QPushButton("Use Current AP/ML As Midpoint")
        midpoint_btn.setProperty("variant", "primary")
        midpoint_btn.style().unpolish(midpoint_btn)
        midpoint_btn.style().polish(midpoint_btn)
        midpoint_btn.clicked.connect(self.capture_midpoint)

        self.mid_ap = self._double_spinbox()
        self.mid_ml = self._double_spinbox()
        self.current_ap_label = QLabel("-")
        self.current_ml_label = QLabel("-")
        self.current_dv_label = QLabel("-")
        self.diameter = self._double_spinbox(value=3.0, minimum=0.1, maximum=20.0)
        self.seed_count = self._spinbox(value=6, minimum=3, maximum=24)
        self.trajectory_points = self._spinbox(value=60, minimum=12, maximum=360)
        self.cut_offset = self._double_spinbox(value=0.0, minimum=-5.0, maximum=5.0)
        self.current_seed_spin = self._spinbox(value=1, minimum=1, maximum=1)
        self.current_seed_spin.valueChanged.connect(self.on_seed_spin_changed)
        self.current_seed_coords = QLabel("Seed: -")
        self.current_seed_coords.setWordWrap(True)
        self.current_seed_coords.setProperty("role", "muted")

        setup_layout.addWidget(QLabel("Current AP"), 0, 0)
        setup_layout.addWidget(self.current_ap_label, 0, 1)
        setup_layout.addWidget(QLabel("Current ML"), 0, 2)
        setup_layout.addWidget(self.current_ml_label, 0, 3)
        setup_layout.addWidget(QLabel("Current DV"), 0, 4)
        setup_layout.addWidget(self.current_dv_label, 0, 5)

        setup_layout.addWidget(midpoint_btn, 1, 0, 1, 2)
        setup_layout.addWidget(QLabel("Mid AP"), 1, 2)
        setup_layout.addWidget(self.mid_ap, 1, 3)
        setup_layout.addWidget(QLabel("Mid ML"), 1, 4)
        setup_layout.addWidget(self.mid_ml, 1, 5)

        setup_layout.addWidget(QLabel("Diameter (mm)"), 2, 0)
        setup_layout.addWidget(self.diameter, 2, 1)
        setup_layout.addWidget(QLabel("Seed Points"), 2, 2)
        setup_layout.addWidget(self.seed_count, 2, 3)
        setup_layout.addWidget(QLabel("Trajectory Points"), 2, 4)
        setup_layout.addWidget(self.trajectory_points, 2, 5)

        setup_layout.addWidget(QLabel("Cut Offset DV"), 3, 0)
        setup_layout.addWidget(self.cut_offset, 3, 1)
        setup_layout.addWidget(QLabel("Current Seed"), 3, 2)
        setup_layout.addWidget(self.current_seed_spin, 3, 3)
        setup_layout.addWidget(self.current_seed_coords, 3, 4, 1, 2)

        generate_btn = QPushButton("Generate Seeds")
        generate_btn.clicked.connect(self.generate_seeds)
        self.move_seed_btn = QPushButton("Move To Current Seed")
        self.move_seed_btn.clicked.connect(self.move_to_current_seed)
        self.capture_surface_btn = QPushButton("At Surface / Capture DV")
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
        setup_layout.addWidget(generate_btn, 4, 0, 1, 2)
        setup_layout.addWidget(clear_btn, 4, 2, 1, 2)
        setup_layout.addWidget(stop_btn, 4, 4, 1, 2)
        setup_layout.addWidget(self.move_seed_btn, 5, 0, 1, 3)
        setup_layout.addWidget(self.capture_surface_btn, 5, 3, 1, 3)

        views_box = QGroupBox("Trajectory View")
        views_layout = QGridLayout(views_box)
        views_layout.setContentsMargins(14, 14, 14, 14)
        content.addWidget(views_box, 1)

        self.top_view = ProjectionWidget("Top View", "AP", "ML")
        self.top_view.setMinimumSize(420, 420)
        self.top_view.setMaximumWidth(620)
        views_layout.addWidget(self.top_view, 0, 0)

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

    def set_status(self, message: str) -> None:
        return

    def refresh_live_position(self) -> None:
        try:
            ap, ml, dv = self.controller.get_current_position()
            self.current_ap_label.setText(f"{ap:.2f}")
            self.current_ml_label.setText(f"{ml:.2f}")
            self.current_dv_label.setText(f"{dv:.2f}")
            if self.seeds:
                self.redraw_views(current_point=(ap, ml))
        except Exception as exc:
            self.set_status(str(exc))

    def capture_midpoint(self) -> None:
        try:
            ap, ml, _dv = self.controller.get_current_position()
            self.mid_ap.setValue(ap)
            self.mid_ml.setValue(ml)
            self.set_status("Captured current AP/ML as the craniotomy midpoint.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def generate_seeds(self) -> None:
        try:
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
            self.redraw_views()
            self.set_status(f"Generated {seed_count} seed points.")
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
                f"Moved to seed {seed.index + 1} target [{seed.ap:.2f}, {seed.ml:.2f}, -1.00]. Lower manually to the skull surface, then click 'At Surface / Capture DV'."
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
            self.controller.stop()
            self.set_status("Sent Stop command.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

    def redraw_views(self, current_point: tuple[float, float] | None = None) -> None:
        top_points = [(ap, ml, 0.0) for ap, ml, _ in self.trajectory]
        top_seeds = [(seed.ap, seed.ml, seed.dv is not None) for seed in self.seeds]
        if current_point is None and self.seeds:
            try:
                current_ap, current_ml, _current_dv = self.controller.get_current_position()
                current_point = (current_ap, current_ml)
            except Exception:
                current_point = None
        self.top_view.set_data(top_points, top_seeds, current_point=current_point if self.seeds else None)
        self.update_seed_selector_label()


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Craniotomy Planner")
    window = CraniotomyWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
