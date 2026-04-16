import math
import sys
from dataclasses import dataclass

from PySide6.QtCore import QPointF, Qt
from PySide6.QtGui import QColor, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import (
    QApplication,
    QDoubleSpinBox,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
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
        self.trajectory: list[tuple[float, float]] = []
        self.seed_points: list[tuple[float, float, bool]] = []
        self.setMinimumHeight(220)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

    def set_data(self, trajectory: list[tuple[float, float]], seed_points: list[tuple[float, float, bool]]) -> None:
        self.trajectory = trajectory
        self.seed_points = seed_points
        self.update()

    def paintEvent(self, _event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QColor("#fbfcfa"))

        pad = 24
        draw_rect = self.rect().adjusted(pad, pad + 20, -pad, -pad)

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
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)
        if math.isclose(min_x, max_x):
            min_x -= 1.0
            max_x += 1.0
        if math.isclose(min_y, max_y):
            min_y -= 1.0
            max_y += 1.0

        def map_point(x: float, y: float) -> QPointF:
            px = draw_rect.left() + (x - min_x) / (max_x - min_x) * draw_rect.width()
            normalized_y = (y - min_y) / (max_y - min_y)
            if self.invert_y:
                py = draw_rect.top() + normalized_y * draw_rect.height()
            else:
                py = draw_rect.bottom() - normalized_y * draw_rect.height()
            return QPointF(px, py)

        if len(self.trajectory) > 1:
            path = QPainterPath()
            path.moveTo(map_point(*self.trajectory[0]))
            for point in self.trajectory[1:]:
                path.lineTo(map_point(*point))
            painter.setPen(QPen(QColor("#0d8a63"), 3))
            painter.drawPath(path)

        for idx, (x, y, sampled) in enumerate(self.seed_points, start=1):
            pt = map_point(x, y)
            color = QColor("#0d8a63" if sampled else "#dd6e42")
            painter.setPen(Qt.NoPen)
            painter.setBrush(color)
            painter.drawEllipse(pt, 4.5, 4.5)
            painter.setPen(color)
            painter.drawText(pt + QPointF(8, -8), str(idx))


class CraniotomyWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.controller = StereoDriveController()
        self.seeds: list[SeedPoint] = []
        self.trajectory: list[tuple[float, float, float]] = []
        self.current_seed_index: int | None = None

        self.setWindowTitle("Craniotomy Planner")
        self.resize(1360, 900)
        self._build_ui()
        self.refresh_live_position()

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        root.setStyleSheet(
            """
            QWidget {
                background: #eef3ea;
                color: #173122;
                font-family: 'Segoe UI';
                font-size: 13px;
            }
            QGroupBox {
                border: 1px solid #d4ded3;
                border-radius: 18px;
                margin-top: 14px;
                background: rgba(255,255,255,0.92);
                font-weight: 600;
                padding-top: 12px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 16px;
                padding: 0 6px;
            }
            QPushButton {
                border: none;
                border-radius: 18px;
                padding: 10px 16px;
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
                border-radius: 12px;
                padding: 8px;
                background: white;
                min-height: 20px;
            }
            QTableWidget {
                border: 1px solid #d4ded3;
                border-radius: 14px;
                background: white;
                gridline-color: #edf2ec;
            }
            QHeaderView::section {
                background: #f6f8f4;
                border: none;
                border-bottom: 1px solid #edf2ec;
                padding: 8px;
                font-weight: 600;
            }
            """
        )

        layout = QVBoxLayout(root)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(18)

        hero = QFrame()
        hero_layout = QHBoxLayout(hero)
        hero_layout.setContentsMargins(0, 0, 0, 0)
        hero_layout.setSpacing(18)
        layout.addWidget(hero)

        hero_text = QGroupBox()
        hero_text_layout = QVBoxLayout(hero_text)
        hero_text_layout.setContentsMargins(22, 22, 22, 22)
        title = QLabel("Craniotomy Planner")
        title.setProperty("role", "hero")
        subtitle = QLabel(
            "Capture the midpoint from the live frame, sample skull surface depth around the rim, "
            "and interpolate a circular trajectory that follows the skull contour."
        )
        subtitle.setWordWrap(True)
        subtitle.setProperty("role", "muted")
        hero_text_layout.addWidget(title)
        hero_text_layout.addWidget(subtitle)
        hero_layout.addWidget(hero_text, 3)

        live_box = QGroupBox("StereoDrive")
        live_layout = QGridLayout(live_box)
        live_layout.setContentsMargins(18, 18, 18, 18)
        live_layout.setHorizontalSpacing(16)
        live_layout.setVerticalSpacing(8)
        self.reference_label = QLabel("-")
        self.ap_label = QLabel("-")
        self.ml_label = QLabel("-")
        self.dv_label = QLabel("-")
        refresh_btn = QPushButton("Refresh Live Position")
        refresh_btn.clicked.connect(self.refresh_live_position)
        live_layout.addWidget(QLabel("Reference"), 0, 0)
        live_layout.addWidget(self.reference_label, 0, 1)
        live_layout.addWidget(refresh_btn, 0, 2, 1, 2)
        live_layout.addWidget(QLabel("AP"), 1, 0)
        live_layout.addWidget(self.ap_label, 1, 1)
        live_layout.addWidget(QLabel("ML"), 1, 2)
        live_layout.addWidget(self.ml_label, 1, 3)
        live_layout.addWidget(QLabel("DV"), 1, 4)
        live_layout.addWidget(self.dv_label, 1, 5)
        hero_layout.addWidget(live_box, 2)

        content = QHBoxLayout()
        content.setSpacing(18)
        layout.addLayout(content, 1)

        left_column = QVBoxLayout()
        left_column.setSpacing(18)
        content.addLayout(left_column, 1)

        setup_box = QGroupBox("Setup")
        setup_layout = QGridLayout(setup_box)
        setup_layout.setContentsMargins(18, 18, 18, 18)
        setup_layout.setHorizontalSpacing(14)
        setup_layout.setVerticalSpacing(12)
        left_column.addWidget(setup_box)

        midpoint_btn = QPushButton("Use Current AP/ML As Midpoint")
        midpoint_btn.setProperty("variant", "primary")
        midpoint_btn.style().unpolish(midpoint_btn)
        midpoint_btn.style().polish(midpoint_btn)
        midpoint_btn.clicked.connect(self.capture_midpoint)

        self.mid_ap = self._double_spinbox()
        self.mid_ml = self._double_spinbox()
        self.travel_dv = self._double_spinbox()
        self.diameter = self._double_spinbox(value=3.0, minimum=0.1, maximum=20.0)
        self.seed_count = self._spinbox(value=6, minimum=3, maximum=24)
        self.trajectory_points = self._spinbox(value=60, minimum=12, maximum=360)
        self.cut_offset = self._double_spinbox(value=0.0, minimum=-5.0, maximum=5.0)

        setup_layout.addWidget(midpoint_btn, 0, 0, 1, 2)
        setup_layout.addWidget(QLabel("Mid AP (mm)"), 0, 2)
        setup_layout.addWidget(self.mid_ap, 0, 3)
        setup_layout.addWidget(QLabel("Mid ML (mm)"), 0, 4)
        setup_layout.addWidget(self.mid_ml, 0, 5)

        setup_layout.addWidget(QLabel("Travel DV"), 1, 0)
        setup_layout.addWidget(self.travel_dv, 1, 1)
        setup_layout.addWidget(QLabel("Diameter (mm)"), 1, 2)
        setup_layout.addWidget(self.diameter, 1, 3)
        setup_layout.addWidget(QLabel("Seed Points"), 1, 4)
        setup_layout.addWidget(self.seed_count, 1, 5)

        setup_layout.addWidget(QLabel("Trajectory Points"), 2, 0)
        setup_layout.addWidget(self.trajectory_points, 2, 1)
        setup_layout.addWidget(QLabel("Cut Offset DV"), 2, 2)
        setup_layout.addWidget(self.cut_offset, 2, 3)

        generate_btn = QPushButton("Generate Seed Points")
        generate_btn.clicked.connect(self.generate_seeds)
        stop_btn = QPushButton("Stop Motion")
        stop_btn.setProperty("variant", "danger")
        stop_btn.style().unpolish(stop_btn)
        stop_btn.style().polish(stop_btn)
        stop_btn.clicked.connect(self.stop_motion)
        setup_layout.addWidget(generate_btn, 2, 4)
        setup_layout.addWidget(stop_btn, 2, 5)

        workflow_box = QGroupBox("Workflow")
        workflow_layout = QVBoxLayout(workflow_box)
        workflow_layout.setContentsMargins(18, 18, 18, 18)
        workflow_layout.setSpacing(12)
        left_column.addWidget(workflow_box, 1)

        button_row = QHBoxLayout()
        self.move_seed_btn = QPushButton("Move To Current Seed")
        self.move_seed_btn.clicked.connect(self.move_to_current_seed)
        self.capture_surface_btn = QPushButton("At Surface / Capture DV")
        self.capture_surface_btn.setProperty("variant", "primary")
        self.capture_surface_btn.style().unpolish(self.capture_surface_btn)
        self.capture_surface_btn.style().polish(self.capture_surface_btn)
        self.capture_surface_btn.clicked.connect(self.capture_surface)
        button_row.addWidget(self.move_seed_btn)
        button_row.addWidget(self.capture_surface_btn)
        workflow_layout.addLayout(button_row)

        self.seed_table = QTableWidget(0, 6)
        self.seed_table.setHorizontalHeaderLabels(["#", "Angle", "AP", "ML", "DV", "State"])
        self.seed_table.verticalHeader().setVisible(False)
        self.seed_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.seed_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.seed_table.cellClicked.connect(self.on_seed_selected)
        workflow_layout.addWidget(self.seed_table, 1)

        self.status_label = QLabel("Ready.")
        self.status_label.setWordWrap(True)
        self.status_label.setProperty("role", "muted")
        workflow_layout.addWidget(self.status_label)

        right_column = QVBoxLayout()
        right_column.setSpacing(18)
        content.addLayout(right_column, 1)

        views_box = QGroupBox("Trajectory Views")
        views_layout = QGridLayout(views_box)
        views_layout.setContentsMargins(18, 18, 18, 18)
        views_layout.setSpacing(14)
        right_column.addWidget(views_box, 1)

        self.top_view = ProjectionWidget("Top View", "AP", "ML")
        self.back_view = ProjectionWidget("Back View", "ML", "DV", invert_y=True)
        self.side_view = ProjectionWidget("Side View", "AP", "DV", invert_y=True)
        views_layout.addWidget(self.top_view, 0, 0)
        views_layout.addWidget(self.back_view, 0, 1)
        views_layout.addWidget(self.side_view, 1, 0, 1, 2)

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
        self.status_label.setText(message)

    def refresh_live_position(self) -> None:
        try:
            ap, ml, dv = self.controller.get_current_position()
            self.reference_label.setText(self.controller.get_reference_selector() or "Unknown")
            self.ap_label.setText(f"{ap:.2f}")
            self.ml_label.setText(f"{ml:.2f}")
            self.dv_label.setText(f"{dv:.2f}")
            if self.travel_dv.value() == 0.0:
                self.travel_dv.setValue(dv)
            self.set_status("StereoDrive position refreshed.")
        except Exception as exc:
            self.set_status(str(exc))

    def capture_midpoint(self) -> None:
        try:
            ap, ml, dv = self.controller.get_current_position()
            self.mid_ap.setValue(ap)
            self.mid_ml.setValue(ml)
            self.travel_dv.setValue(dv)
            self.refresh_live_position()
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
            self.refresh_seed_table()
            self.redraw_views()
            self.set_status(f"Generated {seed_count} seed points.")
        except Exception as exc:
            QMessageBox.critical(self, "Craniotomy", str(exc))

    def refresh_seed_table(self) -> None:
        self.seed_table.setRowCount(len(self.seeds))
        for seed in self.seeds:
            state = "Captured" if seed.dv is not None else "Pending"
            values = [
                str(seed.index + 1),
                f"{seed.angle_deg:.1f}°",
                f"{seed.ap:.2f}",
                f"{seed.ml:.2f}",
                "" if seed.dv is None else f"{seed.dv:.2f}",
                state,
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setTextAlignment(Qt.AlignCenter)
                if seed.index == self.current_seed_index:
                    item.setBackground(QColor("#dff0e5"))
                self.seed_table.setItem(seed.index, column, item)
        if self.current_seed_index is not None and self.current_seed_index < len(self.seeds):
            self.seed_table.selectRow(self.current_seed_index)

    def on_seed_selected(self, row: int, _column: int) -> None:
        self.current_seed_index = row
        self.refresh_seed_table()

    def move_to_current_seed(self) -> None:
        if self.current_seed_index is None or not self.seeds:
            QMessageBox.information(self, "Craniotomy", "Generate seed points first.")
            return
        try:
            seed = self.seeds[self.current_seed_index]
            self.controller.goto_position(seed.ap, seed.ml, self.travel_dv.value())
            self.refresh_live_position()
            self.set_status(
                f"Moved to seed {seed.index + 1}. Lower manually to the skull surface, then click 'At Surface / Capture DV'."
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
            self.refresh_seed_table()
            if next_pending is None:
                self.set_status("Captured all seed points and updated the trajectory.")
            else:
                self.set_status(f"Captured seed {seed.index + 1}. Next pending seed: {next_pending + 1}.")
        except Exception as exc:
            QMessageBox.critical(self, "StereoDrive", str(exc))

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

    def redraw_views(self) -> None:
        top_points = [(ap, ml) for ap, ml, _ in self.trajectory]
        back_points = [(ml, dv) for _, ml, dv in self.trajectory]
        side_points = [(ap, dv) for ap, _, dv in self.trajectory]
        fallback_dv = self.travel_dv.value()
        top_seeds = [(seed.ap, seed.ml, seed.dv is not None) for seed in self.seeds]
        back_seeds = [(seed.ml, seed.dv if seed.dv is not None else fallback_dv, seed.dv is not None) for seed in self.seeds]
        side_seeds = [(seed.ap, seed.dv if seed.dv is not None else fallback_dv, seed.dv is not None) for seed in self.seeds]
        self.top_view.set_data(top_points, top_seeds)
        self.back_view.set_data(back_points, back_seeds)
        self.side_view.set_data(side_points, side_seeds)


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Craniotomy Planner")
    window = CraniotomyWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
