import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPainter, QPen
from PySide6.QtWidgets import QApplication, QWidget


class CircleWidget(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Draw Circle")
        self.resize(320, 320)

    def paintEvent(self, _event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QColor("#ffffff"))

        margin = 24
        diameter = min(self.width(), self.height()) - (margin * 2)
        x = (self.width() - diameter) / 2
        y = (self.height() - diameter) / 2

        painter.setPen(QPen(QColor("#1f2937"), 4))
        painter.setBrush(QColor("#60a5fa"))
        painter.drawEllipse(int(x), int(y), int(diameter), int(diameter))


def main() -> int:
    app = QApplication(sys.argv)
    window = CircleWidget()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
