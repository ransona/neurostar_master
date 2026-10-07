import json
import tkinter as tk
from pathlib import Path

IMAGE_PATH = Path(__file__).resolve().parents[1] / "assets/background_images/skull_bregma_lambda_reference.png"
METADATA_PATH = IMAGE_PATH.with_suffix(".json")

class CalibrationWindow:
    def __init__(self, root):
        self.root = root
        root.title("Bregma/Lambda Image Calibration")
        self.image = tk.PhotoImage(file=str(IMAGE_PATH))
        self.canvas = tk.Canvas(root, width=self.image.width(), height=self.image.height(), cursor="crosshair")
        self.canvas.pack()
        self.canvas.create_image(0, 0, image=self.image, anchor="nw")
        self.points = {}
        self.status = tk.Label(root, text="Click Bregma, then Lambda. Press R to reset", anchor="w")
        self.status.pack(fill="x")
        self.canvas.bind("<Button-1>", self.click)
        root.bind("<Key-r>", self.reset)

    def click(self, event):
        name = "bregma" if "bregma" not in self.points else "lambda" if "lambda" not in self.points else None
        if name is None:
            return
        self.points[name] = (int(event.x), int(event.y))
        color = "red" if name == "bregma" else "blue"
        self.canvas.create_oval(event.x - 6, event.y - 6, event.x + 6, event.y + 6, outline=color, width=2)
        self.canvas.create_text(event.x + 10, event.y - 10, text=name.title(), fill=color, anchor="sw")
        if len(self.points) == 2:
            payload = json.loads(METADATA_PATH.read_text(encoding="utf-8"))
            payload["calibration"]["bregma_pixel"] = list(self.points["bregma"])
            payload["calibration"]["lambda_pixel"] = list(self.points["lambda"])
            payload["calibration"]["status"] = "calibrated"
            METADATA_PATH.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
            self.status.config(text=f"Saved Bregma {self.points['bregma']} and Lambda {self.points['lambda']}")
        else:
            self.status.config(text="Bregma recorded. Now click Lambda.")

    def reset(self, _event=None):
        self.points.clear()
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.image, anchor="nw")
        self.status.config(text="Click Bregma, then Lambda. Press R to reset")

if __name__ == "__main__":
    app = tk.Tk()
    CalibrationWindow(app)
    app.mainloop()
