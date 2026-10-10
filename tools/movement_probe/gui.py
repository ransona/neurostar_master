"""Local operator window. HTTP/worker threads never touch Tk widgets."""
import json
import queue
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk


def event_visible(row, show_polling=False):
    return show_polling or not (
        row.get("event") in ("HTTP_REQUEST", "HTTP_RESULT")
        and row.get("endpoint") in ("/status", "/events")
        and row.get("status", 200) == 200
    )


def format_event(row):
    fields = {key: value for key, value in row.items() if key not in ("utc", "event", "sequence")}
    return f'{row["utc"]}  {row["event"]}  {json.dumps(fields, ensure_ascii=False)}\n'


def run_gui(service, url, log_path):
    root = tk.Tk()
    root.title("StereoDrive Movement Probe")
    root.geometry("900x560")
    root.minsize(680, 400)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(3, weight=1)
    header = ttk.Frame(root, padding=10)
    header.grid(row=0, column=0, sticky="ew")
    header.columnconfigure(1, weight=1)

    def read_only(label, value, row, secret=False):
        ttk.Label(header, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10))
        entry = ttk.Entry(header, show="*" if secret else "")
        entry.insert(0, str(value))
        entry.configure(state="readonly")
        entry.grid(row=row, column=1, sticky="ew", pady=2)
        return entry

    read_only("Server", url, 0)
    read_only("Event log", log_path, 1)

    controls = ttk.Frame(root, padding=(10, 0, 10, 5))
    controls.grid(row=1, column=0, sticky="ew")
    state = ttk.Label(controls, text="READY — same-computer connections only", foreground="darkgreen")
    state.pack(side="left", padx=(0, 15))
    updates = queue.Queue()
    closing = threading.Event()

    def background(action):
        def perform():
            try:
                action()
            except Exception as exc:
                updates.put(("error", str(exc)))
        threading.Thread(target=perform, daemon=True).start()

    def stop():
        background(lambda: service.stop("Local GUI Stop"))

    tk.Button(controls, text="STOP Movement (Esc)", command=stop,
              background="#b22222", foreground="white", padx=12).pack(side="left", padx=4)
    root.bind("<Escape>", lambda event: stop())
    mode = "SIMULATION — no hardware" if service.simulated else "REAL HARDWARE — supervised bench only"
    ttk.Label(controls, text=mode).pack(side="right")

    detail = tk.StringVar(value="Reading Axis position…")
    ttk.Label(root, textvariable=detail, padding=(10, 5), wraplength=850).grid(row=2, column=0, sticky="ew")
    log_box = scrolledtext.ScrolledText(root, wrap="word", state="disabled", font="TkFixedFont")
    log_box.grid(row=3, column=0, sticky="nsew", padx=10, pady=5)
    footer = ttk.Frame(root, padding=10)
    footer.grid(row=4, column=0, sticky="ew")
    polling = tk.BooleanVar(value=False)
    redraw = [True]
    ttk.Checkbutton(footer, text="Show status polling", variable=polling,
                    command=lambda: redraw.__setitem__(0, True)).pack(side="left")
    ttk.Label(footer, text="Incoming requests, results and movement events • UTC • last 1000 events").pack(side="right")

    def refresh_status():
        while not closing.is_set():
            try:
                snapshot = service.status()
                updates.put(("status", snapshot))
            except Exception as exc:
                updates.put(("status_error", str(exc)))
                service.stop("GUI status read failed: " + str(exc), fault=True)
            closing.wait(.3)

    threading.Thread(target=refresh_status, daemon=True).start()
    last_sequence = [0]

    def refresh():
        if closing.is_set():
            return
        while True:
            try:
                kind, value = updates.get_nowait()
            except queue.Empty:
                break
            if kind == "error":
                messagebox.showerror("Probe operation", value, parent=root)
            elif kind == "status_error":
                state.configure(text="READ ERROR — restart required", foreground="firebrick")
                detail.set(value + " — use physical Stop if movement persists.")
            else:
                state.configure(text="READY — local only" if value["ready"] else "FAULT — restart required",
                                foreground="darkgreen" if value["ready"] else "firebrick")
                p = value["position"]
                operation = value["operation"] or {}
                line = f"Axis AP {p[0]:.3f}  ML {p[1]:.3f}  DV {p[2]:.3f} mm | "
                line += f"Limit {value['max_move_mm']:.3f} mm/move | DV {'enabled' if value['allow_dv'] else 'disabled'}"
                line += f" | Injector ≤{value['max_injector_volume_nl']:g} nL/action"
                line += "\nBounds: " + "  ".join(
                    f"{axis} [{bounds[0]:.3f}, {bounds[1]:.3f}]" for axis, bounds in value["bounds"].items())
                if value.get("fault"):
                    line += "\nFault: " + value["fault"]
                if operation:
                    line += f"\n{operation['id']}: {operation['state']} → {operation['target']}"
                    if operation.get("error"):
                        line += " | " + operation["error"]
                if value.get("stop_error"):
                    line += "\nSTOP FAILED: " + value["stop_error"] + " — use physical Stop!"
                detail.set(line)
        with service.log_lock:
            rows = list(service.events[-1000:])
        sequence = rows[-1]["sequence"] if rows else 0
        if sequence != last_sequence[0] or redraw[0]:
            at_bottom = log_box.yview()[1] >= .99
            position = log_box.yview()[0]
            log_box.configure(state="normal")
            log_box.delete("1.0", "end")
            log_box.insert("end", "".join(format_event(row) for row in rows if event_visible(row, polling.get())))
            log_box.configure(state="disabled")
            log_box.see("end") if at_bottom else log_box.yview_moveto(position)
            last_sequence[0] = sequence
            redraw[0] = False
        root.after(200, refresh)

    def close():
        closing.set()
        # Returning to main immediately invokes service.close()/native Stop.
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", close)
    root.after(100, refresh)
    root.mainloop()
