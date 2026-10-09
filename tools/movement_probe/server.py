"""Supervised HTTP movement experiments; no USB packet injection or capture."""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import ipaddress
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import math
from pathlib import Path
import sys
import threading
import time
import uuid

AXES = ("AP", "ML", "DV")
STEPS = (0.01, 0.02, 0.05)


class Rejected(ValueError):
    pass


def number(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise Rejected("Expected a finite JSON number.")
    return float(value)


class SimulatedController:
    """API smoke tests only; does not model USB, mechanics, or GUI timing."""
    def __init__(self):
        self.position = [0.0, 0.0, 0.0]
        self.steps = {}
        self.cancelled = False

    def get_current_axis_position(self):
        return tuple(self.position)

    def prepare_motion(self):
        self.cancelled = False

    def stop(self):
        self.cancelled = True

    def set_nudge_step(self, axis, step):
        self.steps[axis] = step

    def nudge_axis(self, axis, positive):
        if self.cancelled:
            raise RuntimeError("Movement cancelled.")
        self.position[AXES.index(axis)] += self.steps[axis] * (1 if positive else -1)

    def goto_axis_position(self, *target, stop_requested=None):
        if self.cancelled or stop_requested():
            raise RuntimeError("Movement cancelled.")
        self.position[:] = target

    def safety_check(self):
        pass


def real_controller():
    if sys.platform != "win32":
        raise RuntimeError("Real movement requires the Windows computer running StereoDrive. Use --simulate here.")
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from stereodrive_controller import StereoDriveController

    class ProbeController(StereoDriveController):
        # The main GUI intentionally confirms some native warnings. This probe
        # must never approve those dialogs on behalf of an unattended agent.
        def safety_check(self):
            if self._find_below_skull_warning_dialog() is not None:
                raise RuntimeError("StereoDrive skull warning: operator intervention required; NOT accepted.")

        def confirm_below_skull_warning(self, **kwargs):
            self.safety_check()
            return False

        def confirm_no_actual_movement_dialog(self, **kwargs):
            return False

    return ProbeController()


class ProbeService:
    def __init__(self, controller, radius=0.1, max_move=0.05, allow_dv=False,
                 timeout=15.0, log_path=None):
        self.controller = controller
        self.origin = self.read_position()
        self.radius, self.max_move = radius, max_move
        self.allow_dv = allow_dv
        self.timeout = timeout
        self.log_path = log_path
        self.lock = threading.RLock()
        self.log_lock = threading.Lock()
        self.fault = None
        self.stop_event = threading.Event()
        self.closed = threading.Event()
        self.operation = None
        self.operations = {}
        self.events = []
        self.worker = None
        self.stop_error = None
        self.monitor = threading.Thread(target=self._monitor, daemon=True)
        self.monitor.start()
        self.log("START", origin=self.origin, radius_mm=radius, max_move_mm=max_move,
                 allow_dv=allow_dv, simulated=isinstance(controller, SimulatedController))

    def read_position(self):
        values = tuple(number(v) for v in self.controller.get_current_axis_position())
        if len(values) != 3:
            raise Rejected("Expected three mechanical Axis readings.")
        return values

    def log(self, event, **fields):
        with self.log_lock:
            row = dict(sequence=len(self.events) + 1,
                       utc=datetime.now(timezone.utc).isoformat(), event=event, **fields)
            if self.log_path:
                with open(self.log_path, "a", encoding="utf-8") as stream:
                    stream.write(json.dumps(row) + "\n")
                    stream.flush()
            self.events.append(row)

    def in_envelope(self, position):
        return all(abs(p - o) <= self.radius + 1e-9 for p, o in zip(position, self.origin))

    def stop(self, reason="Client Stop", fault=False):
        with self.lock:
            if fault:
                self.fault = reason
            self.stop_event.set()
            try:
                self.controller.stop()
            except Exception as exc:
                self.stop_error = str(exc)
                self.fault = "Native Stop failed: " + str(exc)
            self.log("STOP", reason=reason, stop_error=self.stop_error)

    def status(self):
        with self.lock:
            return dict(ready=not self.closed.is_set() and self.fault is None,
                        fault=self.fault, simulated=isinstance(self.controller, SimulatedController),
                        coordinates="mechanical Axis mm; never native Bregma",
                        position=self.read_position(), origin=self.origin,
                        bounds={a: [o - self.radius, o + self.radius] for a, o in zip(AXES, self.origin)},
                        max_move_mm=self.max_move, allow_dv=self.allow_dv,
                        stop_error=self.stop_error,
                        operation=dict(self.operation) if self.operation else None)

    def plan(self, payload, start):
        kind = payload.get("kind")
        method = payload.get("method", "goto")
        if method not in ("goto", "nudged", "planar"):
            raise Rejected("method must be goto, nudged, or planar.")
        target = list(start)
        if kind in ("nudge", "out_and_back", "axis"):
            axis = payload.get("axis")
            if axis not in AXES:
                raise Rejected("axis must be AP, ML, or DV.")
            index = AXES.index(axis)
            if kind == "axis":
                target[index] = number(payload.get("target_mm"))
            else:
                step = number(payload.get("step_mm"))
                if step not in STEPS:
                    raise Rejected("Single nudge step must be 0.01, 0.02, or 0.05 mm.")
                direction = payload.get("direction")
                if type(direction) is not int or direction not in (-1, 1):
                    raise Rejected("direction must be -1 or 1.")
                target[index] += step * direction
                method = "single_nudge"
        elif kind in ("relative", "absolute"):
            coords = payload.get("coordinates")
            if not isinstance(coords, dict) or not coords or set(coords) - set(AXES):
                raise Rejected("coordinates must contain AP, ML and/or DV.")
            for axis, value in coords.items():
                i = AXES.index(axis)
                target[i] = number(value) + (start[i] if kind == "relative" else 0)
        else:
            raise Rejected("kind must be nudge, out_and_back, axis, relative, or absolute.")
        if method == "goto":
            target = [round(v, 2) for v in target]
        if not self.allow_dv and abs(target[2] - start[2]) > 1e-9:
            raise Rejected("DV motion disabled; requires local startup --allow-dv approval.")
        if method == "planar" and abs(target[2] - start[2]) > 1e-9:
            raise Rejected("planar permits AP/ML only.")
        if math.dist(start, target) > self.max_move + 1e-9 or not self.in_envelope(target):
            raise Rejected("Movement exceeds per-command distance or startup envelope.")
        return method, tuple(target), kind == "out_and_back"

    def submit(self, payload):
        command_id = payload.get("command_id")
        if not isinstance(command_id, str) or not 1 <= len(command_id) <= 80:
            raise Rejected("Provide a unique command_id (1–80 characters) for retry deduplication.")
        with self.lock:
            if command_id in self.operations:
                old = self.operations[command_id]
                if old[0] != payload:
                    raise Rejected("command_id already used with different contents.")
                return dict(old[1])
            if self.closed.is_set() or self.fault is not None:
                raise Rejected("Server closed or faulted; resolve the fault and restart locally.")
            if self.worker and self.worker.is_alive():
                raise Rejected("Busy; no movement queue. Wait or Stop.")
            if len(self.operations) >= 1000:
                raise Rejected("Experiment limit reached; restart under local supervision.")
            start = self.read_position()
            self.controller.safety_check()
            if not self.in_envelope(start):
                self.stop("Position outside envelope", fault=True)
                raise Rejected("Position outside envelope.")
            method, target, reverse = self.plan(payload, start)
            self.operation = dict(id=command_id, state="running", start=start, target=target,
                                  method=method, error=None)
            self.operations[command_id] = (dict(payload), self.operation)
            self.stop_event.clear()
            self.controller.prepare_motion()
            self.log("COMMAND", command=payload, start=start, target=target, method=method)
            self.worker = threading.Thread(target=self._run, args=(method, start, target, reverse), daemon=True)
            self.worker.start()
            return dict(self.operation)

    def _check(self, deadline):
        if self.stop_event.is_set() or time.monotonic() >= deadline:
            raise RuntimeError("Movement cancelled or timed out.")
        self.controller.safety_check()
        current = self.read_position()
        if not self.in_envelope(current):
            raise RuntimeError("Observed position outside envelope.")
        return current

    def _wait(self, target, tolerance, deadline):
        # No blind retries/rearming of target boxes: an ambiguous move stops the
        # experiment instead of sending additional command traffic.
        while True:
            current = self._check(deadline)
            with self.lock:
                self.operation["position"] = current
            if all(abs(a - b) <= tolerance for a, b in zip(current, target)):
                self.log("ARRIVED", position=current, target=target)
                return
            time.sleep(0.05)

    def _move(self, method, target, deadline, nudge_delta=None):
        current = self._check(deadline)
        self.log("MOVE_REQUEST", method=method, target=target, position=current)
        if method == "goto":
            # Mechanical Axis edits are 0.01 mm resolution.
            target = tuple(round(v, 2) for v in target)
            if not self.in_envelope(target) or math.dist(current, target) > self.max_move + 1e-9:
                raise Rejected("Rounded GoTo target exceeds movement limits.")
            self.controller.goto_axis_position(*target, stop_requested=self.stop_event.is_set)
            self._wait(target, 0.006, deadline)
            return
        if method == "single_nudge":
            i = max(range(3), key=lambda i: abs(nudge_delta[i]))
            expected = list(current)
            expected[i] += nudge_delta[i]
            if not self.in_envelope(expected):
                raise Rejected("Observed position makes the requested nudge exceed the envelope.")
            # Reverse the same native step, not a jitter-derived unsupported
            # combo value. Arrival still must match the original start.
            self.controller.set_nudge_step(AXES[i], round(abs(nudge_delta[i]), 3))
            self.controller.nudge_axis(AXES[i], nudge_delta[i] > 0)
            self._wait(target, 0.003, deadline)
            return
        # Fine/planar moves issue one verified 0.01 mm increment at a time,
        # never the main app's rapid drilling loop. Planar interleaves AP/ML.
        for _ in range(100):
            current = self._check(deadline)
            axes = range(2) if method == "planar" else range(3)
            pending = [i for i in axes if abs(target[i] - current[i]) > 0.006]
            if not pending:
                self._wait(target, 0.006, deadline)
                return
            i = max(pending, key=lambda i: abs(target[i] - current[i])) if method == "planar" else pending[0]
            next_target = list(current)
            positive = target[i] > current[i]
            next_target[i] += 0.01 if positive else -0.01
            if not self.in_envelope(next_target):
                raise Rejected("Next fine nudge exceeds envelope.")
            self.controller.set_nudge_step(AXES[i], 0.01)
            self.controller.nudge_axis(AXES[i], positive)
            self._wait(next_target, 0.003, deadline)
        raise RuntimeError("Fine nudge iteration limit reached.")

    def _run(self, method, start, target, reverse):
        deadline = time.monotonic() + self.timeout
        try:
            delta = tuple(b - a for a, b in zip(start, target))
            self._move(method, target, deadline, delta)
            if reverse:
                self._check(deadline)
                self.log("REVERSE_REQUEST", target=start)
                self._move(method, start, deadline, tuple(-v for v in delta))
            self._check(deadline)
            with self.lock:
                self.operation["state"] = "completed"
            self.log("COMPLETED", id=self.operation["id"], position=self.read_position())
        except Exception as exc:
            # A normal Cancel/Stop is not a fault and needs no rearming.
            self.stop(str(exc), fault=not self.stop_event.is_set())
            with self.lock:
                self.operation.update(state="stopped", error=str(exc))

    def _monitor(self):
        while not self.closed.wait(0.1):
            try:
                with self.lock:
                    if self.fault is not None:
                        continue
                    self.controller.safety_check()
                    if not self.in_envelope(self.read_position()):
                        self.stop("Observed position outside envelope", fault=True)
            except Exception as exc:
                self.stop("Monitor failure: " + str(exc), fault=True)

    def close(self):
        self.closed.set()
        self.stop("Server shutdown")
        if self.worker:
            self.worker.join(timeout=5)
        self.monitor.join(timeout=2)


def handler(service):
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass  # Never print authentication headers.

        def reply(self, code, value):
            service.log("HTTP_RESULT", peer=self.client_address[0],
                        endpoint=self.path if self.path in ("/status", "/events", "/move", "/stop") else "<unknown>",
                        status=code, error=value.get("error") if isinstance(value, dict) else None)
            data = json.dumps(value).encode()
            self.send_response(code)
            self.send_header("Content-Type", "application/json")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(data)))
            try:
                self.end_headers()
                self.wfile.write(data)
            except (ConnectionError, OSError):
                # A lost reply is not a new controller fault. The operation ID
                # lets the local client inspect/retry without repeating motion.
                pass

        def dispatch(self):
            self.connection.settimeout(2)
            if not ipaddress.ip_address(self.client_address[0]).is_loopback:
                self.reply(403, {"error": "Same-computer connections only"})
                return
            # Restrict Host too: a malicious website must not use DNS rebinding
            # to reach this unauthenticated loopback service.
            hosts = (f"127.0.0.1:{self.server.server_port}", f"localhost:{self.server.server_port}")
            if self.headers.get("Host", "").lower() not in hosts:
                self.reply(403, {"error": "Loopback Host required"})
                return
            if self.headers.get("Origin") is not None or self.headers.get("Sec-Fetch-Site") is not None:
                self.reply(403, {"error": "Browser requests prohibited"})
                return
            endpoint = self.path if self.path in ("/status", "/events", "/move", "/stop") else "<unknown>"
            service.log("HTTP_REQUEST", peer=self.client_address[0], method=self.command, endpoint=endpoint)
            try:
                if self.command == "GET" and self.path == "/status":
                    result = service.status()
                elif self.command == "GET" and self.path == "/events":
                    with service.log_lock:
                        result = list(service.events)
                elif self.command == "POST":
                    if self.headers.get("Content-Type", "").split(";", 1)[0].strip().lower() != "application/json":
                        self.reply(415, {"error": "Content-Type must be application/json"})
                        return
                    size = int(self.headers.get("Content-Length", "0"))
                    if not 0 < size <= 4096:
                        raise Rejected("JSON request length must be 1–4096 bytes.")
                    payload = json.loads(self.rfile.read(size))
                    if not isinstance(payload, dict):
                        raise Rejected("Expected JSON object.")
                    if self.path == "/move":
                        # Show only the movement schema, never arbitrary extra
                        # headers/fields that could contain credentials.
                        service.log("INCOMING_MOVE", peer=self.client_address[0],
                                    command={key: payload[key] for key in
                                             ("command_id", "kind", "method", "axis", "direction", "step_mm", "target_mm", "coordinates")
                                             if key in payload})
                    if self.path == "/stop":
                        service.stop()
                        result = {"ok": service.stop_error is None, "stop_error": service.stop_error}
                    elif self.path == "/move":
                        result = service.submit(payload)
                    else:
                        self.reply(404, {"error": "Unknown endpoint"})
                        return
                else:
                    self.reply(404, {"error": "Unknown endpoint"})
                    return
                self.reply(200, result)
            except (ValueError, TypeError) as exc:
                self.reply(400, {"error": str(exc)})
            except Exception as exc:
                service.stop("API failure: " + str(exc), fault=True)
                self.reply(503, {"error": str(exc)})

        do_GET = dispatch
        do_POST = dispatch
    return Handler


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bind", default="127.0.0.1", help="Loopback only; LAN/public connections are prohibited")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--simulate", action="store_true")
    parser.add_argument("--console", action="store_true", help="Legacy terminal UI instead of the default small GUI")
    parser.add_argument("--allow-dv", action="store_true", help="Local operator allows DV experiments in an empty/retracted workspace.")
    parser.add_argument("--radius-mm", type=float, default=0.1)
    parser.add_argument("--max-move-mm", type=float, default=0.05)
    parser.add_argument("--log-dir", type=Path, default=Path.home() / "StereoDriveProbeLogs")
    args = parser.parse_args()
    try:
        address = ipaddress.ip_address(args.bind)
        if str(address) != "127.0.0.1":
            raise ValueError()
    except ValueError:
        parser.error("Bind must be 127.0.0.1; this token-free server only accepts same-computer connections.")
    if not (math.isfinite(args.radius_mm) and 0 < args.radius_mm <= 1
            and math.isfinite(args.max_move_mm) and 0 < args.max_move_mm <= min(args.radius_mm, 0.1)):
        parser.error("radius must be >0 and <=1 mm; max move >0 and <=0.1 mm and <=radius.")
    args.log_dir.mkdir(parents=True, exist_ok=True)
    log_path = args.log_dir / (datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ") + "-" + uuid.uuid4().hex[:8] + ".jsonl")
    service = ProbeService(SimulatedController() if args.simulate else real_controller(),
                           radius=args.radius_mm, max_move=args.max_move_mm,
                           allow_dv=args.allow_dv, log_path=log_path)
    server = None
    try:
        server = ThreadingHTTPServer((args.bind, args.port), handler(service))
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        if not args.console:
            from gui import run_gui
            run_gui(service, f"http://{args.bind}:{server.server_port}", log_path)
            return
        print(f"URL: http://{args.bind}:{server.server_port}\nLog: {log_path}")
        print(json.dumps(service.status(), indent=2))
        print("READY for local commands. No specimen; tool safely retracted; drill off. Close other automation.")
        print("Operator commands: STOP, STATUS, QUIT. No token, arming or heartbeat required.")
        while True:
            command = input("probe> ").strip().upper()
            if command == "STOP":
                service.stop("Local operator Stop")
                print("Stop requested. Use physical Stop if motion persists.")
            elif command == "STATUS":
                print(json.dumps(service.status(), indent=2))
            elif command == "QUIT":
                break
    except (KeyboardInterrupt, EOFError):
        pass
    finally:
        service.close()
        if server:
            server.shutdown()
            server.server_close()


if __name__ == "__main__":
    main()
