"""Network probe tests using a fake controller; never contact hardware."""
import importlib.util
import json
from pathlib import Path
import sys
import threading
import time
import tempfile
from types import SimpleNamespace
from unittest.mock import patch
import unittest
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer

spec = importlib.util.spec_from_file_location("movement_probe_server", Path(__file__).resolve().parents[1] / "tools/movement_probe/server.py")
probe = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = probe
spec.loader.exec_module(probe)

# GUI formatting tests need no Tk installation/display.
import ast
gui_tree = ast.parse((Path(__file__).resolve().parents[1] / "tools/movement_probe/gui.py").read_text())
gui_helpers = {"json": json}
exec(compile(ast.Module(body=[node for node in gui_tree.body if isinstance(node, ast.FunctionDef)
                            and node.name in ("event_visible", "format_event")], type_ignores=[]),
             "gui_helpers", "exec"), gui_helpers)


class ProbeTests(unittest.TestCase):
    def setUp(self):
        self.controller = probe.SimulatedController()
        self.service = probe.ProbeService(self.controller)
        self.addCleanup(self.service.close)
        self.service.arm_local()

    def submit(self, **kwargs):
        payload = dict(command_id=str(time.monotonic_ns()), **kwargs)
        result = self.service.submit(payload)
        self.service.worker.join(timeout=2)
        self.assertFalse(self.service.worker.is_alive())
        return self.service.status()["operation"]

    def test_all_movement_modes(self):
        self.assertEqual(self.submit(kind="nudge", axis="AP", direction=1, step_mm=.01)["state"], "completed")
        self.submit(kind="relative", coordinates={"ML": .02})
        self.submit(kind="axis", axis="AP", target_mm=.03, method="nudged")
        self.submit(kind="absolute", coordinates={"AP": .01, "ML": .01}, method="planar")
        self.assertAlmostEqual(self.controller.position[0], .01)
        self.assertAlmostEqual(self.controller.position[1], .01)

    def test_out_and_back_verified(self):
        self.submit(kind="out_and_back", axis="ML", direction=-1, step_mm=.02)
        self.assertEqual(tuple(self.controller.position), (0, 0, 0))
        names = [e["event"] for e in self.service.events]
        self.assertLess(names.index("ARRIVED"), names.index("REVERSE_REQUEST"))

    def test_bad_commands_and_dv_rejected(self):
        for payload in [dict(kind="relative", coordinates={"AP": .06}),
                        dict(kind="relative", coordinates={"DV": .01}),
                        dict(kind="axis", axis="AP", target_mm=float("nan")),
                        dict(kind="nudge", axis="AP", step_mm=.001, direction=1),
                        dict(kind="absolute", coordinates={"AP": True}),
                        dict(kind="relative", coordinates={"Bregma": .01})]:
            with self.subTest(payload=payload), self.assertRaises(probe.Rejected):
                self.service.submit(dict(command_id=str(time.monotonic_ns()), **payload))
        self.assertEqual(self.controller.position, [0, 0, 0])

    def test_envelope_prevents_accumulated_drift(self):
        self.controller.position[0] = .09
        with self.assertRaises(probe.Rejected):
            self.submit(kind="relative", coordinates={"AP": .02})

    def test_duplicate_is_not_executed_again(self):
        payload = dict(command_id="same", kind="nudge", axis="AP", step_mm=.01, direction=1)
        self.service.submit(payload)
        self.service.worker.join(timeout=2)
        self.service.submit(payload)
        self.assertAlmostEqual(self.controller.position[0], .01)
        with self.assertRaises(probe.Rejected):
            self.service.submit(dict(payload, direction=-1))

    def test_stop_requires_local_rearm(self):
        self.service.stop()
        with self.assertRaises(probe.Rejected):
            self.service.heartbeat()
        with self.assertRaises(probe.Rejected):
            self.submit(kind="relative", coordinates={"AP": .01})
        self.service.arm_local()
        self.submit(kind="relative", coordinates={"AP": .01})

    def test_stale_heartbeat_cannot_revive_lease(self):
        self.service.last_heartbeat -= 4
        with self.assertRaises(probe.Rejected):
            self.service.heartbeat()
        self.assertFalse(self.service.armed)
        self.assertTrue(self.controller.cancelled)

    def test_monitor_stops_after_heartbeat_loss(self):
        self.service.last_heartbeat -= 4
        self.service.monitor.join(timeout=.3)
        self.assertFalse(self.service.armed)
        self.assertTrue(self.controller.cancelled)

    def test_failed_forward_does_not_reverse(self):
        self.controller.nudge_axis = lambda *args: (_ for _ in ()).throw(RuntimeError("readout failed"))
        result = self.submit(kind="out_and_back", axis="AP", direction=1, step_mm=.01)
        self.assertEqual(result["state"], "stopped")
        self.assertFalse(self.service.armed)
        self.assertNotIn("REVERSE_REQUEST", [e["event"] for e in self.service.events])

    def test_stop_cancels_worker_and_busy_is_rejected(self):
        self.controller.goto_axis_position = lambda *args, **kwargs: None
        self.service.submit(dict(command_id="waiting", kind="relative", coordinates={"AP": .01}))
        with self.assertRaises(probe.Rejected):
            self.service.submit(dict(command_id="queued", kind="relative", coordinates={"AP": .01}))
        self.service.stop()
        self.service.worker.join(timeout=2)
        self.assertEqual(self.service.status()["operation"]["state"], "stopped")
        self.assertFalse(self.service.worker.is_alive())

    def test_timeout_disarms(self):
        self.service.timeout = .05
        self.controller.goto_axis_position = lambda *args, **kwargs: None
        result = self.submit(kind="relative", coordinates={"AP": .01})
        self.assertEqual(result["state"], "stopped")
        self.assertFalse(self.service.armed)

    def test_dv_explicitly_enabled(self):
        self.service.allow_dv = True
        self.submit(kind="relative", coordinates={"DV": -.01})
        self.assertAlmostEqual(self.controller.position[2], -.01)

    def test_http_auth_origin_and_endpoints(self):
        self.service.lease_seconds = 30  # Network setup latency isn't a lease test.
        server = ThreadingHTTPServer(("127.0.0.1", 0), probe.handler(self.service, "test-token"))
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(server.server_close)
        self.addCleanup(server.shutdown)
        url = f"http://127.0.0.1:{server.server_port}"

        def call(path, payload=None, token="test-token", origin=None):
            headers = {"Authorization": "Bearer " + token}
            if origin:
                headers["Origin"] = origin
            req = urllib.request.Request(url + path, headers=headers,
                                         data=None if payload is None else json.dumps(payload).encode())
            opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
            with opener.open(req, timeout=2) as response:
                return json.load(response)

        with self.assertRaises(urllib.error.HTTPError) as caught:
            call("/status", token="wrong")
        self.assertEqual(caught.exception.code, 401)
        with self.assertRaises(urllib.error.HTTPError) as caught:
            call("/stop", {}, origin="https://example.com")
        self.assertEqual(caught.exception.code, 403)
        self.assertTrue(call("/status")["armed"])
        self.assertTrue(call("/heartbeat", {})["ok"])
        call("/move", dict(command_id="http", kind="nudge", axis="AP", direction=1, step_mm=.01))
        self.service.worker.join(timeout=2)
        self.assertEqual(call("/status")["operation"]["state"], "completed")
        self.assertTrue(call("/events"))
        call("/stop", {})
        self.assertFalse(call("/status")["armed"])
        events = self.service.events
        incoming = [row for row in events if row["event"] == "INCOMING_MOVE"]
        self.assertEqual(incoming[0]["command"]["command_id"], "http")
        self.assertIn("peer", incoming[0])
        self.assertTrue(any(row["event"] == "HTTP_RESULT" and row["status"] == 401 for row in events))
        self.assertNotIn("test-token", json.dumps(events))

    def test_rejected_movement_is_visible_in_audit(self):
        server = ThreadingHTTPServer(("127.0.0.1", 0), probe.handler(self.service, "secret"))
        threading.Thread(target=server.serve_forever, daemon=True).start()
        self.addCleanup(server.server_close)
        self.addCleanup(server.shutdown)
        payload = dict(command_id="bad", kind="relative", coordinates={"DV": .01}, password="do-not-log")
        req = urllib.request.Request(f"http://127.0.0.1:{server.server_port}/move",
                                     data=json.dumps(payload).encode(), headers={"Authorization": "Bearer secret"})
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with self.assertRaises(urllib.error.HTTPError) as caught:
            opener.open(req, timeout=2)
        self.assertEqual(caught.exception.code, 400)
        self.assertTrue(any(row["event"] == "INCOMING_MOVE" for row in self.service.events))
        self.assertTrue(any(row["event"] == "HTTP_RESULT" and row["status"] == 400 for row in self.service.events))
        self.assertNotIn("do-not-log", json.dumps(self.service.events))

    def test_gui_event_filter_keeps_commands_and_failures(self):
        visible = gui_helpers["event_visible"]
        self.assertFalse(visible(dict(event="HTTP_RESULT", endpoint="/heartbeat", status=200)))
        self.assertTrue(visible(dict(event="HTTP_RESULT", endpoint="/heartbeat", status=400)))
        self.assertTrue(visible(dict(event="HTTP_REQUEST", endpoint="/move")))
        self.assertTrue(visible(dict(event="HTTP_REQUEST", endpoint="/status"), True))
        row = dict(sequence=1, utc="2026-10-09T10:00:00+00:00", event="INCOMING_MOVE", command={"kind":"nudge"})
        self.assertIn('"kind": "nudge"', gui_helpers["format_event"](row))

    def test_real_adapter_does_not_accept_skull_warning(self):
        class FakeNative:
            def _find_below_skull_warning_dialog(self):
                return 123
        with patch.object(probe.sys, "platform", "win32"), \
                patch.dict(sys.modules, {"stereodrive_controller": SimpleNamespace(StereoDriveController=FakeNative)}):
            native = probe.real_controller()
            with self.assertRaisesRegex(RuntimeError, "NOT accepted"):
                native.confirm_below_skull_warning(timeout_seconds=.01)
            self.assertFalse(native.confirm_no_actual_movement_dialog())

    def test_log_is_saved_per_event_without_credentials(self):
        with tempfile.TemporaryDirectory() as directory:
            service = probe.ProbeService(probe.SimulatedController(), log_path=Path(directory) / "test.jsonl")
            try:
                rows = [json.loads(row) for row in service.log_path.read_text().splitlines()]
                self.assertEqual(rows[0]["event"], "START")
                self.assertIn("utc", rows[0])
                self.assertNotIn("token", rows[0])
            finally:
                service.close()


if __name__ == "__main__":
    unittest.main()
