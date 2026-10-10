"""Network probe tests using a fake controller; never contact hardware."""
import importlib.util
import json
from pathlib import Path
import sys
import threading
import time
import tempfile
from types import SimpleNamespace
from unittest.mock import Mock, patch
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
        self.clock = SimpleNamespace(now=0., on_sleep=None)
        self.clock.monotonic = lambda: self.clock.now
        def sleep(duration):
            self.clock.now += duration
            if self.clock.on_sleep:
                self.clock.on_sleep()
            time.sleep(.001)  # Yield to Stop/API threads without real settling delays.
        self.clock.sleep = sleep
        clock_patch = patch.object(probe, "time", self.clock)
        clock_patch.start()
        self.addCleanup(clock_patch.stop)
        self.controller = probe.SimulatedController()
        self.service = probe.ProbeService(self.controller)
        self.addCleanup(self.service.close)

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
        for payload in [dict(kind="relative", coordinates={"AP": 1.01}),
                        dict(kind="relative", coordinates={"DV": .01}),
                        dict(kind="axis", axis="AP", target_mm=float("nan")),
                        dict(kind="nudge", axis="AP", step_mm=.001, direction=1),
                        dict(kind="absolute", coordinates={"AP": True}),
                        dict(kind="relative", coordinates={"Bregma": .01})]:
            with self.subTest(payload=payload), self.assertRaises(probe.Rejected):
                self.service.submit(dict(command_id=str(time.monotonic_ns()), **payload))
        self.assertEqual(self.controller.position, [0, 0, 0])

    def test_envelope_prevents_accumulated_drift(self):
        self.controller.position[0] = .99
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

    def test_stop_allows_next_command_without_rearming(self):
        self.service.stop()
        self.assertTrue(self.service.status()["ready"])
        self.submit(kind="relative", coordinates={"AP": .01})

    def test_ready_without_arm_or_heartbeat(self):
        self.assertTrue(self.service.status()["ready"])
        self.assertFalse(hasattr(self.service, "heartbeat"))
        self.assertFalse(hasattr(self.service, "arm_local"))
        self.submit(kind="relative", coordinates={"AP": .01})

    def test_fault_blocks_further_commands(self):
        self.service.stop("Hardware fault", fault=True)
        with self.assertRaises(probe.Rejected):
            self.submit(kind="relative", coordinates={"AP": .01})
        self.assertFalse(self.service.status()["ready"])
        self.assertTrue(self.controller.cancelled)

    def test_monitor_stops_outside_envelope(self):
        self.controller.position[0] = 1.2
        self.service.monitor.join(timeout=.3)
        self.assertFalse(self.service.status()["ready"])
        self.assertTrue(self.controller.cancelled)

    def test_failed_forward_does_not_reverse(self):
        self.controller.nudge_axis = lambda *args: (_ for _ in ()).throw(RuntimeError("readout failed"))
        result = self.submit(kind="out_and_back", axis="AP", direction=1, step_mm=.01)
        self.assertEqual(result["state"], "stopped")
        self.assertFalse(self.service.status()["ready"])
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

    def test_timeout_faults(self):
        self.service.timeout = .05
        self.controller.goto_axis_position = lambda *args, **kwargs: None
        result = self.submit(kind="relative", coordinates={"AP": .01})
        self.assertEqual(result["state"], "stopped")
        self.assertFalse(self.service.status()["ready"])

    def test_dv_explicitly_enabled(self):
        self.service.allow_dv = True
        self.submit(kind="relative", coordinates={"DV": -.01})
        self.assertAlmostEqual(self.controller.position[2], -.01)

    def test_default_one_mm_limits_and_full_distance_methods(self):
        self.assertEqual(self.service.status()["bounds"]["AP"], [-1, 1])
        self.assertEqual(self.service.status()["max_move_mm"], 1)
        for method in ("goto", "nudged", "planar"):
            with self.subTest(method=method):
                self.controller.position[:] = [0, 0, 0]
                result = self.submit(kind="relative", coordinates={"AP": 1}, method=method)
                self.assertEqual(result["state"], "completed")
                self.assertAlmostEqual(self.controller.position[0], 1)

    def test_one_mm_out_and_back_and_diagonal_limit(self):
        result = self.submit(kind="out_and_back", axis="ML", step_mm=1, direction=-1)
        self.assertEqual(result["state"], "completed")
        self.assertEqual(self.controller.position, [0, 0, 0])
        with self.assertRaises(probe.Rejected):
            self.submit(kind="relative", coordinates={"AP": .8, "ML": .8})

    def test_one_mm_dv_when_enabled(self):
        self.service.allow_dv = True
        self.assertEqual(self.submit(kind="relative", coordinates={"DV": -1})["state"], "completed")
        self.assertAlmostEqual(self.controller.position[2], -1)

    def test_injector_actions_do_not_move_axes(self):
        for direction in ("up", "down"):
            self.assertEqual(self.submit(kind="injector_step", direction=direction, volume_nl=10)["state"], "completed")
        self.assertEqual(self.submit(kind="injector_inject", volume_nl=20)["state"], "completed")
        self.assertEqual(self.controller.injector_actions, [("up",10),("down",10),("inject",20)])
        self.assertEqual(self.controller.position, [0,0,0])

    def test_injector_out_and_back(self):
        original = self.controller.injector_position
        self.submit(kind="injector_out_and_back", direction="down", volume_nl=50)
        self.assertEqual(self.controller.injector_position, original)
        self.assertEqual(self.controller.injector_actions, [("down",50),("up",50)])

    def test_injector_limits_and_directions(self):
        for payload in [dict(kind="injector_inject",volume_nl=200),
                        dict(kind="injector_step",direction="up",volume_nl=15),
                        dict(kind="injector_step",direction=1,volume_nl=10),
                        dict(kind="injector_step",direction="inject",volume_nl=10),
                        dict(kind="injector_inject",volume_nl=float("nan"))]:
            with self.subTest(payload=payload), self.assertRaises(probe.Rejected):
                self.submit(**payload)
        self.assertEqual(self.controller.injector_actions, [])
        self.service.max_injector_volume = 500
        with self.assertRaises(probe.Rejected):
            self.submit(kind="injector_inject",volume_nl=500)

    def test_injector_duplicate_does_not_repeat_dose(self):
        payload = dict(command_id="dose",kind="injector_inject",volume_nl=100)
        self.service.submit(payload)
        self.service.worker.join(timeout=2)
        self.service.submit(payload)
        self.assertEqual(self.controller.injector_actions, [("inject",100)])

    def test_injector_failed_forward_does_not_reverse(self):
        self.controller.probe_injector_action = lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("injector failed"))
        self.controller.probe_stop_injector = Mock()
        result = self.submit(kind="injector_out_and_back", direction="up", volume_nl=10)
        self.assertEqual(result["state"], "stopped")
        self.controller.probe_stop_injector.assert_called_once()
        self.assertNotIn("INJECTOR_REVERSE_REQUEST", [e["event"] for e in self.service.events])

    def test_stop_and_timeout_cancel_injector(self):
        entered = threading.Event()
        def blocked(action, volume_nl, stop_requested, timeout_seconds):
            entered.set()
            while not stop_requested():
                self.clock.sleep(.02)
            raise RuntimeError("cancelled")
        self.controller.probe_injector_action = blocked
        self.controller.probe_stop_injector = Mock()
        self.service.submit(dict(command_id="busy-injector",kind="injector_step",direction="up",volume_nl=10))
        self.assertTrue(entered.wait(timeout=1))
        with self.assertRaises(probe.Rejected):
            self.service.submit(dict(command_id="axis-while-injecting",kind="relative",coordinates={"AP":.01}))
        self.service.stop()
        self.service.worker.join(timeout=2)
        self.assertFalse(self.service.worker.is_alive())
        self.assertTrue(self.controller.probe_stop_injector.called)
        self.assertTrue(self.service.status()["ready"])
        self.service.timeout = .05
        result = self.submit(kind="injector_inject", volume_nl=10)
        self.assertEqual(result["state"], "stopped")
        self.assertFalse(self.service.status()["ready"])

    def test_direct_backend_requires_zero_setup(self):
        with self.assertRaisesRegex(RuntimeError, "measured zero"):
            probe.real_controller(None, simulate=True)

    def test_three_axis_fine_move_can_exceed_100_increments(self):
        self.service.allow_dv = True
        result = self.submit(kind="relative", coordinates={"AP": .57, "ML": .57, "DV": .57}, method="nudged")
        self.assertEqual(result["state"], "completed")
        for value in self.controller.position:
            self.assertAlmostEqual(value, .57)

    def test_arrival_requires_native_controls_ready_and_stable_readings(self):
        self.service.operation = {"axes": ["AP"]}
        self.controller.position[0] = .01
        self.controller.motion_controls_ready = lambda axes: self.clock.now >= .3
        self.service._wait((.01, 0, 0), .006, deadline=2)
        self.assertGreaterEqual(self.clock.now, .5)
        self.assertEqual(self.service.events[-1]["settled_seconds"], .2)

    def test_transient_matching_reading_does_not_report_arrival(self):
        self.service.operation = {"axes": ["AP"]}
        self.controller.position[0] = .01
        self.clock.on_sleep = lambda: self.controller.position.__setitem__(0, .03)
        with self.assertRaisesRegex(RuntimeError, "timed out"):
            self.service._wait((.01, 0, 0), .006, deadline=.5)
        self.assertNotIn("ARRIVED", [e["event"] for e in self.service.events])

    def test_arrival_timer_resets_while_display_still_changes(self):
        self.service.operation = {"axes": ["AP"]}
        self.controller.position[0] = .005
        def update():
            if self.clock.now >= .15:
                self.controller.position[0] = .01
        self.clock.on_sleep = update
        self.service._wait((.01, 0, 0), .006, deadline=2)
        self.assertGreaterEqual(self.clock.now, .35)

    def test_busy_controls_never_report_arrival(self):
        self.service.operation = {"axes": ["AP"]}
        self.controller.position[0] = .01
        self.controller.motion_controls_ready = lambda axes: False
        with self.assertRaises(RuntimeError):
            self.service._wait((.01, 0, 0), .006, deadline=.5)
        self.assertNotIn("ARRIVED", [e["event"] for e in self.service.events])

    def test_small_goto_through_actual_adapter_and_probe(self):
        from test_movement import SimController, controller as native_module
        class NativeFake(SimController):
            def safety_check(self):
                pass
            def _is_control_enabled(self, control_id):
                return True
        native = NativeFake()
        with patch.object(native_module, "time", self.clock):
            service = probe.ProbeService(native)
            try:
                service.submit(dict(command_id="actual-adapter", kind="relative", coordinates={"AP": .01}))
                service.worker.join(timeout=2)
                self.assertEqual(service.status()["operation"]["state"], "completed")
                self.assertAlmostEqual(native.position[0], 30.01)
                self.assertEqual(native.clicks.count(native_module.GOTO_ID), 1)
            finally:
                service.close()

    def test_token_free_http_rejects_browser_and_bad_host(self):
        server = ThreadingHTTPServer(("127.0.0.1", 0), probe.handler(self.service))
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(server.server_close)
        self.addCleanup(server.shutdown)
        url = f"http://127.0.0.1:{server.server_port}"

        def call(path, payload=None, origin=None, host=None, content_type="application/json"):
            headers = {"Content-Type": content_type}
            if host:
                headers["Host"] = host
            if origin:
                headers["Origin"] = origin
            req = urllib.request.Request(url + path, headers=headers,
                                         data=None if payload is None else json.dumps(payload).encode())
            opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
            with opener.open(req, timeout=2) as response:
                return json.load(response)

        with self.assertRaises(urllib.error.HTTPError) as caught:
            call("/status", host="malicious.example")
        self.assertEqual(caught.exception.code, 403)
        with self.assertRaises(urllib.error.HTTPError) as caught:
            call("/stop", {}, origin="https://example.com")
        self.assertEqual(caught.exception.code, 403)
        with self.assertRaises(urllib.error.HTTPError) as caught:
            call("/move", {}, content_type="text/plain")
        self.assertEqual(caught.exception.code, 415)
        self.assertTrue(call("/status")["ready"])
        with self.assertRaises(urllib.error.HTTPError) as caught:
            call("/heartbeat", {})
        self.assertEqual(caught.exception.code, 404)
        call("/move", dict(command_id="http", kind="nudge", axis="AP", direction=1, step_mm=.01))
        self.service.worker.join(timeout=2)
        self.assertEqual(call("/status")["operation"]["state"], "completed")
        call("/move", dict(command_id="http-injector", kind="injector_inject", volume_nl=10))
        self.service.worker.join(timeout=2)
        self.assertEqual(call("/status")["operation"]["state"], "completed")
        self.assertEqual(self.controller.injector_actions, [("inject", 10)])
        self.assertTrue(call("/events"))
        call("/stop", {})
        self.assertTrue(call("/status")["ready"])
        events = self.service.events
        incoming = [row for row in events if row["event"] == "INCOMING_MOVE"]
        self.assertEqual(incoming[0]["command"]["command_id"], "http")
        self.assertEqual(incoming[1]["command"]["volume_nl"], 10)
        self.assertIn("peer", incoming[0])
        self.assertTrue(any(row["event"] == "HTTP_RESULT" and row["status"] == 403 for row in events))

    def test_rejected_movement_is_visible_in_audit(self):
        server = ThreadingHTTPServer(("127.0.0.1", 0), probe.handler(self.service))
        threading.Thread(target=server.serve_forever, daemon=True).start()
        self.addCleanup(server.server_close)
        self.addCleanup(server.shutdown)
        payload = dict(command_id="bad", kind="relative", coordinates={"DV": .01}, password="do-not-log")
        req = urllib.request.Request(f"http://127.0.0.1:{server.server_port}/move",
                                     data=json.dumps(payload).encode(), headers={"Content-Type": "application/json"})
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with self.assertRaises(urllib.error.HTTPError) as caught:
            opener.open(req, timeout=2)
        self.assertEqual(caught.exception.code, 400)
        self.assertTrue(any(row["event"] == "INCOMING_MOVE" for row in self.service.events))
        self.assertTrue(any(row["event"] == "HTTP_RESULT" and row["status"] == 400 for row in self.service.events))
        self.assertNotIn("do-not-log", json.dumps(self.service.events))

    def test_gui_event_filter_keeps_commands_and_failures(self):
        visible = gui_helpers["event_visible"]
        self.assertFalse(visible(dict(event="HTTP_RESULT", endpoint="/status", status=200)))
        self.assertTrue(visible(dict(event="HTTP_RESULT", endpoint="/status", status=400)))
        self.assertTrue(visible(dict(event="HTTP_REQUEST", endpoint="/move")))
        self.assertTrue(visible(dict(event="HTTP_REQUEST", endpoint="/status"), True))
        row = dict(sequence=1, utc="2026-10-09T10:00:00+00:00", event="INCOMING_MOVE", command={"kind":"nudge"})
        self.assertIn('"kind": "nudge"', gui_helpers["format_event"](row))

    def test_direct_adapter_uses_public_api_for_axis_and_piston(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory)
            setup=root/'setup.json'
            setup.write_text(json.dumps(dict(calibration=dict(axis_zero_counts=dict(AP=105280,ML=75864,DV=41767),
                piston_3000_count=-8572,anchor_backlash=dict(AP=0,ML=261,DV=0,PISTON=0)),
                verified_backlash=dict(AP=0,ML=261,DV=0,PISTON=0))))
            with patch.dict('os.environ', {'LOCALAPPDATA':directory}):
                controller=probe.real_controller(setup,allow_dv=True,allow_piston=True,simulate=True)
            try:
                controller.prepare_motion()
                controller.goto_axis_position(.01,.01,.01)
                position=controller.get_current_axis_position()
                self.assertAlmostEqual(position[0],.01,delta=1/5225)
                controller.probe_injector_action('up',10,lambda:False,5)
                self.assertAlmostEqual(controller.read_injectomate_calibrate_scale_nl(),3010,delta=1/161.36)
                with self.assertRaisesRegex(ValueError,'unvalidated'):
                    controller.probe_injector_action('inject',10,lambda:False,5)
                self.assertFalse(controller.confirm_below_skull_warning())
            finally:controller.close()

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
