"""Hardware-free regression tests for mechanical movement and GUI references.

Run: python -B -m unittest discover -s tests -v
The Windows API is stubbed during controller import. GUI classes are loaded
from their original AST without Qt startup/imports; their method bodies are
unchanged. These tests do not contact StereoDrive or issue USB commands.
"""
import ast
import ctypes
import importlib.util
import math
import json
import sys
import threading
import time
import unittest
from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch


ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("movement_test_controller", ROOT / "tools/stereodrive_controller.py")
controller = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = controller
with patch.object(ctypes, "WinDLL", return_value=Mock(), create=True), \
        patch.object(ctypes, "WINFUNCTYPE", ctypes.CFUNCTYPE, create=True):
    spec.loader.exec_module(controller)


def point(x=0.0, y=None):
    if y is None and hasattr(x, "x"):
        x, y = x.x(), x.y()
    y = 0.0 if y is None else y
    return SimpleNamespace(x=lambda: x, y=lambda: y)


def load_gui():
    path = ROOT / "tools/craniotomy_qt.py"
    tree = ast.parse(path.read_text())
    classes = {"ProjectionWidget", "CraniotomyWindow", "SeedPoint", "StoredLocation", "InjectionSite", "InjectionProtocolSettings", "CraniotomyConfig"}
    nodes = [node for node in tree.body if isinstance(node, ast.ClassDef) and node.name in classes]
    module = ast.Module(body=[ast.ImportFrom(module="__future__", names=[ast.alias(name="annotations")], level=0)] + nodes,
                        type_ignores=[])
    widget = type("Widget", (), {name: lambda *args: None for name in
                  ("mousePressEvent", "mouseReleaseEvent", "mouseMoveEvent", "mouseDoubleClickEvent")})
    namespace = dict(dataclass=dataclass, QMainWindow=object, QWidget=widget, QPointF=point, Signal=lambda *args: Mock(),
                     StereoDriveError=controller.StereoDriveError, math=math, threading=threading,
                     time=time, json=json, Path=Path, QMessageBox=Mock(Yes=1, No=0, Cancel=2),
                     QApplication=Mock(), Qt=Mock(), DEFAULT_INJECTION_VOLUME_NL=100,
                     MOVE_SPEED_OPTIONS_MM=controller.NUDGE_STEP_OPTIONS_MM)
    exec(compile(ast.fix_missing_locations(module), str(path), "exec"), namespace)
    return namespace


GUI = load_gui()
Window = GUI["CraniotomyWindow"]
Site = GUI["InjectionSite"]


class Clock:
    def __init__(self):
        self.now = 0.0
        self.on_sleep = None

    def monotonic(self):
        return self.now

    def sleep(self, duration):
        self.now += max(duration, 0.001)
        if self.on_sleep:
            self.on_sleep()


class SimController(controller.StereoDriveController):
    def _find_main_window(self):
        return 1

    def __init__(self):
        super().__init__()
        self.position = [30.0, 31.0, 19.0]
        self.native_origin = [29.0, 30.0, 18.0]
        self.fields = {}
        self.steps = {}
        self.clicks = []
        self.history = []
        self.overshoot_once = False
        self.stalled_dv_once = False

    def _control_handle(self, control_id, **kwargs):
        return control_id

    def _parse_float(self, control_id):
        if control_id in (1138, 1139, 1140):
            return self.position[control_id - 1138]
        if control_id in (1144, 1145, 1146):
            return self.position[control_id - 1144] - self.native_origin[control_id - 1144]
        raise AssertionError(control_id)

    def _get_text(self, control_id):
        return self.fields.get(control_id, "")

    def _set_edit_control_text(self, control_id, text):
        self.fields[control_id] = text

    def set_nudge_step(self, axis, step):
        self.steps[axis] = step

    def _click(self, control_id):
        self.clicks.append(control_id)
        buttons = {1102: (0, 1), 1103: (0, -1), 1105: (1, 1),
                   1104: (1, -1), 1107: (2, 1), 1106: (2, -1)}
        if control_id in buttons:
            index, direction = buttons[control_id]
            distance = self.steps[("AP", "ML", "DV")[index]]
            if self.overshoot_once:
                distance += 0.01
                self.overshoot_once = False
            self.position[index] += direction * distance
        elif control_id == controller.GOTO_ID:
            for index, field in enumerate((1141, 1142, 1143)):
                if self.stalled_dv_once and index == 2:
                    self.fields[field] = ""
                    self.stalled_dv_once = False
                elif self.fields.get(field):
                    self.position[index] = float(self.fields[field])
        self.history.append(tuple(self.position))

    def confirm_below_skull_warning(self, **kwargs):
        return False

    def confirm_no_actual_movement_dialog(self, **kwargs):
        return False


def window():
    w = Window.__new__(Window)
    w.controller = SimController()
    w.bregma_axis = (30.0, 31.0, 19.0)
    w.coordinate_mode = "bregma"
    w.craniotomy_coordinate_system = "bregma"
    w.injection_sites_coordinate_system = "bregma"
    w.validation_clearance_mm = 0.5
    w.home_axis = w.work_axis = None
    w.injection_sites = []
    w.drill_thread = w.injection_thread = w.benchmark_thread = w.usb_probe_thread = None
    w.validation_move_active = w.validation_modal_active = False
    w.validation_move_cancel_callback = None
    w.nudge_all_sites_active = False
    w.move_speed_step_mm = 0.05
    w.quick_locations = {}
    w.injection_stop_requested = threading.Event()
    w.injection_pause_requested = threading.Event()
    w.drill_stop_requested = threading.Event()
    w.drill_pause_requested = threading.Event()
    for name in ("injection_progress", "injection_site_progress", "start_injection_btn",
                 "pause_injection_btn", "injection_finished_signal", "sequence_step_signal",
                 "active_injection_site_signal", "status_signal", "injection_progress_signal"):
        setattr(w, name, Mock())
    w.set_status = Mock()
    return w


class MechanicalMovementTests(unittest.TestCase):
    def setUp(self):
        self.c = SimController()
        self.clock = Clock()
        self.clock_patch = patch.object(controller, "time", self.clock)
        self.clock_patch.start()
        self.addCleanup(self.clock_patch.stop)

    def test_axis_readouts_ignore_native_reference(self):
        self.assertEqual(self.c.get_current_axis_position(), (30., 31., 19.))
        self.assertEqual(self.c.get_current_axis("DV"), 19.)
        self.assertEqual(self.c.get_current_position(), (1., 1., 1.))

    def test_fine_dv_targets_mechanical_axis(self):
        self.c.move_axis_to_target("DV", 19.5)
        self.assertAlmostEqual(self.c.position[2], 19.5)

    def test_continuous_xyz_uses_same_axis_reference(self):
        self.c.move_to_position_nudged(30.1, 31.05, 19.02, step_mm=.01)
        for actual, target in zip(self.c.position, (30.1, 31.05, 19.02)):
            self.assertAlmostEqual(actual, target, places=3)

    def test_overshoot_is_corrected_before_success(self):
        self.c.overshoot_once = True
        self.c.move_axis_to_target("DV", 19.005, step_mm=.005)
        self.assertLessEqual(abs(self.c.position[2] - 19.005), .003)

    def test_cancel_during_entry_does_not_issue_goto(self):
        self.clock.on_sleep = self.c.stop
        with self.assertRaisesRegex(controller.StereoDriveError, "cancelled"):
            self.c.goto_axis_position(32., 33., 20.)
        self.assertNotIn(controller.GOTO_ID, self.c.clicks)

    def test_cancel_during_goto_delay_does_not_restart(self):
        def cancel():
            if self.clock.now >= .5:
                self.c.stop()
        self.clock.on_sleep = cancel
        with self.assertRaisesRegex(controller.StereoDriveError, "cancelled"):
            self.c.goto_axis_position(32., 33., 20.)
        self.assertNotIn(controller.GOTO_ID, self.c.clicks)

    def test_stop_latches_until_new_owned_operation(self):
        self.c.stop()
        with self.assertRaises(controller.StereoDriveError):
            self.c.nudge_axis("DV", True)
        self.c.prepare_motion()
        self.c.set_nudge_step("DV", .005)
        self.c.nudge_axis("DV", True)
        self.assertAlmostEqual(self.c.position[2], 19.005)

    def test_missing_dv_is_recovered_in_shared_arrival_wait(self):
        self.c.stalled_dv_once = True
        self.c.goto_axis_position(32., 33., 20.)
        self.assertEqual(self.c.position[2], 19.)
        self.c.wait_for_axis_position(32., 33., 20.)
        self.assertEqual(self.c.position, [32., 33., 20.])

    def test_recovery_never_reissues_reached_dv(self):
        self.assertFalse(self.c.rearm_axis_dv_and_goto_if_needed(19.))
        self.assertEqual(self.c.clicks, [])

    def test_wait_cancelled_before_retry(self):
        self.c.stop()
        with self.assertRaisesRegex(controller.StereoDriveError, "cancelled"):
            self.c.wait_for_axis_position(32., 33., 20.)
        self.assertNotIn(controller.GOTO_ID, self.c.clicks)

    def test_non_finite_targets_are_rejected_before_motor_commands(self):
        for value in (float("nan"), float("inf")):
            with self.assertRaisesRegex(controller.StereoDriveError, "finite"):
                self.c.move_axis_to_target("DV", value)
            with self.assertRaisesRegex(controller.StereoDriveError, "finite"):
                self.c.goto_axis_position(30., 31., value)
        self.assertEqual(self.c.clicks, [])

    def test_named_motion_does_not_claim_no_movement_is_arrival(self):
        with self.assertRaisesRegex(controller.StereoDriveError, "arrival cannot be verified"):
            self.c.wait_for_named_motion(tuple(self.c.position))

    def test_named_motion_waits_for_observed_motion_to_settle(self):
        initial = tuple(self.c.position)
        self.clock.on_sleep = lambda: self.c.position.__setitem__(0, 32.)
        self.c.wait_for_named_motion(initial)
        self.assertGreaterEqual(self.clock.now, 1.)

    def test_stop_confirmation_detects_slow_accumulated_drift(self):
        self.clock.on_sleep = lambda: self.c.position.__setitem__(2, self.c.position[2] + .001)
        with self.assertRaisesRegex(controller.StereoDriveError, "Could not confirm"):
            self.c.wait_until_stopped(timeout_seconds=1.)


class GuiCoordinateTests(unittest.TestCase):
    def setUp(self):
        self.w = window()

    def test_bregma_button_targets_origin_in_axis_display(self):
        self.w.coordinate_mode = "axis"
        self.w._move_to_axis_position_with_progress = Mock(return_value=True)
        self.w.goto_bregma()
        self.assertEqual(self.w._move_to_axis_position_with_progress.call_args.args[0], (30., 31., 19.))

    def test_home_and_work_use_fixed_axis_targets_after_anchor_change(self):
        self.w.home_axis = (10., 11., 12.)
        self.w.work_axis = (20., 21., 22.)
        self.w.bregma_axis = (34., 35., 23.)
        self.w._move_to_axis_position_with_progress = Mock(return_value=True)
        self.w.goto_home()
        self.assertEqual(self.w._move_to_axis_position_with_progress.call_args.args[0], (10., 11., 12.))
        self.w.goto_work()
        self.assertEqual(self.w._move_to_axis_position_with_progress.call_args.args[0], (20., 21., 22.))
        self.assertEqual(self.w.controller.clicks, [])

    def test_setting_home_captures_axis_even_in_bregma_mode(self):
        self.w._save_general_settings = Mock()
        self.w.set_persistent_axis_location("home")
        self.assertEqual(self.w.home_axis, (30., 31., 19.))
        self.assertTrue(self.w._save_general_settings.called)

    def test_unset_home_does_not_use_native_home(self):
        self.w.controller.goto_home = Mock()
        self.w.goto_home()
        self.assertFalse(self.w.controller.goto_home.called)
        self.assertEqual(self.w.controller.clicks, [])

    def test_permanent_positions_and_validation_height_settings_roundtrip(self):
        self.w.home_axis, self.w.work_axis = (10., 11., 12.), (20., 21., 22.)
        self.w.validation_clearance_mm = 1.25
        self.w.movement_key_edits = self.w.movement_key_bindings = self.w.syringe_key_bindings = {}
        self.w.anchor_axis = self.w.anchor_bregma = None
        self.w.recent_injection_grid_configs = []
        self.w._config_root_dir = lambda: Mock()
        self.w._general_settings_path = lambda: Mock(exists=lambda: True)
        self.w._write_config_file = Mock()
        self.w._save_general_settings()
        saved = json.loads(json.dumps(self.w._write_config_file.call_args.args[1]))
        self.w.home_axis = self.w.work_axis = None
        self.w.validation_clearance_mm = .5
        self.w._read_config_file = lambda path: saved
        self.w.validation_clearance_edit = Mock()
        self.w.update_coordinate_mode_buttons = Mock()
        self.w.top_view = self.w.injection_sites_view = Mock()
        self.w._load_general_settings()
        self.assertEqual(self.w.home_axis, (10., 11., 12.))
        self.assertEqual(self.w.work_axis, (20., 21., 22.))
        self.assertEqual(self.w.validation_clearance_mm, 1.25)

    def test_cancelled_navigation_does_not_report_arrival(self):
        self.w._move_to_axis_position_with_progress = Mock(return_value=False)
        self.w.goto_bregma()
        self.assertFalse(self.w.set_status.called)

    def test_manual_site_capture_uses_bregma_in_axis_display(self):
        self.w.coordinate_mode = "axis"
        self.w.controller.position = [31., 33., 19.2]
        self.w.refresh_injection_sites_list = Mock()
        self.w.add_injection_site()
        site = self.w.injection_sites[0]
        self.assertEqual((site.ap, site.ml), (1., 2.))
        self.assertAlmostEqual(site.dv, .2)

    def test_map_click_adds_unvalidated_bregma_site_without_moving(self):
        self.w.add_sites_on_map_checkbox = Mock(isChecked=lambda: True)
        self.w.refresh_injection_sites_list = Mock()
        self.w.injection_sites_list = Mock()
        self.w._autosave_project_session = Mock()
        self.w.add_injection_site_from_map(2., 1.)
        self.assertEqual(self.w.injection_sites, [Site(1., 2., None, True)])
        self.assertEqual(self.w.controller.clicks, [])
        self.assertTrue(self.w._autosave_project_session.called)

    def test_axis_map_click_is_stored_relative_to_gui_bregma(self):
        self.w.coordinate_mode = "axis"
        self.w.add_sites_on_map_checkbox = Mock(isChecked=lambda: True)
        self.w.refresh_injection_sites_list = Mock()
        self.w.injection_sites_list = Mock()
        self.w._autosave_project_session = Mock()
        self.w.add_injection_site_from_map(33., 31.)
        self.assertEqual(self.w.injection_sites, [Site(1., 2., None, True)])

    def test_busy_operation_blocks_map_site_placement(self):
        self.w.add_sites_on_map_checkbox = Mock(isChecked=lambda: True)
        self.w.injection_thread = Mock(is_alive=lambda: True)
        self.w.add_injection_site_from_map(2., 1.)
        self.assertEqual(self.w.injection_sites, [])

    def test_map_placement_requires_gui_bregma(self):
        self.w.bregma_axis = None
        self.w.add_sites_on_map_checkbox = Mock()
        self.w.injection_sites_view = Mock()
        self.w.set_add_sites_on_map(True)
        self.w.add_sites_on_map_checkbox.setChecked.assert_called_once_with(False)
        self.w.injection_sites_view.set_add_site_mode.assert_called_once_with(False)

    def test_quick_positions_are_mode_independent(self):
        self.w.coordinate_mode = "axis"
        self.w.controller.position = [31., 33., 19.2]
        self.w.set_quick_location("A")
        self.w.bregma_axis = (34., 35., 23.)
        self.w._move_to_axis_position_with_progress = Mock()
        self.w.goto_quick_location("A")
        self.assertEqual(self.w._move_to_axis_position_with_progress.call_args.args[0][:2], (35., 37.))
        self.assertAlmostEqual(self.w._move_to_axis_position_with_progress.call_args.args[0][2], 23.2)

    def test_anchor_recalibration_changes_origin_only(self):
        self.w.anchor_bregma = (1., 2., .3)
        self.w.anchor_axis = (31., 33., 19.3)
        self.w.controller.position = [35., 37., 23.3]
        self.w.injection_sites = [Site(2., 3., .1)]
        self.w._save_general_settings = Mock()
        self.w.update_coordinate_mode_buttons = Mock()
        self.w.refresh_injection_sites_list = Mock()
        self.w.redraw_views = Mock()
        self.w.top_view = self.w.injection_sites_view = Mock()
        self.w.current_ap_label = self.w.current_ml_label = self.w.current_dv_label = Mock()
        with patch.object(GUI["QMessageBox"], "warning", return_value=1):
            self.w.at_anchor()
        self.assertEqual(self.w.bregma_axis[:2], (34., 35.))
        self.assertAlmostEqual(self.w.bregma_axis[2], 23.)
        self.assertEqual(self.w.injection_sites, [Site(2., 3., .1)])

    def test_injection_snapshot_uses_new_origin_and_survives_mode_change(self):
        original = Site(1., 2., .2)
        self.w.bregma_axis = (34., 35., 23.)
        self.w.coordinate_mode = "axis"
        with patch.object(threading, "Thread") as thread:
            self.w._start_injection_sequence([original], Mock(), [100], False, 50, 0, 1, "Ready")
            sent = thread.call_args.kwargs["args"][0][0]
        self.assertEqual((sent.ap, sent.ml, sent.dv), (35., 37., 23.2))
        self.w.bregma_axis = (90., 90., 90.)
        self.assertEqual(sent.dv, 23.2)
        self.assertEqual(original.dv, .2)

    def test_clearance_retracts_before_xy_and_never_lowers_first(self):
        self.w.controller.position = [30., 31., 21.]
        self.assertEqual(self.w._axis_clearance_path((33., 34., 19.), 18.5),
                         [(30., 31., 18.5), (33., 34., 18.5), (33., 34., 19.)])
        self.w.controller.position[2] = 17.
        self.assertEqual(self.w._axis_clearance_path((33., 34., 19.), 18.5)[0][2], 17.)

    def test_approach_waits_at_every_clearance_stage(self):
        commands = []
        self.w.controller.goto_axis_position = lambda *p, **kw: commands.append(("go", p))
        self.w.controller.wait_for_axis_position = lambda *p, **kw: commands.append(("wait", p))
        self.w._approach_axis_position((33., 34., 19.), 18.5, lambda: False)
        self.assertEqual([kind for kind, _ in commands], ["go", "wait"] * 3)
        self.assertEqual(commands[0][1], (30., 31., 18.5))

    def test_validation_uses_custom_height_above_recalibrated_bregma(self):
        self.w.bregma_axis = (34., 35., 23.)
        self.w.validation_clearance_mm = 1.5
        self.w._move_through_axis_positions_with_progress = Mock(return_value=True)
        self.w._move_to_injection_site_for_validation(Site(1., 2., None, True))
        path = self.w._move_through_axis_positions_with_progress.call_args.args[0]
        self.assertEqual(path[-1], (35., 37., 21.5))
        self.assertEqual(path[0][:2], (30., 31.))

    def test_validation_height_change_is_saved(self):
        self.w.validation_clearance_edit = Mock(value=lambda: 1.25)
        self.w._save_general_settings = Mock()
        self.w.save_validation_clearance()
        self.assertEqual(self.w.validation_clearance_mm, 1.25)
        self.assertTrue(self.w._save_general_settings.called)

    def test_right_click_validation_runs_only_selected_site(self):
        self.w.injection_sites = [Site(1., 2., None), Site(3., 4., None)]
        self.w._run_injection_site_validation = Mock()
        self.w.validate_injection_site(1)
        self.w._run_injection_site_validation.assert_called_once_with(1, validate_all_sites=True, single_site=True)

    def test_single_site_validation_captures_surface_without_advancing(self):
        self.w.injection_sites = [Site(1., 2., None), Site(3., 4., None)]
        self.w.controller.position = [33.1, 35.2, 19.3]
        self.w.refresh_injection_sites_list = Mock()
        self.w._move_to_injection_site_for_validation = Mock(return_value=True)
        self.w._validation_dialog = Mock(return_value=("validate", None))
        self.w._run_injection_site_validation(1, True, single_site=True)
        self.w._move_to_injection_site_for_validation.assert_called_once()
        self.assertIsNone(self.w.injection_sites[0].dv)
        self.assertAlmostEqual(self.w.injection_sites[1].ap, 3.1)
        self.assertAlmostEqual(self.w.injection_sites[1].ml, 4.2)
        self.assertAlmostEqual(self.w.injection_sites[1].dv, .3)

    def test_single_site_skip_keeps_unvalidated_position(self):
        original = Site(3., 4., None, True)
        self.w.injection_sites = [original, Site(5., 6., None)]
        self.w.refresh_injection_sites_list = Mock()
        self.w._move_to_injection_site_for_validation = Mock(return_value=True)
        self.w._validation_dialog = Mock(return_value=("next", None))
        self.w._run_injection_site_validation(0, True, single_site=True)
        self.assertEqual(self.w.injection_sites[0], original)
        self.w._move_to_injection_site_for_validation.assert_called_once()

    def test_context_menu_routes_to_clicked_row(self):
        item = Mock()
        action = Mock()
        menu = Mock(addAction=lambda text: action, exec=lambda position: action)
        self.w.injection_sites_list = Mock(itemAt=lambda position: item, row=lambda candidate: 2)
        self.w.validate_injection_site = Mock()
        with patch.dict(GUI, QMenu=lambda parent: menu):
            self.w.show_injection_site_context_menu(point(10., 20.))
        self.w.validate_injection_site.assert_called_once_with(2)

    def test_busy_worker_blocks_manual_nudges_and_recalibration(self):
        self.w.injection_thread = Mock(is_alive=lambda: True)
        self.w.controller.nudge_axis = Mock()
        self.w.keyboard_nudge("DV", True, "Down")
        self.assertFalse(self.w.controller.nudge_axis.called)
        before = self.w.bregma_axis
        self.w.set_local_bregma()
        self.assertEqual(self.w.bregma_axis, before)

    def test_validation_dialog_keeps_manual_movement_available(self):
        self.w.validation_modal_active = True
        self.w._focus_is_editable = lambda: False
        self.w.update_move_speed_label = Mock()
        self.w.refresh_live_position = Mock()
        self.w.keyboard_nudge("DV", True, "Down")
        self.assertAlmostEqual(self.w.controller.position[2], 19.05)

    def test_global_stop_cancels_both_workers_and_hardware(self):
        self.w.injection_thread = Mock(is_alive=lambda: True)
        self.w.controller.stop_injectomate_motion = Mock()
        self.w.stop_motion()
        self.assertTrue(self.w.drill_stop_requested.is_set())
        self.assertTrue(self.w.injection_stop_requested.is_set())
        self.assertTrue(self.w.controller._motion_cancelled.is_set())
        self.assertTrue(self.w.controller.stop_injectomate_motion.called)

    def test_project_map_uses_axis_coordinates_in_axis_mode(self):
        self.w.coordinate_mode = "axis"
        self.assertEqual(self.w._project_map_position(1., 2., "bregma"), (31., 33.))

    def test_legacy_ambiguous_sites_cannot_start(self):
        self.w.injection_sites_coordinate_system = "unknown"
        with self.assertRaisesRegex(controller.StereoDriveError, "ambiguous"):
            self.w._active_injection_sites()

    def test_cancelled_clear_preserves_legacy_frame_guard(self):
        self.w.craniotomy_coordinate_system = "unknown"
        self.w.seeds = [GUI["SeedPoint"](0, 0., 1., 2., .1)]
        self.w.trajectory = [(1., 2., .1)]
        with patch.object(GUI["QMessageBox"], "warning", return_value=2):
            self.w.clear_craniotomy()
        self.assertEqual(self.w.craniotomy_coordinate_system, "unknown")
        self.assertEqual(self.w.trajectory, [(1., 2., .1)])

    def test_injection_depth_loop_uses_frozen_mechanical_targets(self):
        settings = GUI["InjectionProtocolSettings"](0, 0., 100., .2, 1000., .1, 0.)
        self.w.controller.position = [35., 37., 22.1]
        self.w.coordinate_mode = "axis"
        clock = Clock()
        with patch.object(controller, "time", clock), patch.dict(GUI, time=clock):
            self.w._run_protocol_at_site(Site(35., 37., 23.1), settings, [],
                                        dict(approach=0, advance=1, retract=2, pause=3, return_=4, **{"return": 4}), 1, 1)
        deepest = max(position[2] for position in self.w.controller.history)
        self.assertAlmostEqual(deepest, 23.4, places=2)
        self.assertAlmostEqual(self.w.controller.position[2], 22.1)
        self.assertTrue(any(abs(position[2] - 23.3) < .003 for position in self.w.controller.history))

    def test_drilling_snapshots_all_targets_with_recalibrated_origin(self):
        self.w.coordinate_mode = "axis"
        self.w.bregma_axis = (34., 35., 23.)
        self.w.trajectory = [(1., 2., .1), (1.1, 2., .1), (1., 2., .1)]
        self.w.seeds = [GUI["SeedPoint"](0, 0., 1., 2., .1), GUI["SeedPoint"](1, 180., 1.1, 2., .1)]
        self.w.drilled_depths = [0.] * 3
        self.w.frozen_points = [False] * 3
        self.w.current_target_depth_mm = .1
        self.w.drill_depth = Mock(value=lambda: .2)
        self.w.round_time_seconds = Mock(value=lambda: 1.)
        self.w.cut_offset = Mock(value=lambda: 0.)
        self.w.mid_ap = Mock(value=lambda: 1.)
        self.w.mid_ml = Mock(value=lambda: 2.)
        self.w.start_round_btn = Mock()
        self.w.update_current_target_depth_label = Mock()
        with patch.object(threading, "Thread") as thread, \
                patch.object(GUI["QMessageBox"], "question", return_value=1):
            self.w.start_drilling_round()
            args = thread.call_args.kwargs["args"]
        self.assertEqual(args[0][0], (35., 37., 23.1))
        self.assertEqual(args[-1], (35., 37., 21.1))
        self.assertEqual(self.w.trajectory[0], (1., 2., .1))

    def test_drilling_declined_confirmation_never_starts_motion(self):
        self.w.trajectory = [(1., 2., .1)]
        self.w.seeds = [GUI["SeedPoint"](0, 0., 1., 2., .1), GUI["SeedPoint"](1, 180., 1., 2., .1)]
        self.w.drill_depth = Mock(value=lambda: .2)
        self.w.current_target_depth_mm = .1
        self.w.update_current_target_depth_label = Mock()
        self.w.controller.prepare_motion = Mock()
        with patch.object(GUI["QMessageBox"], "question", return_value=0) as question, \
                patch.object(threading, "Thread") as thread:
            self.w.start_drilling_round()
        self.assertTrue(question.called)
        self.assertFalse(self.w.controller.prepare_motion.called)
        self.assertFalse(thread.called)
        self.assertEqual(self.w.controller.clicks, [])

    def test_pause_during_active_drilling_does_not_ask_for_drill_confirmation(self):
        self.w.drill_thread = Mock(is_alive=lambda: True)
        with patch.object(GUI["QMessageBox"], "question") as question:
            self.w.start_drilling_round()
        self.assertTrue(self.w.drill_pause_requested.is_set())
        self.assertFalse(question.called)

    def test_continuous_drilling_segment_passes_axis_xyz(self):
        self.w.controller.position = [35., 37., 23.1]
        self.w.skull_thickness_mm = Mock(value=lambda: .25)
        self.w.drilled_depths = [0., 0.]
        clock = Clock()
        with patch.object(controller, "time", clock), patch.dict(GUI, time=clock):
            count = self.w._follow_continuous_round_segment(
                [(35., 37., 23.1), (35.05, 37.05, 23.12)], 0, 1, .1, .005,
                .003, .003, 0, 20, 0., 0., 2)
        self.assertGreater(count, 0)
        for actual, target in zip(self.w.controller.position, (35.05, 37.05, 23.22)):
            self.assertAlmostEqual(actual, target, places=2)

    def test_close_stops_hardware_and_waits_for_live_worker(self):
        self.w.injection_thread = Mock(is_alive=lambda: True)
        self.w.controller.stop_injectomate_motion = Mock()
        self.w.controller.wait_until_stopped = Mock()
        event = Mock()
        self.w.closeEvent(event)
        self.assertTrue(self.w.controller._motion_cancelled.is_set())
        self.assertTrue(self.w.injection_stop_requested.is_set())
        self.assertTrue(self.w.controller.wait_until_stopped.called)
        self.assertTrue(event.ignore.called)

    def test_new_session_restores_explicit_frames_and_old_session_is_blocked(self):
        self.w.mid_ap = self.w.mid_ml = Mock(value=lambda: 0.)
        self.w.current_seed_spin = Mock()
        self.w._initial_target_depth = lambda: .1
        self.w.update_coordinate_mode_buttons = Mock()
        self.w.top_view = self.w.injection_sites_view = Mock()
        self.w.update_seed_selector_label = Mock()
        self.w.update_current_target_depth_label = Mock()
        self.w.refresh_injection_sites_list = Mock()
        self.w.redraw_views = Mock()
        self.w.zoom_mode_combo = self.w.injection_sites_zoom_combo = Mock(currentIndex=lambda: 0)
        self.w.overlay_combo = Mock(count=lambda: 0)
        self.w.tabs = Mock(count=lambda: 2, currentIndex=lambda: 0)
        payload = dict(bregma_axis=[34., 35., 23.], coordinate_mode="axis",
                       craniotomy_coordinate_system="bregma", injection_sites_coordinate_system="bregma",
                       injection_sites=[dict(ap=1., ml=2., dv=.1)])
        self.w._restore_project_session(payload)
        self.assertEqual(self.w._active_injection_sites(), [Site(1., 2., .1)])
        del payload["injection_sites_coordinate_system"]
        self.w._restore_project_session(payload)
        with self.assertRaisesRegex(controller.StereoDriveError, "ambiguous"):
            self.w._active_injection_sites()

    def test_invalid_restored_origin_is_rejected(self):
        self.assertIsNone(self.w._session_axis([30., 31., float("nan")]))


class MapInteractionTests(unittest.TestCase):
    def setUp(self):
        self.view = GUI["ProjectionWidget"].__new__(GUI["ProjectionWidget"])
        self.view._coordinate_bounds = (-2., 2., -4., 4.)
        self.view._draw_rect = SimpleNamespace(left=lambda: 50., top=lambda: 24., width=lambda: 200.,
            height=lambda: 200., contains=lambda p: 50. <= p.x() <= 250. and 24. <= p.y() <= 224.)
        self.view.invert_y = False
        self.view.add_site_mode = True
        self.view.navigation_enabled = True
        self.view.freeze_mode = self.view.unfreeze_mode = False
        self.view.navigation_pan = point(0., 0.)
        self.view._pan_anchor = None
        self.view._pan_dragged = False
        self.view.location_clicked = Mock()
        self.view.location_double_clicked = Mock()
        self.view.setCursor = Mock()
        self.view.update = Mock()

    def event(self, x=200., y=74.):
        return Mock(button=lambda: GUI["Qt"].LeftButton, position=lambda: point(x, y))

    def test_single_click_uses_actual_map_rectangle(self):
        event = self.event()
        self.view.mousePressEvent(event)
        self.view.mouseReleaseEvent(event)
        self.view.location_clicked.emit.assert_called_once_with(1., 2.)
        self.assertFalse(self.view.location_double_clicked.emit.called)

    def test_drag_pans_without_adding(self):
        self.view.mousePressEvent(self.event(100., 100.))
        with patch.object(GUI["QApplication"], "startDragDistance", return_value=10):
            self.view.mouseMoveEvent(self.event(120., 110.))
        self.view.mouseReleaseEvent(self.event(120., 110.))
        self.assertFalse(self.view.location_clicked.emit.called)
        self.assertAlmostEqual(self.view.navigation_pan.x(), -.4)
        self.assertAlmostEqual(self.view.navigation_pan.y(), .4)

    def test_double_click_in_placement_mode_adds_once_and_never_moves(self):
        event = self.event()
        self.view.mousePressEvent(event)
        self.view.mouseReleaseEvent(event)
        self.view.mouseDoubleClickEvent(event)
        self.view.mouseReleaseEvent(event)
        self.view.location_clicked.emit.assert_called_once_with(1., 2.)
        self.assertFalse(self.view.location_double_clicked.emit.called)

    def test_normal_double_click_uses_same_map_transform(self):
        self.view.add_site_mode = False
        self.view.mouseDoubleClickEvent(self.event())
        self.view.location_double_clicked.emit.assert_called_once_with(1., 2.)
        self.assertFalse(self.view.location_clicked.emit.called)

    def test_outside_map_is_not_a_site(self):
        event = self.event(20., 74.)
        self.view.mousePressEvent(event)
        self.view.mouseReleaseEvent(event)
        self.assertFalse(self.view.location_clicked.emit.called)

    def test_inverse_transform_respects_panned_bounds_and_y_direction(self):
        self.view._coordinate_bounds = (8., 12., 6., 14.)
        self.view.invert_y = True
        self.assertEqual(self.view._position_to_coordinates(point(200., 74.)), (11., 8.))


if __name__ == "__main__":
    unittest.main()
