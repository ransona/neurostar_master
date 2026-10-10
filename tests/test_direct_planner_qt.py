"""Real Qt integration checks, simulation only; skip if PySide6 unavailable."""
import importlib.util
import json
from pathlib import Path
import sys
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

HAS_QT=importlib.util.find_spec('PySide6') is not None
if HAS_QT:
    sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'tools'))
    from PySide6.QtCore import QTimer
    from PySide6.QtWidgets import QApplication,QPushButton,QCheckBox,QDialog,QComboBox
    import craniotomy_qt as planner
    from direct_api_controller import StereoDriveController
    from stereodrive_api import Calibration


@unittest.skipUnless(HAS_QT,'Install PySide6 to run actual planner integration tests')
class PlannerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):cls.app=QApplication.instance() or QApplication([])

    def setUp(self):
        self.directory=tempfile.TemporaryDirectory()
        self.root=Path(self.directory.name)
        self.settings_patch=patch.object(planner.CraniotomyWindow,'_config_root_dir',lambda _:self.root)
        self.settings_patch.start()
        self.dialogs_patch=patch.object(planner,'QMessageBox')
        self.dialogs=self.dialogs_patch.start()
        self.window=planner.CraniotomyWindow(StereoDriveController())
        self.window.show()

    def tearDown(self):
        self.window.controller.close()
        self.window.close();self.app.processEvents()
        self.dialogs_patch.stop();self.settings_patch.stop();self.directory.cleanup()

    def connect(self,**kwargs):
        c=self.window.controller
        options=dict(allow_dv=True,allow_piston=True,travel_limits=dict(AP=(-40,40),ML=(-40,40),DV=(-40,40),PISTON=(0,5000)));options.update(kwargs)
        c.connect(Calibration(dict(AP=105280,ML=75864,DV=41767),-8572,dict(AP=0,ML=261,DV=0,PISTON=0)),
                  dict(AP=0,ML=261,DV=0,PISTON=0),self.root/'api.json',**options)
        self.window.refresh_live_position()

    def wait_idle(self):
        deadline=time.monotonic()+4
        while self.window.controller.has_active_motion() and time.monotonic()<deadline:
            self.app.processEvents();time.sleep(.01)
        self.assertFalse(self.window.controller.has_active_motion())

    def test_disconnected_keyboard_does_not_move(self):
        with patch.object(self.window,'_focus_is_editable',return_value=False):
            self.window.keyboard_nudge('AP',True,'AP')
        self.assertIsNone(self.window.controller.drive)
        self.dialogs.critical.assert_called()

    def test_motion_options_defaults_apply_persist_and_busy_guard(self):
        for axis, edits in self.window.direct_limit_edits.items():
            self.assertEqual([e.value() for e in edits],[0,5000 if axis=='PISTON' else 40])
        self.window.direct_limit_edits['AP'][1].setValue(25)
        self.window.direct_speed_combo.setCurrentText('2')
        self.window._apply_direct_motion_options()
        self.assertEqual(json.loads((self.root/'direct-control.json').read_text())['travel_limits']['AP'],[0,25])
        self.window.close();self.app.processEvents()
        self.window=planner.CraniotomyWindow(StereoDriveController())
        self.assertEqual(self.window.direct_limit_edits['AP'][1].value(),25)
        self.assertEqual(self.window.direct_speed_combo.currentText(),'2')
        self.connect()
        self.window._apply_direct_motion_options()
        self.assertEqual(self.window.controller.drive.travel_limits['AP'],(0,25))
        self.assertEqual(self.window.controller.drive.speed_mm_s,2)
        self.window.direct_speed_combo.setCurrentText('1')
        with patch.object(self.window,'_motion_is_active',return_value=True):
            self.window._apply_direct_motion_options()
        self.assertEqual(self.window.controller.drive.speed_mm_s,2)
        self.window.direct_limit_edits['AP'][0].setValue(30)
        self.window._apply_direct_motion_options()
        self.dialogs.warning.assert_called()
        self.assertEqual(self.window.controller.drive.travel_limits['AP'],(0,25))

    def test_keyboard_absolute_progress_and_internal_bregma(self):
        self.connect()
        self.window.set_local_bregma()
        self.window.keyboard_nudge('AP',True,'AP')
        self.wait_idle()
        self.window.controller.prepare_motion()
        self.assertTrue(self.window._move_through_axis_positions_with_progress([(.02,.02,-.01)],title='test',message='simulated move'))
        self.assertAlmostEqual(self.window.get_bregma_position()[0],.02,delta=1/5225)

    def test_invalid_later_path_leg_is_rejected_before_any_target(self):
        self.connect()
        drive=self.window.controller.drive
        before=len(drive._session.transport.packets)
        with self.assertRaises(ValueError):
            self.window._move_through_axis_positions_with_progress([(0,0,-.1),(40.1,0,-.1)],title='test',message='test')
        self.assertFalse(any(p[1]==0x0c for p in drive._session.transport.packets[before:]))

    def test_protocols_blocked_before_any_hardware_target(self):
        self.connect()
        self.window.start_single_injection();self.window.start_drilling_round();self.window.resume_injection_from_selected()
        self.assertEqual(self.dialogs.warning.call_count,3)
        self.assertFalse(any(p[1]==0x0c for p in self.window.controller.drive._session.transport.packets))

    def test_disconnect_clears_stale_cross_and_readings(self):
        self.connect()
        self.window.refresh_live_position()
        self.window.controller.close()
        self.window.refresh_live_position()
        self.assertEqual(self.window.current_ap_label.text(),'—')
        self.assertIsNone(self.window.top_view.current_point)
        self.assertIsNone(self.window.injection_sites_view.current_point)

    def test_direct_settings_separate_and_survive_startup(self):
        self.window.direct_control_settings=dict(calibration=dict(axis_zero_counts=dict(AP=105280,ML=75864,DV=41767),
            piston_3000_count=-8572,anchor_backlash=dict(AP=0,ML=261,DV=0,PISTON=0)),speed_mm_s=2,
            allow_dv=True,allow_piston=True,allow_drill=False,verified=True,new_reference=True)
        self.window.home_axis=(.1,.2,.3)
        self.window._save_general_settings()
        saved=json.loads((self.root/'direct-control.json').read_text())
        general=json.loads(self.window._general_settings_path().read_text())
        self.assertEqual(saved['speed_mm_s'],2)
        self.assertNotIn('verified',saved);self.assertNotIn('new_reference',saved)
        self.assertNotIn('home_axis',general);self.assertNotIn('axis_zero_reference',general)
        self.window.close();self.app.processEvents()
        self.window=planner.CraniotomyWindow(StereoDriveController())
        self.assertEqual(self.window.direct_control_settings['speed_mm_s'],2)
        self.assertEqual(self.window.home_axis,(.1,.2,.3))
        self.assertIsNone(self.window.controller.drive)

    def test_legacy_direct_fields_migrate_without_loss(self):
        path=self.window._direct_control_settings_path()
        self.window.close();self.app.processEvents()
        path.unlink() # Test-owned temporary file only.
        self.window._general_settings_path().write_text(json.dumps(dict(home_axis=[.1,.2,.3],work_axis=[0,0,0],
            axis_zero_reference=dict(AP=105280,ML=75603,DV=41767))))
        self.window=planner.CraniotomyWindow(StereoDriveController())
        saved=json.loads(path.read_text())
        self.assertEqual(saved['home_axis'],[.1,.2,.3])
        self.assertEqual(saved['axis_zero_reference']['ML'],75603)

    def test_wrong_mode_direct_settings_fail_closed_and_are_preserved(self):
        path=self.window._direct_control_settings_path()
        payload=json.loads(path.read_text());payload['live']=True
        original=json.dumps(payload)
        path.write_text(original)
        self.window._load_direct_control_settings()
        self.window._save_general_settings()
        self.assertEqual(self.window.direct_control_settings,{})
        self.assertIsNone(self.window.home_axis)
        self.assertEqual(path.read_text(),original)

    def test_manual_piston_uses_worker_and_verified_position(self):
        self.connect();self.window.set_syringe_position(3000)
        self.window.manual_injection_volume_nl=10
        self.window.manual_syringe_step(True)
        self.assertTrue(self.window.controller.has_active_motion())
        self.wait_idle();self.app.processEvents()
        self.assertAlmostEqual(self.window.current_syringe_position(),3010,delta=1/161.36)

    def pulse_settings(self):
        return planner.InjectionProtocolSettings(main_volume_nl=10,insertion_rate_nl_min=60000,
            main_rate_nl_min=60000,injection_depth_mm=.005,insert_retract_speed_um_s=1000,
            overshoot_mm=0,post_inject_pause_s=0)

    def test_pulsed_injection_worker_uses_gui_bregma_and_verified_counter(self):
        self.connect(allow_pulsed=True)
        self.window.controller.prepare_motion();self.window.controller.goto_axis_position(.02,.01,.01)
        self.window.set_local_bregma()
        self.window.validation_clearance_mm=.02
        self.dialogs.warning.return_value=planner.QMessageBox.Yes
        self.window._start_injection_sequence([planner.InjectionSite(0,0,0)],self.pulse_settings(),[10],False,10,0,1,'test')
        deadline=time.monotonic()+8
        while self.window.injection_thread.is_alive() and time.monotonic()<deadline:
            self.app.processEvents();time.sleep(.01)
        self.assertFalse(self.window.injection_thread.is_alive());self.app.processEvents()
        p=self.window.controller.get_current_axis_position()
        self.assertAlmostEqual(p[0],.02,delta=1/5225)
        self.assertAlmostEqual(p[1],.01,delta=1/5225)
        self.assertAlmostEqual(p[2],-.01,delta=1/5225)
        self.assertAlmostEqual(self.window.current_syringe_position(),2980,delta=1/161.36)
        self.assertIsNone(self.window.controller.error)

    def test_pulsed_workflow_preflight_rejects_later_site_and_whole_dose(self):
        self.connect(allow_pulsed=True);self.window.injection_clearance_axis_dv=-.02
        self.window.pulsed_clearance_mm=.02
        with self.assertRaises(ValueError):
            self.window._preflight_pulsed_injections([planner.InjectionSite(0,0,0),planner.InjectionSite(40.1,0,0)],self.pulse_settings(),False,10)
        with self.assertRaises(Exception):
            self.window._preflight_pulsed_injections([planner.InjectionSite(0,0,0)]*251,self.pulse_settings(),False,10)

        self.window._preflight_pulsed_injections([planner.InjectionSite(0,0,0)]*6,self.pulse_settings(),False,10)
        self.assertFalse(any(p[1]==0x0c for p in self.window.controller.drive._session.transport.packets))

    def test_resume_selected_routes_remaining_sites_to_pulsed_sequence(self):
        self.connect(allow_pulsed=True);self.window.set_local_bregma()
        self.window.injection_sites=[planner.InjectionSite(0,0,0),planner.InjectionSite(.01,.02,0)]
        self.window.refresh_injection_sites_list();self.window.injection_sites_list.setCurrentRow(1)
        with patch.object(self.window,'_injection_protocol_settings',return_value=self.pulse_settings()), \
                patch.object(self.window,'_start_injection_sequence') as start:
            self.window.resume_injection_from_selected()
        self.assertEqual(start.call_args.kwargs['start_site_offset'],1)
        self.assertEqual(len(start.call_args.kwargs['sites']),1)
        self.assertEqual(start.call_args.kwargs['sites'][0].ml,.02)

    def test_pulsed_drill_pause_retracts_without_losing_reference(self):
        self.connect(allow_pulsed=True,allow_drill=True)
        self.window.controller.drive.drill_on()
        self.window.pulsed_clearance_mm=.02;self.window.pulsed_drill_rate=1
        self.window.drill_clearance_axis_dv=-.02
        self.window.drill_pause_requested.set()
        surfaces=[(0,0,0),(.01,0,0),(0,.01,0),(0,0,0)]
        self.window._run_pulsed_drilling(surfaces,[0]*4,[.005]*4,[False]*4,.1,(0,0,-.02))
        self.assertIsNone(self.window.controller.error)
        self.assertTrue(self.window.controller.drive._session.store.read()['valid'])
        self.assertFalse(self.window.controller.drive.drill_state())
        self.assertAlmostEqual(self.window.controller.get_current_axis('DV'),-.02,delta=1/5225)
        self.assertTrue(self.window.drilling_paused)

    def test_pulsed_drilling_start_completes_closed_perimeter(self):
        self.connect(allow_pulsed=True,allow_drill=True)
        self.window.set_local_bregma();self.window.controller.drive.drill_on()
        self.window.validation_clearance_mm=.02
        self.window.trajectory=[(0,0,0),(.01,0,0),(0,.01,0),(0,0,0)]
        self.window.seeds=[planner.SeedPoint(0,0,0,0,0),planner.SeedPoint(1,180,.01,0,0)]
        self.window.drilled_depths=[0]*4;self.window.frozen_points=[False]*4
        self.window.current_target_depth_mm=.005;self.window.drill_depth.setValue(.005)
        self.window.drill_rate_mm_per_s.setValue(1);self.window.round_time_seconds.setValue(1)
        self.window.cut_offset.setValue(0);self.window.mid_ap.setValue(0);self.window.mid_ml.setValue(0)
        self.window.auto_start_rounds.setChecked(False)
        self.dialogs.question.return_value=planner.QMessageBox.Yes
        self.dialogs.warning.return_value=planner.QMessageBox.Yes
        outcomes=[];self.window.drill_round_finished_signal.connect(outcomes.append)
        self.window.start_drilling_round()
        self.assertIsNotNone(self.window.drill_thread)
        deadline=time.monotonic()+10
        while self.window.drill_thread.is_alive() and time.monotonic()<deadline:
            self.app.processEvents();time.sleep(.01)
        self.assertFalse(self.window.drill_thread.is_alive());self.app.processEvents()
        self.assertEqual(outcomes,['completed'])
        self.assertIsNone(self.window.controller.error)
        self.assertEqual(self.window.drilled_depths,[.005]*4)
        self.assertAlmostEqual(self.window.controller.get_current_axis('DV'),-.02,delta=1/5225)

    def test_zero_setup_wizard_simulation(self):
        # Persistent mechanical targets survive reconnect with the same Axis zero.
        self.window.axis_zero_reference=dict(AP=105280,ML=75864-261,DV=41767)
        self.window.home_axis=(.1,0,0)
        invoked=[False]
        timer=QTimer()
        deadline=time.monotonic()+8
        def automate():
            dialog=next((d for d in self.window.findChildren(QDialog) if d.isVisible() and 'measured zero' in d.windowTitle()),None)
            if not dialog:return
            if time.monotonic()>deadline:
                dialog.reject();timer.stop();return
            if not invoked[0]:
                buttons=dialog.findChildren(QPushButton)
                next(b for b in buttons if b.text().startswith('Use SIMULATION')).click()
                next(c for c in dialog.findChildren(QComboBox) if c.count()==2 and c.itemText(0)=='1').setCurrentText('2')
                next(c for c in dialog.findChildren(QCheckBox) if c.text().startswith('Enable DV')).setChecked(True)
                next(c for c in dialog.findChildren(QCheckBox) if c.text().startswith('Measured zero/anchor')).setChecked(True)
                next(b for b in buttons if b.text().startswith('Connect / restore')).click()
                invoked[0]=True
            if self.window.controller.drive is not None:
                button=next(b for b in dialog.findChildren(QPushButton) if b.text()=='Close setup')
                if button.isEnabled():button.click();timer.stop()
        timer.timeout.connect(automate);timer.start(100)
        self.window.open_direct_usb_setup()
        self.assertTrue(invoked[0]);self.assertIsNotNone(self.window.controller.drive)
        self.assertTrue(self.window.controller.drive.position()['calibrated'])
        self.assertEqual(self.window.home_axis,(.1,0,0))
        saved=json.loads(self.window._direct_control_settings_path().read_text())
        self.assertIn('calibration',saved)
        self.assertEqual(saved['speed_mm_s'],2)
        self.assertTrue(saved['allow_dv'])
        self.assertNotIn('verified_backlash',saved)


if __name__=='__main__':unittest.main()
