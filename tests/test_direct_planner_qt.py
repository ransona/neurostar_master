"""Real Qt integration checks, simulation only; skip if PySide6 unavailable."""
import importlib.util
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
    from PySide6.QtWidgets import QApplication,QPushButton,QCheckBox,QDialog
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

    def connect(self):
        c=self.window.controller
        c.connect(Calibration(dict(AP=105280,ML=75864,DV=41767),-8572,dict(AP=0,ML=261,DV=0,PISTON=0)),
                  dict(AP=0,ML=261,DV=0,PISTON=0),self.root/'api.json',allow_dv=True,allow_piston=True)
        self.window.refresh_live_position()

    def wait_idle(self):
        deadline=time.monotonic()+4
        while self.window.controller.has_active_motion() and time.monotonic()<deadline:
            self.app.processEvents();time.sleep(.01)
        self.assertFalse(self.window.controller.has_active_motion())

    def test_disconnected_keyboard_does_not_move(self):
        self.window.keyboard_nudge('AP',True,'AP')
        self.assertIsNone(self.window.controller.drive)
        self.dialogs.critical.assert_called()

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
            self.window._move_through_axis_positions_with_progress([(0,0,-.1),(1.1,0,-.1)],title='test',message='test')
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

    def test_manual_piston_uses_worker_and_verified_position(self):
        self.connect();self.window.set_syringe_position(3000)
        self.window.manual_injection_volume_nl=10
        self.window.manual_syringe_step(True)
        self.assertTrue(self.window.controller.has_active_motion())
        self.wait_idle();self.app.processEvents()
        self.assertAlmostEqual(self.window.current_syringe_position(),3010,delta=1/161.36)

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


if __name__=='__main__':unittest.main()
