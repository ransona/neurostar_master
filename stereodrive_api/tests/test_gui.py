"""Hidden-window simulation tests; never use the real serial transport."""
import tempfile
from pathlib import Path
import sys
import time
import tkinter as tk
import unittest
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from stereodrive_api import gui

class GuiTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.old_data = gui.DATA; gui.DATA = Path(self.directory.name)
        self.root = tk.Tk(); self.root.withdraw()
        self.app = gui.App(self.root)

    def wait(self, predicate):
        deadline = time.monotonic() + 4
        while time.monotonic() < deadline:
            self.root.update()
            if predicate(): return
            time.sleep(.01)
        self.fail('GUI state did not settle: ' + self.app.status.get())

    def tearDown(self):
        if self.app.worker:
            self.app.disconnect(); self.wait(lambda: self.app.worker is None)
        self.root.destroy(); gui.DATA = self.old_data; self.directory.cleanup()

    def test_buttons_reserve_single_move_and_restore_saved_position(self):
        self.app.connect(); self.wait(lambda: self.app.connected)
        self.app.steps['AP'].set('0.1')
        self.app.move('AP', 1); self.app.move('AP', 1)
        self.assertTrue(self.app.busy)
        self.assertTrue(all(str(b['state']) == 'disabled' for b in self.app.arrows))
        self.wait(lambda: not self.app.busy)
        self.assertEqual(self.app.mm['AP'].get(), '+0.100 mm')
        targets = [p for p in self.app.worker.session.transport.packets if p[1] == 0x0c]
        self.assertEqual(len(targets), 1)
        self.app.disconnect(); self.wait(lambda: self.app.worker is None)
        self.app.connect(); self.wait(lambda: self.app.connected)
        self.assertEqual(self.app.mm['AP'].get(), '+0.100 mm')
        self.assertEqual(self.app.worker.session.states['AP'], 0)

    def test_esc_stop_cancels_without_automatic_reverse(self):
        self.app.connect(); self.wait(lambda: self.app.connected)
        self.app.allow_dv.set(True)
        # Enablement is fixed at connection; reconnect to load it.
        self.app.disconnect();self.wait(lambda:self.app.worker is None)
        self.app.connect();self.wait(lambda:self.app.connected)
        self.app.move('DV', 1)
        self.wait(lambda: self.app.worker.session.busy)
        self.app.stop()
        self.wait(lambda: self.app.faulted)
        self.assertFalse(self.app.worker.session.store.read()['valid'])
        self.assertTrue(all(str(b['state']) == 'disabled' for b in self.app.arrows))

    def test_gui_piston_and_drill_use_public_api(self):
        self.app.allow_piston.set(True);self.app.allow_drill.set(True)
        self.app.connect();self.wait(lambda:self.app.connected)
        drive=self.app.worker.drive
        self.assertIsInstance(drive,gui.StereoDrive)
        self.app.inject('up');self.wait(lambda:not self.app.busy)
        self.assertIn('10.00 nL',self.app.piston_position.get())
        self.app.inject('down');self.wait(lambda:not self.app.busy)
        self.app.drill(True);self.wait(lambda:not self.app.busy)
        self.assertEqual(self.app.drill_status.get(),'Drill power: ON')
        self.app.drill(False);self.wait(lambda:not self.app.busy)
        self.assertEqual(self.app.drill_status.get(),'Drill power: OFF')
        self.assertIn(bytes.fromhex('af1101'),drive._session.transport.packets)

    def test_default_speed_and_no_unapproved_controls(self):
        self.assertEqual(self.app.speed.get(),'2')
        self.app.connect();self.wait(lambda:self.app.connected)
        self.assertEqual(str(self.app.drill_on_button['state']),'disabled')
        self.assertTrue(all(str(b['state'])=='disabled' for b in self.app.piston_buttons))
        self.assertEqual(self.app.worker.drive.speed_mm_s,2)

if __name__ == '__main__': unittest.main()
