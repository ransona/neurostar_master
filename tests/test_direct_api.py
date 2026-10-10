"""Direct backend safety/integration tests, using API simulation only."""
import tempfile
from pathlib import Path
import sys
import threading
import time
import unittest
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/"tools"))
from direct_api_controller import StereoDriveController, StereoDriveError
from stereodrive_api import StereoDrive, Calibration

STATES=dict(AP=0,ML=261,DV=0,PISTON=0)
CAL=Calibration(dict(AP=105280,ML=75864,DV=41767),-8572,STATES)
LIMITS=dict(AP=(-40,40),ML=(-40,40),DV=(-40,40),PISTON=(0,5000))


class DirectTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.path=Path(self.tmp.name)/"state.json"

    def drive(self,**kwargs):
        kwargs.setdefault('travel_limits',LIMITS)
        d=StereoDrive(state_path=self.path,calibration=CAL,**kwargs)
        d.connect(verified_backlash=STATES);self.addCleanup(d.close)
        return d

    def test_zero_required_for_all_motion_and_drill_on(self):
        d=StereoDrive(state_path=self.path,allow_dv=True,allow_piston=True,allow_drill=True)
        d.connect(verified_backlash=STATES);self.addCleanup(d.close)
        for action in (lambda:d.move_mm('AP',.01),lambda:d.piston_step('up',10),
                       lambda:d.move_axis_to('AP',.01),lambda:d.drill_on()):
            with self.assertRaises(ValueError):action()
        self.assertFalse(any(p[1] in (0x0c,0x11) for p in d._session.transport.packets))
        d.position();d.stop() # Read/Stop must remain available without calibration.
        with self.assertRaises(ValueError):StereoDrive(simulate=False,require_calibration=False)

    def test_absolute_targets_and_preflight_all_axes(self):
        d=self.drive(allow_dv=True)
        d.move_axes_to(dict(AP=.1,ML=.2,DV=-.1))
        p=d.position()['axes_mm']
        for a,t in dict(AP=.1,ML=.2,DV=-.1).items():self.assertAlmostEqual(p[a],t,delta=1/5225)
        packets=len(d._session.transport.packets)
        with self.assertRaises(ValueError):d.move_axes_to(dict(AP=.2,ML=40.1))
        self.assertFalse(any(p[1]==0x0c for p in d._session.transport.packets[packets:]))
        packets=len(d._session.transport.packets)
        with self.assertRaises(ValueError):d.validate_axis_path([dict(DV=-.2),dict(AP=40.1)])
        self.assertFalse(any(p[1]==0x0c for p in d._session.transport.packets[packets:]))

    def test_absolute_dv_opt_in_nonfinite_and_combined_distance(self):
        d=self.drive()
        for target in (dict(DV=.01),dict(AP=float('nan')),dict(AP=True),dict(AP=40.1,ML=.8),dict(PISTON=10)):
            with self.assertRaises(ValueError):d.move_axes_to(target)
        self.assertFalse(any(p[1]==0x0c for p in d._session.transport.packets))

    def test_absolute_moves_preserve_fractional_targets(self):
        d=self.drive()
        for target in (.001,.002,.003,.004,.005):d.move_axis_to('AP',target)
        self.assertAlmostEqual(d.position()['axes_mm']['AP'],.005,delta=1/5225)

    def test_rounded_held_axis_at_envelope_edge_is_not_a_new_target(self):
        d=self.drive(travel_limits=dict(LIMITS,AP=(0,1.0001)))
        d.move_axis_to('AP',.0001);d.close();d.connect() # Simulator-only restoration test.
        d.move_axis_to('AP',1.0001)
        p=d.position()['axes_mm']
        self.assertGreater(p['AP'],1.0001) # Quantized motor readout exceeds fractional target.
        d.move_axes_to(dict(p,ML=.01))
        self.assertAlmostEqual(d.position()['axes_mm']['ML'],.01,delta=1/5225)

    def test_adapter_menu_keyboard_injector_and_stop_routes_api(self):
        c=StereoDriveController();self.addCleanup(c.close)
        with self.assertRaises(StereoDriveError):c.prepare_motion()
        c.connect(CAL,STATES,self.path,allow_dv=True,allow_piston=True,travel_limits=LIMITS)
        c.prepare_motion();c.goto_axis_position(.02,.01,-.01)
        c.wait_for_axis_position(.02,.01,-.01)
        c.prepare_motion();c.set_nudge_step('AP',.01);c.nudge_axis('AP',True)
        c.wait_until_stopped()
        self.assertAlmostEqual(c.get_current_axis('AP'),.03,delta=1/5225)
        c.prepare_motion();c.syringe_step('10 nl',up=True)
        self.assertAlmostEqual(c.read_injectomate_calibrate_scale_nl(),3010,delta=1/161.36)
        with self.assertRaises(StereoDriveError):c.benchmark_axis_moves(distances_mm=[2])
        c.stop()
        self.assertFalse(c.drive.drill_state())

    def test_idle_display_read_cannot_interrupt_protocol_command_start(self):
        c=StereoDriveController();self.addCleanup(c.close)
        c.connect(CAL,STATES,self.path);c.prepare_motion()
        entered=threading.Event();release=threading.Event();errors=[]
        original=c.drive.position
        def delayed_position():
            if threading.current_thread().name=='display-reader':
                with c.drive._lock:
                    entered.set();release.wait(2)
            return original()
        def move():
            try:c.goto_axis_position(.01,0,0)
            except Exception as exc:errors.append(exc)
        with patch.object(c.drive,'position',side_effect=delayed_position):
            reader=threading.Thread(target=c.get_current_axis_position,name='display-reader');reader.start()
            self.assertTrue(entered.wait(1))
            mover=threading.Thread(target=move);mover.start()
            time.sleep(.02)
            self.assertTrue(mover.is_alive())
            release.set();reader.join(2);mover.join(3)
        self.assertFalse(reader.is_alive());self.assertFalse(mover.is_alive())
        self.assertEqual(errors,[])
        self.assertAlmostEqual(c.get_current_axis('AP'),.01,delta=1/5225)

    def test_absolute_stop_cancels_remaining_legs(self):
        d=self.drive(allow_dv=True);errors=[]
        def move():
            try:d.move_axes_to(dict(AP=.2,ML=.2,DV=-.2))
            except Exception as exc:errors.append(exc)
        thread=threading.Thread(target=move);thread.start()
        deadline=time.monotonic()+2
        while not d._session.busy and time.monotonic()<deadline:time.sleep(.005)
        self.assertTrue(d._session.busy)
        self.assertIsInstance(d.live_position()['raw_counts'],dict)
        d.stop();thread.join(3)
        self.assertFalse(thread.is_alive());self.assertTrue(errors)
        self.assertEqual(sum(p[1]==0x0c for p in d._session.transport.packets),1)
        self.assertFalse(d._session.store.read()['valid'])

    def test_stop_always_requests_drill_off_without_on_optin(self):
        d=self.drive()
        d._session.transport.drill_on=True
        d.stop()
        self.assertFalse(d.drill_state())
        self.assertIn(bytes.fromhex('af1100'),d._session.transport.packets)

    def test_verified_history_must_match_restore_and_rejects_boolean(self):
        d=self.drive()
        d.move_mm('AP',-.01);d.close()
        other=StereoDrive(state_path=self.path,calibration=CAL,travel_limits=LIMITS)
        self.addCleanup(other.close)
        with self.assertRaisesRegex(ValueError,'differs'):
            other.connect(verified_backlash=STATES)
        with self.assertRaisesRegex(ValueError,'Invalid verified'):
            other.connect(verified_backlash=dict(STATES,AP=False),new_reference=True)
        other.connect() # Restore saved history without overriding it.
        self.assertEqual(other.position()['backlash_counts']['AP'],522)
        with self.assertRaises(ValueError):StereoDrive(speed_mm_s=True)

    def test_piston_capacity_guard_before_target(self):
        cal=Calibration(CAL.axis_zero_counts,-8572-round(1995*161.36),STATES)
        d=StereoDrive(state_path=self.path,calibration=cal,allow_piston=True)
        d.connect(verified_backlash=STATES);self.addCleanup(d.close)
        with self.assertRaisesRegex(ValueError,'travel range'):
            d.piston_step('up',10)
        self.assertFalse(any(p[1]==0x0c for p in d._session.transport.packets))

    def test_default_ranges_larger_moves_and_custom_limits(self):
        d=self.drive(travel_limits=None,allow_piston=True)
        self.assertEqual(d.travel_limits['AP'],(0,40))
        d.move_axes_to(dict(AP=2,ML=3))
        d.move_mm('AP',2)
        for direction in ('up','up','down','down'): d.piston_step(direction,100)
        d.validate_piston_steps([-10]*200) # No connection-relative volume cap.
        before=len(d._session.transport.packets)
        for action in (lambda:d.move_axis_to('AP',-.01),lambda:d.move_axis_to('ML',40.01)):
            with self.assertRaises(ValueError):action()
        self.assertFalse(any(p[1]==0x0c for p in d._session.transport.packets[before:]))
        d.configure_motion(speed_mm_s=1,travel_limits=dict(d.travel_limits,AP=(3,5),PISTON=(2995,3005)))
        self.assertEqual(d._session.speed_mm_s,1)
        with self.assertRaises(ValueError):d.move_mm('AP',2)
        with self.assertRaises(ValueError):d.piston_step('up',10)
        with self.assertRaises(ValueError):d.validate_piston_steps([-10])
        d.move_axis_to('AP',5)

    def test_invalid_limits_rejected_and_busy_settings_unchanged(self):
        from stereodrive_api.limits import DEFAULT_LIMITS
        for pair in ((2,1),(0,float('nan')),(False,40)):
            with self.assertRaises(ValueError):StereoDrive(travel_limits=dict(DEFAULT_LIMITS,AP=pair))
        with self.assertRaises(ValueError):StereoDrive(travel_limits=dict(DEFAULT_LIMITS,PISTON=(0,5001)))
        d=self.drive()
        with d._lock:
            with self.assertRaises(RuntimeError):d.configure_motion(speed_mm_s=1,travel_limits=DEFAULT_LIMITS)
        self.assertEqual(d.speed_mm_s,2)

    def test_larger_syringe_volume_verified_serial_steps_and_full_preflight(self):
        c=StereoDriveController();self.addCleanup(c.close)
        c.connect(CAL,STATES,self.path,allow_piston=True)
        updates=[];c.prepare_motion()
        result=c.syringe_step('150 nl',on_completed=updates.append)
        self.assertAlmostEqual(result['PISTON'],3150,delta=1/161.36)
        self.assertEqual(len(updates),2)
        self.assertEqual(sum(p[1]==0x0c for p in c.drive._session.transport.packets),2)
        before=len(c.drive._session.transport.packets)
        with self.assertRaises(ValueError):c.syringe_step('2000 nl')
        self.assertFalse(any(p[1]==0x0c for p in c.drive._session.transport.packets[before:]))
        self.assertIsNone(c.error)

    def test_empty_fill_respect_custom_limits_and_do_not_invent_partial_steps(self):
        c=StereoDriveController();self.addCleanup(c.close)
        c.connect(CAL,STATES,self.path,allow_piston=True,travel_limits=dict(LIMITS,PISTON=(2985,3035)))
        c.prepare_motion();result=c.empty_syringe()
        self.assertAlmostEqual(result['PISTON'],2990,delta=1/161.36)
        c.prepare_motion();result=c.fill_syringe()
        self.assertAlmostEqual(result['PISTON'],3030,delta=1/161.36)
        self.assertEqual(c.get_current_axis_position(),(0,0,0))
        self.assertFalse(any(p[1]==0x0c and p[2]!=0x70 for p in c.drive._session.transport.packets))

    def test_piston_plan_uses_commanded_fractional_counts_not_rounded_readout(self):
        d=self.drive(allow_piston=True)
        d.piston_step('up',20)
        steps=d.plan_piston_to(0)
        self.assertEqual(sum(steps),-3020)
        d.validate_piston_steps(steps)
        self.assertEqual(sum(d.plan_piston_to(5000)),1980)
        with self.assertRaises(ValueError):d.plan_piston_to(5001)

    def test_benchmark_preflights_all_axes_and_stops_without_reversal(self):
        c=StereoDriveController();self.addCleanup(c.close)
        c.connect(CAL,STATES,self.path)
        with self.assertRaises(ValueError):c.benchmark_axis_moves(['AP','DV'],[.01],1)
        self.assertFalse(any(p[1]==0x0c for p in c.drive._session.transport.packets))
        c.prepare_motion();rows=c.benchmark_axis_moves(['AP'],[.01],1)
        self.assertEqual([row['direction'] for row in rows],['+','-'])
        self.assertAlmostEqual(c.get_current_axis('AP'),0,delta=1/5225)
        stopped=threading.Event();c.prepare_motion()
        before=len(c.drive._session.transport.packets)
        with self.assertRaises(StereoDriveError):
            c.benchmark_axis_moves(['AP'],[.01],1,stop_requested=stopped.is_set,
                                   progress_callback=lambda *args:stopped.set())
        self.assertEqual(sum(p[1]==0x0c for p in c.drive._session.transport.packets[before:]),1)
        self.assertAlmostEqual(c.get_current_axis('AP'),.01,delta=1/5225)
        self.assertIsNone(c.error)
        self.assertTrue(c.drive._session.store.read()['valid'])


if __name__=='__main__':unittest.main()
