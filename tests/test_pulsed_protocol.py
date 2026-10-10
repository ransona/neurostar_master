"""Pulse plan, timing and API checks: simulation only, never hardware."""
from pathlib import Path
import sys
import tempfile
import threading
import time
from types import SimpleNamespace
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'tools'))
import pulsed_protocol as pulse
from direct_api_controller import StereoDriveController,StereoDriveError
from stereodrive_api import Calibration

STATES=dict(AP=0,ML=261,DV=0,PISTON=0)
CAL=Calibration(dict(AP=105280,ML=75864,DV=41767),-8572,STATES)

def settings(**kwargs):
    values=dict(main_volume_nl=10,injection_depth_mm=.005,overshoot_mm=.005,
        insert_retract_speed_um_s=1000,insertion_rate_nl_min=60000,main_rate_nl_min=60000,post_inject_pause_s=0)
    return SimpleNamespace(**dict(values,**kwargs))


class PulseTests(unittest.TestCase):
    def test_insertion_and_main_volumes_are_separate_and_quantized(self):
        plan=pulse.injection_plan((0,0,0),SimpleNamespace(ap=0,ml=0,dv=0),settings(),-.02,.02)
        self.assertEqual(sum(e.volume_nl for e in plan if e.phase=='insertion_dose'),20)
        self.assertEqual(sum(e.volume_nl for e in plan if e.phase=='main_dose'),10)
        self.assertEqual([e.due_s for e in plan],sorted(e.due_s for e in plan))
        targets=[e.target for e in plan if e.target is not None]
        self.assertIn((0,0,.01),targets) # Do not skip overshoot endpoint.
        self.assertIn((0,0,.005),targets)
        self.assertEqual(targets[-1],(0,0,-.02))

    def test_zero_depth_has_no_insertion_dose(self):
        self.assertEqual(pulse.insertion_volume(settings(injection_depth_mm=0,overshoot_mm=0)),0)

    def test_scheduler_delays_overdue_pulses_instead_of_catching_up(self):
        now=[0];starts=[]
        class Controller:
            def syringe_step(self,*args,**kwargs):
                starts.append(now[0]);now[0]+=2 # Command slower than 1s requested interval.
                return {'PISTON':3000-10*len(starts)}
        events=[pulse.Pulse(i,'main_dose',volume_nl=10) for i in (1,2,3)]
        pulse.execute(Controller(),events,clock=lambda:now[0],sleep=lambda s:now.__setitem__(0,now[0]+s))
        self.assertEqual(starts,[1,3,5])

    def test_pause_waits_at_idle_and_shifts_timeline(self):
        now=[0];starts=[]
        class Controller:
            def syringe_step(self,*args,**kwargs):starts.append(now[0]);return {'PISTON':2990}
        pulse.execute(Controller(),[pulse.Pulse(1,'dose',volume_nl=10)],pause_requested=lambda:now[0]<2,
            clock=lambda:now[0],sleep=lambda s:now.__setitem__(0,now[0]+s))
        self.assertGreaterEqual(starts[0],3)
        with self.assertRaises(pulse.PulsePaused):
            pulse.execute(Controller(),[pulse.Pulse(0,'dose',volume_nl=10)],pause_requested=lambda:True,retract_on_pause=True)

    def test_cancel_prevents_every_remaining_dose(self):
        stopped=[False];starts=[]
        class Controller:
            def syringe_step(self,*args,**kwargs):starts.append(1);stopped[0]=True;return {'PISTON':2990}
        with self.assertRaises(pulse.PulseCancelled):
            pulse.execute(Controller(),[pulse.Pulse(0,'dose',volume_nl=10)]*3,stop_requested=lambda:stopped[0])
        self.assertEqual(len(starts),1)

    def test_drilling_gap_is_crossed_at_clearance_not_cutting_depth(self):
        surfaces=[(0,0,0),(.02,0,0),(.02,.02,0),(0,.02,0),(0,0,0)]
        plan=pulse.drilling_plan((0,0,-.02),surfaces,[0]*5,[.005]*5,[False,True,False,False,False],
            -.02,.02,(.01,.01,-.02),.1,.1)
        self.assertFalse(any(e.point_index==1 for e in plan))
        self.assertFalse(any(e.phase=='cut' and e.target[0]==.02 and e.target[1]==0 for e in plan))
        between=[e for e in plan if e.phase=='approach' and e.target[0]==.02]
        self.assertTrue(any(e.target[2]==-.02 for e in between))
        self.assertEqual([e.due_s for e in plan],sorted(e.due_s for e in plan))

    def test_spaced_borehole_sampling_is_uniform_and_respects_frozen_segments(self):
        perimeter=[(0,0,0),(2,0,0),(2,2,0),(0,2,0),(0,0,0)]
        holes=pulse.sample_boreholes(perimeter,1.1,[False,True,False,False,False])
        self.assertEqual(len(holes),8)
        self.assertAlmostEqual(holes[0][0],0)
        self.assertAlmostEqual(holes[0][1],0)
        self.assertTrue(any(hole[3] for hole in holes))
        distances=[((holes[(i+1)%len(holes)][0]-hole[0])**2+
                    (holes[(i+1)%len(holes)][1]-hole[1])**2)**.5 for i,hole in enumerate(holes)]
        self.assertLess(max(distances),1.1+1e-9)

    def test_per_hole_freeze_overrides_can_freeze_or_unfreeze_inherited_holes(self):
        holes=[(0,0,0,True),(1,0,0,False),(2,0,0,False)]
        effective=pulse.apply_borehole_freeze_overrides(holes,[False,True,None])
        self.assertEqual([hole[3] for hole in effective],[False,True,False])

    def test_spaced_boreholes_infer_surface_from_perimeter_profile(self):
        perimeter=[(0,0,0),(2,0,2),(2,2,2),(0,2,0),(0,0,0)]
        holes=pulse.sample_boreholes(perimeter,1.0,[False]*len(perimeter))
        self.assertAlmostEqual(holes[1][2],1.0)

    def test_borehole_plan_retracts_before_lateral_move_and_limits_depth(self):
        holes=[(0,0,0,False),(.5,0,0,False)]
        plan=pulse.borehole_drilling_plan((0,0,-.05),holes,[0,0],[.01,.01],-.02,.02,
            (.25,0,-.05),.1)
        drill=[event for event in plan if event.phase=='borehole_drill']
        self.assertTrue(drill)
        self.assertEqual({event.point_index for event in drill}, {0, 1})
        self.assertTrue(all(event.target[2] <= .010001 for event in drill))
        self.assertTrue(all(event.target[2] >= -1e-9 for event in drill))
        # Circuit duration is no longer used to insert waits between holes.
        self.assertLess(max(event.due_s for event in plan), 1.0)
        for left,right in zip(plan,plan[1:]):
            if left.target and right.target and abs(left.target[0]-right.target[0])>1e-9:
                self.assertLessEqual(left.target[2],-.02+1e-9)
                self.assertLessEqual(right.target[2],-.02+1e-9)

    def test_piston_preflight_rejects_whole_overbudget_without_targets(self):
        with tempfile.TemporaryDirectory() as directory:
            c=StereoDriveController();c.connect(CAL,STATES,Path(directory)/'api.json',allow_piston=True)
            try:
                with self.assertRaises(ValueError):c.validate_piston_steps([-10]*301)
                self.assertFalse(any(p[1]==0x0c for p in c.drive._session.transport.packets))
                c.validate_piston_steps([-10]*300)
            finally:c.close()

    def test_direct_microstep_adapter_honors_steps_and_dwell(self):
        with tempfile.TemporaryDirectory() as directory:
            c=StereoDriveController();c.connect(CAL,STATES,Path(directory)/'api.json',allow_dv=True,allow_pulsed=True)
            try:
                c.prepare_motion();started=time.monotonic()
                c.move_axis_to_target('DV',.01,step_mm=.005,dwell_seconds=.3)
                self.assertGreaterEqual(time.monotonic()-started,.6)
                self.assertEqual(sum(p[1]==0x0c for p in c.drive._session.transport.packets),2)
                self.assertAlmostEqual(c.get_current_axis('DV'),.01,delta=1/5225)
            finally:c.close()


if __name__=='__main__':unittest.main()
