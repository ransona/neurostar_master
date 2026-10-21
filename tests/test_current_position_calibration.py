import json
import math
import unittest
from unittest.mock import patch
from stereodrive_api import StereoDrive, Calibration
from stereodrive_api.protocol import SIGNS
from stereodrive_api.transport import Simulator

class CalibrationTests(unittest.TestCase):
    def test_measured_position_round_trip_with_backlash(self):
        raw=dict(AP=105802,ML=75864,DV=41819,PISTON=-2747)
        states=dict(AP=522,ML=261,DV=52,PISTON=5825)
        mm=dict(AP=42.75,ML=39.47,DV=32.01);nl=2983.25
        c=Calibration.from_current_position(raw,mm,nl,states)
        self.assertEqual(c.anchor_backlash,dict.fromkeys(raw,0))
        loaded=Calibration(**json.loads(json.dumps(c.__dict__)))
        scales=dict(AP=5225,ML=5225,DV=5225,PISTON=161.36);ref=loaded.reference(scales)
        for a in mm:self.assertAlmostEqual((raw[a]-states[a]-ref[a])/(SIGNS[a]*scales[a]),mm[a],delta=.5/5225)
        self.assertAlmostEqual((raw['PISTON']-states['PISTON']-ref['PISTON'])/161.36,nl,delta=.5/161.36)

    def test_invalid_measurements_rejected(self):
        raw=dict(AP=1,ML=2,DV=3,PISTON=4);states=dict.fromkeys(raw,0)
        for bad in (float('nan'),float('inf'),True):
            with self.assertRaises(ValueError):Calibration.from_current_position(raw,dict(AP=bad,ML=0,DV=0),3000,states)
        for bad in (-1,5001):
            with self.assertRaises(ValueError):Calibration.from_current_position(raw,dict(AP=0,ML=0,DV=0),bad,states)

    def test_counter_capture_sends_status_only_and_closes(self):
        port=Simulator(lambda *a,**k:None)
        with patch('stereodrive_api.api.Simulator',return_value=port):
            self.assertEqual(StereoDrive.read_idle_counts(),port.raw)
        self.assertTrue(port.closed)
        self.assertTrue(all(p[1]==0x0e for p in port.packets))
        self.assertEqual(len(port.packets),8)

    def test_capture_rejects_changing_or_moving_counts(self):
        for moving in (False,True):
            port=Simulator(lambda *a,**k:None);original=port.exchange;calls=[0]
            def exchange(data,size):
                calls[0]+=1
                if calls[0]==5:port.raw['AP']+=1
                reply=bytearray(original(data,size))
                if moving:reply[10]=1
                return bytes(reply)
            port.exchange=exchange
            with patch('stereodrive_api.api.Simulator',return_value=port):
                with self.assertRaises(RuntimeError):StereoDrive.read_idle_counts()
            self.assertTrue(port.closed)
            self.assertTrue(all(p[1]==0x0e for p in port.packets))

    def test_connected_capture_uses_existing_port_and_preserves_reference(self):
        drive=StereoDrive(simulate=True,require_calibration=False,persist_simulation=False)
        drive.connect(verified_backlash=dict(AP=0,ML=261,DV=0,PISTON=0))
        try:
            before=drive.position();port=drive._session.transport;start=len(port.packets)
            snapshot=drive.capture_calibration_counts()
            self.assertEqual(snapshot['raw_counts'],before['raw_counts'])
            self.assertEqual(snapshot['backlash_counts'],before['backlash_counts'])
            self.assertTrue(all(p[1]==0x0e for p in port.packets[start:]))
            self.assertEqual(drive.position(),before)
            drive._lock.acquire()
            try:
                with self.assertRaises(RuntimeError):drive.capture_calibration_counts()
            finally:drive._lock.release()
        finally:drive.close()

    def test_uncalibrated_jog_all_channels_bounds_and_history(self):
        from stereodrive_api.calibration_jog import CalibrationJog
        import tempfile
        from pathlib import Path
        with tempfile.TemporaryDirectory() as folder:
            jog=CalibrationJog(dict(AP=0,ML=261,DV=0,PISTON=0),journal_path=Path(folder)/'jog.json')
            try:
                before=jog.capture_calibration_counts()
                for a in ('AP','ML','DV','PISTON'):
                    step=20 if a=='PISTON' else .01
                    jog.move(a,step);jog.move(a,-step)
                after=jog.capture_calibration_counts()
                # Mechanical return may use a different raw backlash state.
                for a in before['raw_counts']:
                    self.assertEqual(after['raw_counts'][a]-after['backlash_counts'][a],before['raw_counts'][a]-before['backlash_counts'][a])
                jog.move('AP',2);jog.move('AP',1)
                with self.assertRaises(ValueError):jog.move('AP',.01)
                self.assertTrue(json.loads((Path(folder)/'jog.json').read_text())['valid'])
            finally:jog.close()

    def test_uncalibrated_direction_history_requires_all_initial_moves(self):
        from stereodrive_api.calibration_jog import CalibrationJog,INITIAL_DIRECTION_MOVE
        import tempfile
        from pathlib import Path
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'jog.json'
            jog=CalibrationJog(None,journal_path=path)
            try:
                self.assertFalse(json.loads(path.read_text())['valid'])
                jog.move('AP',INITIAL_DIRECTION_MOVE['AP'])
                self.assertFalse(json.loads(path.read_text())['valid'])
                for axis in ('ML','DV','PISTON'):
                    jog.move(axis,INITIAL_DIRECTION_MOVE[axis])
                self.assertTrue(json.loads(path.read_text())['valid'])
                snapshot=jog.capture_calibration_counts()
                self.assertTrue(all(snapshot['backlash_counts'][axis] is not None for axis in snapshot['backlash_counts']))
            finally:jog.close()

    def test_uncalibrated_jog_stop_cancels_and_invalidates(self):
        from stereodrive_api.calibration_jog import CalibrationJog
        import tempfile,time,threading
        from pathlib import Path
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'jog.json';jog=CalibrationJog(dict(AP=0,ML=261,DV=0,PISTON=0),journal_path=path)
            failures=[]
            def move():
                try:jog.move('AP',.1)
                except Exception as exc:failures.append(exc)
            thread=threading.Thread(target=move);thread.start()
            deadline=time.monotonic()+2
            while not any(p[1]==0x0c for p in jog.transport.packets) and time.monotonic()<deadline:time.sleep(.005)
            jog.stop();thread.join(2)
            self.assertTrue(failures);self.assertFalse(json.loads(path.read_text())['valid'])
            self.assertEqual(sum(p[1]==0x0c for p in jog.transport.packets),1)
            jog.close()
