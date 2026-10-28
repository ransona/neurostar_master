import json
from pathlib import Path
import struct
import tempfile
import threading
import time
import unittest
from stereodrive_api import StereoDrive, Calibration
from stereodrive_api.protocol import target, BACKLASH

INITIAL = dict(AP=0, ML=261, DV=0, PISTON=0)

class APITests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = Path(self.tmp.name)/'state.json'

    def drive(self, **kwargs):
        kwargs.setdefault('require_calibration',False) # Relative-reference tests are simulator-only.
        d = StereoDrive(state_path=self.path, **kwargs)
        d.connect(verified_backlash=INITIAL)
        self.addCleanup(d.close)
        return d

    def test_live_configuration_does_not_validate_simulator_fixture_against_hardware_limits(self):
        d = StereoDrive(simulate=False, calibration=Calibration(
            dict(AP=105280, ML=75864, DV=41767), -8572, INITIAL),
            travel_limits=dict(AP=(5,70), ML=(5,65), DV=(5,65), PISTON=(503,4500)))
        self.assertEqual(d.travel_limits['AP'], (5.0, 70.0))
        self.assertEqual(d.simulation_start_position['AP'], 0.0)
        self.assertFalse(d.simulate)

    def test_packets_match_captured_profiles(self):
        self.assertEqual(target('AP', 105228, 1).hex(), 'af0c400c9b01001a05401f401f00010104')
        self.assertEqual(target('PISTON', -1133, 1).hex(), 'af0c7093fbffff1c07401f401f00010004')
        self.assertEqual(target('DV', 41245, 1)[15], 1)
        self.assertEqual(target('DV', 41819, -1)[15], 0)

    def test_new_reference_requires_all_verified_states(self):
        d=StereoDrive(state_path=self.path)
        with self.assertRaises(ValueError): d.connect()
        with self.assertRaises(ValueError): d.connect(verified_backlash={'AP':0})

    def test_all_axes_return_and_other_channels_stay_fixed(self):
        d=self.drive(allow_dv=True)
        for axis in ('AP','ML','DV'):
            raw=d.position()['raw_counts']
            d.move_mm(axis,.1)
            self.assertAlmostEqual(d.position()['axes_mm'][axis],.1,delta=1/5225)
            for other in raw:
                if other!=axis: self.assertEqual(d.position()['raw_counts'][other],raw[other])
            d.move_mm(axis,-.1)
            self.assertAlmostEqual(d.position()['axes_mm'][axis],0,delta=1/5225)

    def test_event_log_records_requested_and_completed_motion_with_utc(self):
        d=self.drive();d.move_mm('AP',.01)
        records=[json.loads(line) for line in self.path.with_suffix('.jsonl').read_text().splitlines()]
        requested=next(record for record in records if record['event']=='MOVE_REQUESTED')
        arrived=next(record for record in records if record['event']=='ARRIVED')
        self.assertEqual(requested['axis'],'AP')
        self.assertEqual(requested['target_raw'],arrived['raw'])
        self.assertTrue(requested['utc'].endswith('+00:00'))
        self.assertTrue(any(record['event']=='OUT' for record in records))
        self.assertTrue(any(record['event']=='IN' for record in records))

    def test_piston_reversal_and_three_same_direction_moves(self):
        d=self.drive(allow_piston=True)
        # 20 nL triples remain inside the cumulative 100 nL envelope.
        raw=d.position()['raw_counts']['PISTON']; initial_raw=raw; deltas=[]
        for direction in ('up','up','up','down','down','down'):
            d.piston_step(direction,20)
            now=d.position()['raw_counts']['PISTON'];deltas.append(now-raw);raw=now
        self.assertEqual(deltas[0],9052)
        self.assertLessEqual(abs(deltas[1]-deltas[2]),1)
        self.assertLessEqual(abs(abs(deltas[3])-9052),1)
        self.assertEqual(raw, initial_raw)
        self.assertAlmostEqual(d.position()['piston_nl_estimate'], 3000.0, delta=1 / 161.36)

    def test_opt_in_invalid_inputs_and_cumulative_bounds(self):
        d=self.drive()
        with self.assertRaises(ValueError): d.move_mm('DV',.01)
        with self.assertRaises(ValueError): d.piston_step('up',10)
        for v in (float('nan'),float('inf'),0,40.01):
            with self.assertRaises(ValueError): d.move_mm('AP',v)
        d.move_mm('AP',.8)
        with self.assertRaises(ValueError): d.move_mm('AP',40)
        self.assertAlmostEqual(d.position()['axes_mm']['AP'],.8,delta=1/5225)

    def test_persistence_restores_nonzero_reference_position(self):
        d=self.drive();d.move_mm('ML',.05);d.close()
        e=StereoDrive(state_path=self.path).connect()
        try: self.assertAlmostEqual(e.position()['axes_mm']['ML'],.05,delta=1/5225)
        finally: e.close()

    def test_state_is_invalid_before_target_and_valid_after_arrival(self):
        d=self.drive(); original=d._session.transport.exchange
        def observe(data,size):
            if data[1]==0x0c:self.assertFalse(json.loads(self.path.read_text())['valid'])
            return original(data,size)
        d._session.transport.exchange=observe
        d.move_mm('AP',.01)
        self.assertTrue(json.loads(self.path.read_text())['valid'])

    def test_external_movement_invalidates_state(self):
        d=self.drive();d._session.transport.raw['AP']+=1
        with self.assertRaises(RuntimeError):d.position()
        self.assertFalse(json.loads(self.path.read_text())['valid'])

    def test_stop_from_other_thread_no_retry_and_busy_rejected(self):
        d=self.drive();errors=[]
        start=d.position()['axes_mm']['AP']
        end=start+.5
        raw_start=d._session.raw['AP']
        from stereodrive_api.protocol import SIGNS, BACKLASH
        raw_target=round(d._session.command_normal['AP'] + SIGNS['AP']*(end-start)*d._session.scales['AP'])
        def move():
            try:d.move_mm('AP',end-start)
            except Exception as e:errors.append(e)
        thread=threading.Thread(target=move);thread.start()
        deadline=time.monotonic()+2
        while 'AP' not in d._session.transport.pending and time.monotonic()<deadline:time.sleep(.005)
        with self.assertRaises(RuntimeError):d.move_mm('ML',.01)
        time.sleep(.2)
        d.stop();thread.join(2)
        self.assertFalse(thread.is_alive());self.assertTrue(errors)
        stopped=d.position()
        position=stopped['axes_mm']['AP']
        self.assertGreater(position,start)
        self.assertLess(position,end)
        state=json.loads(self.path.read_text())
        self.assertTrue(state['valid'])
        self.assertGreater(state['raw']['AP'], min(raw_start,raw_target))
        self.assertLess(state['raw']['AP'], max(raw_start,raw_target))
        self.assertEqual(state['backlash_state']['AP'],BACKLASH['AP'] if SIGNS['AP'] > 0 else 0)
        self.assertTrue(any('MOVE_CANCELLED_RECOVERED' in line for line in (self.path.with_suffix('.jsonl').read_text()).splitlines()))
        d.move_mm('ML',.01)
        self.assertAlmostEqual(d.position()['axes_mm']['ML'],.01,delta=1/5225)
        packets=d._session.transport.packets
        self.assertEqual(sum(p[1]==0x0c for p in packets),2)
        self.assertEqual({p[2] for p in packets if p[1]==0x0f},{0x40,0x50,0x60,0x70})

    def test_stop_during_backlash_takeup_persists_partial_offset_and_restores(self):
        d=self.drive()
        d.move_mm('AP',1.0)
        start=d.position()['axes_mm']['AP']
        errors=[]
        def reverse():
            try:d.move_mm('AP',-.1)
            except Exception as exc:errors.append(exc)
        thread=threading.Thread(target=reverse);thread.start()
        deadline=time.monotonic()+2
        while 'AP' not in d._session.transport.pending and time.monotonic()<deadline:time.sleep(.005)
        time.sleep(.12)
        d.stop();thread.join(2)
        self.assertFalse(thread.is_alive());self.assertTrue(errors)
        saved=json.loads(self.path.read_text())
        partial=saved['backlash_state']['AP']
        self.assertTrue(saved['valid'])
        self.assertGreater(partial,0)
        self.assertLess(partial,BACKLASH['AP'])
        self.assertAlmostEqual(d.position()['axes_mm']['AP'],start,delta=1/5225)
        d.close()

        restored=StereoDrive(state_path=self.path).connect()
        self.addCleanup(restored.close)
        self.assertEqual(restored._session.states['AP'],partial)
        self.assertAlmostEqual(restored.position()['axes_mm']['AP'],start,delta=1/5225)
        restored.move_mm('AP',-.2)
        self.assertAlmostEqual(restored.position()['axes_mm']['AP'],start-.2,delta=1/5225)

    def test_corrupt_ack_stops_and_faults(self):
        d=self.drive(); original=d._session.transport.exchange
        def bad(data,size):
            result=original(data,size)
            return bytes(9) if data[1]==0x0c else result
        d._session.transport.exchange=bad
        with self.assertRaises(RuntimeError):d.move_mm('AP',.01)
        self.assertFalse(json.loads(self.path.read_text())['valid'])
        self.assertEqual(sum(p[1]==0x0c for p in d._session.transport.packets),1)
        records=[json.loads(line) for line in self.path.with_suffix('.jsonl').read_text().splitlines()]
        failed=next(record for record in records if record['event']=='MOVE_FAILED')
        self.assertTrue(failed['error'])
        self.assertTrue(any(record['event']=='STOP_COMMAND_WRITE_COMPLETED' for record in records))

    def test_stop_failure_invalidates_saved_state(self):
        d=self.drive();original=d._session.transport.write
        def fail(data):
            if data[1]==0x0f:raise OSError('Stop write failed')
            return original(data)
        d._session.transport.write=fail
        try:
            with self.assertRaises(RuntimeError):d.stop()
            self.assertFalse(json.loads(self.path.read_text())['valid'])
        finally:d._session.transport.write=original

    def test_invalid_saved_state_is_refused_without_reset(self):
        d=self.drive();d._session.invalidate('interrupted');d.close()
        e=StereoDrive(state_path=self.path)
        with self.assertRaises(RuntimeError):e.connect()

    def test_absolute_calibration_anchors_and_restore(self):
        cal=Calibration(dict(AP=105280,ML=75864,DV=41767), -8572,
                        anchor_backlash=INITIAL)
        d=self.drive(calibration=cal,allow_piston=True)
        p=d.position()
        self.assertTrue(p['calibrated'])
        self.assertEqual(p['axes_mm'],dict(AP=0.,ML=0.,DV=0.))
        self.assertAlmostEqual(p['piston_nl_estimate'],3000)
        d.move_mm('AP',.01);d.piston_step('up',10)
        self.assertAlmostEqual(d.position()['axes_mm']['AP'],.01,delta=1/5225)
        self.assertAlmostEqual(d.position()['piston_nl_estimate'],3010,delta=1/161.36)
        d.close()
        e=StereoDrive(state_path=self.path,calibration=cal).connect()
        try:self.assertAlmostEqual(e.position()['piston_nl_estimate'],3010,delta=1/161.36)
        finally:e.close()
        bad=Calibration(dict(AP=105281,ML=75864,DV=41767),-8572,INITIAL)
        with self.assertRaises(RuntimeError):StereoDrive(state_path=self.path,calibration=bad).connect()

    def test_anchor_backlash_and_validation(self):
        cal=Calibration(dict(AP=105802,ML=75864,DV=41819),-2747,
                        dict(AP=522,ML=261,DV=52,PISTON=5825))
        ref=cal.reference(dict(AP=5225,ML=5225,DV=5225,PISTON=161.36))
        self.assertEqual(ref['AP'],105280)
        self.assertAlmostEqual(ref['PISTON'],-8572-3000*161.36)
        with self.assertRaises(ValueError):Calibration({'AP':0},0).reference({})
        with self.assertRaises(ValueError):StereoDrive(speed_mm_s=3)

    def test_moving_query_during_blocking_move(self):
        d=self.drive();errors=[]
        def move():
            try:d.move_mm('AP',.1)
            except Exception as exc:errors.append(exc)
        thread=threading.Thread(target=move);thread.start()
        deadline=time.monotonic()+2
        while not d._session.busy and time.monotonic()<deadline:time.sleep(.005)
        self.assertTrue(d.is_moving('AP'))
        self.assertFalse(d.is_moving('ML'))
        thread.join(2);self.assertFalse(errors)
        self.assertFalse(d.is_moving())

    def test_drill_commands_query_opt_in_and_stop_off(self):
        d=self.drive()
        with self.assertRaises(ValueError):d.drill_on()
        self.assertFalse(d.drill_off())
        d.close()
        e=StereoDrive(state_path=self.path,allow_drill=True,require_calibration=False).connect()
        try:
            self.assertFalse(e.drill_state())
            self.assertTrue(e.drill_on());self.assertTrue(e.drill_state())
            self.assertIn(bytes.fromhex('af1101'),e._session.transport.packets)
            e.stop();self.assertFalse(e.drill_state())
            self.assertIn(bytes.fromhex('af1100'),e._session.transport.packets)
        finally:e.close()

    def test_drill_timeout_attempts_off_without_on_retry(self):
        d=self.drive(allow_drill=True);original=d._session.transport.exchange
        def unknown(data,size):
            if data==bytes.fromhex('af12'):return bytes(14)
            return original(data,size)
        d._session.transport.exchange=unknown
        try:
            with self.assertRaises(RuntimeError):d.drill_on()
            self.assertEqual(d._session.transport.packets.count(bytes.fromhex('af1101')),1)
            self.assertIn(bytes.fromhex('af1100'),d._session.transport.packets)
            self.assertFalse(json.loads(self.path.read_text())['valid'])
        finally:d._session.transport.exchange=original

if __name__=='__main__': unittest.main()
