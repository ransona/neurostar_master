"""Explicit bounded pre-calibration jogging; no invented absolute coordinates."""
import math
import threading
import time
from .controller import StateStore
from .protocol import AXES,BACKLASH,SIGNS,poll,stop,target,decode_status,validate_ack,drill_query,decode_drill
from .transport import Simulator,WindowsSerial

MIN_CALIBRATION_TRAVEL = dict(AP=0.2, ML=0.2, DV=0.2, PISTON=50.0)
INITIAL_DIRECTION_MOVE = dict(AP=0.1, ML=0.1, DV=0.1, PISTON=50.0)

class CalibrationJog:
    def __init__(self,states=None,*,simulate=True,journal_path):
        if states is not None and (set(states)!=set(AXES) or any(type(states[a]) is not int or states[a] not in (0,BACKLASH[a]) for a in AXES)):
            raise ValueError('Verify all four direction histories before calibration jogging')
        self.lock=threading.Lock();self.cancel=threading.Event();self.fault=None
        self.states=dict(states) if states is not None else None;self.store=StateStore(journal_path)
        self.transport=Simulator(lambda *a,**k:None) if simulate else WindowsSerial(lambda *a,**k:None)
        self.closed=False
        try:
            first=self._snapshot();time.sleep(.2);second=self._snapshot()
            if any(first[a].moving or second[a].moving or first[a].raw!=second[a].raw for a in AXES):raise RuntimeError('Controller must be stationary')
            self.raw={a:second[a].raw for a in AXES}
            self.normal={a:self.raw[a]-((self.states or {}).get(a) or 0) for a in AXES}
            self.origin=dict(self.normal)
            self.travelled=dict.fromkeys(AXES,0.0)
            self._save(self.states is not None)
        except Exception:self.transport.close();raise

    def _snapshot(self):
        return {a:decode_status(a,self.transport.exchange(poll(a),21 if a=='AP' else 26)) for a in AXES}

    def _save(self,valid):
        history_complete=self.states is not None and all(self.states[a] is not None for a in AXES)
        self.store.write(dict(version=1,mode='relative-calibration-jog',valid=bool(valid and history_complete),
            identity=self.transport.identity,raw_counts=self.raw,backlash_counts=self.states,
            normal_counts=self.normal,origin_counts=self.origin,fault=self.fault))

    def capture_calibration_counts(self):
        if not self.lock.acquire(False):raise RuntimeError('Jog busy')
        try:
            if self.closed or self.fault:raise RuntimeError('Calibration connection closed or faulted')
            one=self._snapshot();time.sleep(.2);two=self._snapshot()
            if any(one[a].moving or two[a].moving or one[a].raw!=two[a].raw or two[a].raw!=self.raw[a] for a in AXES):
                raise RuntimeError('Calibration counts not stationary or changed unexpectedly')
            return dict(raw_counts=dict(self.raw),backlash_counts=dict(self.states) if self.states is not None else None,
                        travelled=dict(self.travelled))
        finally:self.lock.release()

    def set_initial_states(self,states):
        if set(states)!=set(AXES) or any(type(states[a]) is not int or states[a] not in (0,BACKLASH[a]) for a in AXES):
            raise ValueError('Verify all four direction histories before calibration jogging')
        if not self.lock.acquire(False):raise RuntimeError('Jog busy')
        try:
            if self.closed or self.fault:raise RuntimeError('Calibration connection closed or faulted')
            if self.states is not None and self.states!=states:raise RuntimeError('Backlash history is already established for this session')
            self.states=dict(states);self.normal={a:self.raw[a]-self.states[a] for a in AXES};self.origin=dict(self.normal)
            self._save(True)
        finally:self.lock.release()

    def validate_move(self,axis,delta):
        if axis not in AXES or isinstance(delta,bool) or not isinstance(delta,(int,float)) or not math.isfinite(delta) or delta==0:
            raise ValueError('Invalid jog')
        cap=100 if axis=='PISTON' else 2
        if abs(delta)>cap:raise ValueError(f'Calibration move is limited to +/-{cap:g} {"nL" if axis=="PISTON" else "mm"}.')
        if not self.lock.acquire(False):raise RuntimeError('Jog busy')
        try:
            if self.closed or self.fault:raise RuntimeError('Calibration connection closed or faulted')
            if self.cancel.is_set():raise InterruptedError('Stop latched; reopen calibration after verifying position/history')
            if self.states is None or any(self.states[a] is None for a in AXES):
                raise RuntimeError('Initialize all movement directions before jogging')
            scale=161.36 if axis=='PISTON' else 5225
            normal=self.normal[axis]+SIGNS[axis]*scale*delta
            envelope=300 if axis=='PISTON' else 3
            if abs(normal-self.origin[axis])/scale>envelope+1e-7:
                raise ValueError(f'Calibration travel envelope: +/-{envelope:g} {"nL" if axis=="PISTON" else "mm"} from connection')
            state=BACKLASH[axis] if SIGNS[axis]*delta>0 else 0
            requested=round(normal+state)
            if not -(2**31)<=requested<2**31:raise ValueError('Target exceeds signed 32-bit motor count range.')
            return True
        finally:self.lock.release()

    def move(self,axis,delta,speed_mm_s=1):
        if axis not in AXES or isinstance(delta,bool) or not isinstance(delta,(int,float)) or not math.isfinite(delta) or delta==0:raise ValueError('Invalid jog')
        if isinstance(speed_mm_s,bool) or speed_mm_s not in (1,2):raise ValueError('Axis speed must be 1 or 2 mm/s')
        cap=100 if axis=='PISTON' else 2
        if abs(delta)>cap:raise ValueError('Jog too large')
        if not self.lock.acquire(False):raise RuntimeError('Jog busy')
        try:
            if self.closed or self.fault:raise RuntimeError('Calibration connection closed or faulted')
            if self.cancel.is_set():raise InterruptedError('Stop latched; reopen calibration after verifying position/history')
            if decode_drill(self.transport.exchange(drill_query(),14)):raise RuntimeError('Drill must be OFF before calibration jogging')
            before=self._snapshot()
            if any(before[a].moving or before[a].raw!=self.raw[a] for a in AXES):raise RuntimeError('Unexpected motor state')
            scale=161.36 if axis=='PISTON' else 5225
            normal=self.normal[axis]+SIGNS[axis]*scale*delta
            envelope=300 if axis=='PISTON' else 3
            if abs(normal-self.origin[axis])/scale>envelope+1e-7:raise ValueError(f'Calibration travel envelope: +/-{envelope:g} {"nL" if axis=="PISTON" else "mm"} from connection')
            state=BACKLASH[axis] if SIGNS[axis]*delta>0 else 0;requested=round(normal+state)
            self._save(False)
            try:
                if self.cancel.is_set():raise InterruptedError('Stop requested')
                reply=self.transport.exchange(target(axis,requested,1 if delta>0 else -1,speed_mm_s),9)
                validate_ack(axis,reply,self.raw[axis],requested)
                end=time.monotonic()+10;settled=None
                while time.monotonic()<end:
                    if self.cancel.is_set():raise InterruptedError('Stop requested')
                    readings=self._snapshot()
                    if any(readings[a].moving or readings[a].raw!=self.raw[a] for a in AXES if a!=axis):raise RuntimeError('Uncommanded channel moved')
                    current=readings[axis]
                    if not min(self.raw[axis],requested)-2<=current.raw<=max(self.raw[axis],requested)+2:raise RuntimeError('Motor left requested range')
                    if not current.moving and current.raw==requested:
                        if settled is None:settled=time.monotonic()
                        elif time.monotonic()-settled>=.2:
                            self.raw[axis]=requested;self.normal[axis]=normal
                            if self.states is None:self.states=dict.fromkeys(AXES)
                            self.states[axis]=state
                            self.travelled[axis]=max(self.travelled[axis],abs(normal-self.origin[axis])/scale)
                            self._save(True);return
                    else:settled=None
                    time.sleep(.05)
                raise RuntimeError('Jog timeout')
            except Exception as exc:
                self.fault=str(exc)
                try:self._save(False)
                finally:self._stop_packets()
                raise
        finally:self.lock.release()

    def _stop_packets(self):
        errors=[]
        for a in AXES:
            try:self.transport.write(stop(a))
            except Exception as exc:errors.append(str(exc))
        try:self.transport.write(bytes((0xaf,0x11,0)))
        except Exception as exc:errors.append(str(exc))
        if errors:raise RuntimeError('Stop failed: '+'; '.join(errors))

    def stop(self):
        self.cancel.set()
        with self.lock:self._stop_packets()

    def close(self):
        if self.closed:return
        try:self.stop()
        finally:self.closed=True;self.transport.close()
