"""Blocking, single-owner direct USB API. No connection or motion at import."""
import json
import os
from pathlib import Path
import threading
import time
from datetime import datetime, timezone
from .controller import Session, StateStore
from .protocol import AXES, BACKLASH, drill_power, drill_query, decode_drill
from .calibration import Calibration
from .transport import Simulator, WindowsSerial

class StereoDrive:
    """Use connect(), move_mm(), piston_step(), position(), stop(), close().

    Positions use supplied absolute anchors or a relative saved reference; neither is encoder feedback.
    Call stop() from another thread to cancel a blocking move.
    """
    def __init__(self, *, simulate=True, state_path=None, allow_dv=False,
                 allow_piston=False, allow_drill=False, speed_mm_s=2, calibration=None,
                 counts_per_mm=5225., piston_counts_per_nl=161.36):
        if speed_mm_s not in (1, 2): raise ValueError('Axis speed must be 1 or 2 mm/s')
        if calibration is not None and not isinstance(calibration, Calibration):
            raise TypeError('calibration must be a Calibration object')
        self.calibration = calibration
        self.speed_mm_s = int(speed_mm_s)
        self.allow_drill = bool(allow_drill)
        self.simulate = bool(simulate)
        self.allow_dv, self.allow_piston = bool(allow_dv), bool(allow_piston)
        base = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'StereoDrivePythonAPI'
        self.path = Path(state_path) if state_path else base / ('simulation.json' if simulate else 'live.json')
        self.scales = {a: float(counts_per_mm[a]) for a in ('AP','ML','DV')} if isinstance(counts_per_mm, dict) else dict.fromkeys(('AP','ML','DV'), float(counts_per_mm))
        self.scales['PISTON'] = float(piston_counts_per_nl)
        self._lock = threading.Lock()
        self._stopping = threading.Event()
        self._session = None

    def _log(self, event, **values):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.path.with_suffix('.jsonl').open('a', encoding='utf-8') as f:
            f.write(json.dumps(dict(utc=datetime.now(timezone.utc).isoformat(),
                                   event=event, **values), allow_nan=False) + '\n')

    def connect(self, *, verified_backlash=None, new_reference=False):
        """Restore valid matching state, or explicitly establish a verified reference.

        New references require all four backlash states in motor counts.
        Positive raw direction: maximum backlash; negative raw direction: zero.
        Never use new_reference to recover unknown or interrupted movement.
        """
        with self._lock:
            if self._session is not None: raise RuntimeError('Already connected')
            saved = StateStore(self.path).read()
            if new_reference or saved is None:
                if verified_backlash is None or set(verified_backlash) != set(AXES):
                    raise ValueError('New reference requires verified AP/ML/DV/PISTON backlash states')
                if any(verified_backlash[a] not in (0, BACKLASH[a]) for a in AXES):
                    raise ValueError('Invalid verified backlash states')
            self.path.parent.mkdir(parents=True, exist_ok=True)
            transport = Simulator(self._log, device_file=self.path.with_suffix('.sim-device.json')) if self.simulate else WindowsSerial(self._log)
            try:
                self._session = Session(transport, StateStore(self.path), self.scales,
                                        verified_backlash or {}, self._log, new_reference,
                                        self.calibration.reference(self.scales) if self.calibration else None,
                                        self.speed_mm_s, self.allow_drill)
            except Exception:
                transport.close()
                raise
        return self

    def _require(self):
        if self._session is None: raise RuntimeError('Not connected')
        return self._session

    def position(self):
        """Return calibrated (or relative) axis mm, piston nL estimate, raw counts and backlash."""
        if not self._lock.acquire(blocking=False): raise RuntimeError('Controller busy')
        try:
            s = self._require()
            if s.fault: raise RuntimeError('Session fault: ' + s.fault)
            try: p = s.refresh()
            except Exception as exc:
                try: s.invalidate(exc)
                finally: s.emergency_stop()
                raise
            return dict(axes_mm={a:p[a] for a in ('AP','ML','DV')},
                        piston_nl_estimate=p['PISTON'], raw_counts=dict(s.raw),
                        backlash_counts=dict(s.states), simulated=self.simulate,
                        calibrated=s.absolute_calibration)
        finally: self._lock.release()

    def _move(self, axis, delta):
        delta = float(delta)
        # Validate before accessing hardware or clearing cancellation.
        import math
        if not math.isfinite(delta) or delta == 0: raise ValueError('Finite nonzero movement required')
        if not self._lock.acquire(blocking=False): raise RuntimeError('Controller busy; move rejected')
        try:
            s = self._require()
            if s.fault: raise RuntimeError('Session fault: ' + s.fault)
            s.cancel.clear()
            if self._stopping.is_set(): raise InterruptedError('Stop in progress')
            try: s.move(axis, 1 if delta > 0 else -1, abs(delta))
            except (ValueError,): raise
            except Exception as exc:
                try: s.invalidate(exc)
                finally: s.emergency_stop()
                raise
            return dict(s.positions())
        finally: self._lock.release()

    def move_mm(self, axis, delta_mm):
        """One axis; <=1 mm/action and within +/-1 mm of connection position.

        Uses captured 2 mm/s profile by default (optional 1 mm/s). No automatic speed zones or collision path.
        """
        axis = str(axis).upper()
        if axis not in ('AP','ML','DV'): raise ValueError('Axis must be AP, ML or DV')
        if axis == 'DV' and not self.allow_dv: raise ValueError('DV disabled; opt in after checking clearance')
        return self._move(axis, delta_mm)

    def piston_step(self, direction, volume_nl):
        """Native up/down labels; <=100 nL and within +/-100 nL of connection.

        Supported volumes: 10,20,50,100. Reversal may aspirate fluid/air.
        This uses the captured free-piston profile, not controlled injection rate.
        """
        if not self.allow_piston: raise ValueError('Piston disabled; opt in after verifying syringe setup')
        if direction not in ('up','down') or volume_nl not in (10,20,50,100):
            raise ValueError('Use up/down and 10,20,50,100 nL')
        return self._move('PISTON', volume_nl if direction == 'up' else -volume_nl)

    def moving_status(self):
        """Fresh controller motion flags; callable while a move runs in another thread.

        Reports motor/controller state, not physical displacement or drill RPM.
        Queries share the serial lock with movement polls; never open a second port.
        """
        s = self._require()
        readings = s.snapshot()
        flags = {a: v.moving for a, v in readings.items()}
        return dict(any_moving=any(flags.values()), channels=flags)

    def is_moving(self, channel=None):
        """Return any moving flag, or one AP/ML/DV/PISTON flag."""
        if channel is not None and str(channel).upper() not in AXES:
            raise ValueError('Unknown motion channel')
        status = self.moving_status()
        return status['any_moving'] if channel is None else status['channels'][str(channel).upper()]

    def _drill_state_locked(self, session):
        with session.io_lock:
            return decode_drill(session.transport.exchange(drill_query(), 14))

    def drill_state(self):
        """Reported drill power state; not proof of rotation or RPM."""
        if not self._lock.acquire(blocking=False): raise RuntimeError('Controller busy')
        try: return self._drill_state_locked(self._require())
        finally: self._lock.release()

    def _set_drill(self, enabled):
        if enabled and not self.allow_drill: raise ValueError('Drill ON requires allow_drill=True')
        if not self._lock.acquire(blocking=False): raise RuntimeError('Controller busy; use stop() to cancel')
        try:
            s = self._require()
            if enabled:
                if s.fault: raise RuntimeError('Session fault: ' + s.fault)
                s.refresh()
                if self._stopping.is_set(): raise InterruptedError('Stop in progress')
            try:
                with s.io_lock: s.transport.write(drill_power(enabled))
                deadline = time.monotonic() + 3
                while time.monotonic() < deadline:
                    if self._stopping.is_set() and enabled: raise InterruptedError('Stop requested')
                    if self._drill_state_locked(s) == enabled:
                        self._log('DRILL_POWER_CONFIRMED', enabled=enabled)
                        return enabled
                    time.sleep(.1)
                raise RuntimeError('Drill power state did not match; no automatic ON retry')
            except Exception as exc:
                # A lost ON response must always attempt OFF, including faulted sessions.
                try:
                    with s.io_lock: s.transport.write(drill_power(False))
                finally:
                    try: s.invalidate(exc)
                    finally: s.emergency_stop()
                raise
        finally: self._lock.release()

    def drill_on(self):
        return self._set_drill(True)

    def drill_off(self):
        return self._set_drill(False)

    def _stop_locked(self, session):
        try:
            session.stopped()
            if self.allow_drill:
                deadline = time.monotonic() + 3
                while self._drill_state_locked(session):
                    if time.monotonic() >= deadline: raise RuntimeError('Drill OFF not confirmed after Stop')
                    time.sleep(.1)
        except Exception as exc:
            session.fault = str(exc)
            try: session.invalidate(exc)
            except Exception: pass
            raise

    def stop(self):
        """Cancel from another thread; send Stop once the current exchange exits.

        Best effort software Stop, not a hardware emergency stop.
        """
        s = self._require()
        self._stopping.set()
        s.cancel.set()
        try:
            with self._lock:
                self._stop_locked(s)
        finally: self._stopping.clear()

    def close(self):
        s = self._session
        if s is None: return
        self._stopping.set()
        s.cancel.set()
        with self._lock:
            try: self._stop_locked(s)
            finally:
                try: s.close()
                finally:
                    self._session = None
                    self._stopping.clear()

    def __enter__(self):
        self._require()
        return self

    def __exit__(self, *exc):
        self.close()
