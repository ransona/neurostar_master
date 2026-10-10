"""Blocking, single-owner direct USB API. No connection or motion at import."""
import json
import os
from pathlib import Path
import threading
import time
from datetime import datetime, timezone
from .controller import Session, StateStore, MemoryStateStore
from .protocol import AXES, BACKLASH, SIGNS, drill_power, drill_query, decode_drill
from .calibration import Calibration
from .transport import Simulator, WindowsSerial
from .limits import validate_limits, check_target

class StereoDrive:
    """Use connect(), move_mm(), piston_step(), position(), stop(), close().

    Positions use supplied absolute anchors or a relative saved reference; neither is encoder feedback.
    Call stop() from another thread to cancel a blocking move.
    """
    def __init__(self, *, simulate=True, state_path=None, allow_dv=False,
                 allow_piston=False, allow_drill=False, speed_mm_s=2, calibration=None,
                 counts_per_mm=5225., piston_counts_per_nl=161.36, require_calibration=True, travel_limits=None,
                 simulation_start_position=None, persist_simulation=None):
        if not require_calibration and not simulate:
            raise ValueError('Live movement always requires measured zero calibration.')
        self.require_calibration = bool(require_calibration)
        if isinstance(speed_mm_s, bool) or speed_mm_s not in (1, 2): raise ValueError('Axis speed must be 1 or 2 mm/s')
        if calibration is not None and not isinstance(calibration, Calibration):
            raise TypeError('calibration must be a Calibration object')
        self.calibration = calibration
        self.speed_mm_s = int(speed_mm_s)
        self.travel_limits = validate_limits(travel_limits)
        self.simulation_log = []
        self.simulate = bool(simulate)
        self.persist_simulation = (
            (bool(state_path is not None) if self.simulate else True)
            if persist_simulation is None else bool(persist_simulation)
        )
        if not self.simulate and not self.persist_simulation:
            raise ValueError('Live API state must remain persistent and recoverable.')
        if simulation_start_position is None:
            # Keep explicitly state-pathed legacy protocol fixtures at zero, while
            # the normal throwaway simulator starts at useful nonzero coordinates.
            simulation_start_position = (
                dict(AP=0.0, ML=0.0, DV=0.0, PISTON=0.0) if self.persist_simulation
                else dict(AP=30.0, ML=30.0, DV=30.0, PISTON=2500.0)
            )
        if not isinstance(simulation_start_position, dict) or set(simulation_start_position) != set(AXES):
            raise ValueError('Simulation start position must specify AP, ML, DV in mm and PISTON in nL.')
        import math
        if any(isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value)
               for value in simulation_start_position.values()):
            raise ValueError('Simulation start coordinates must be finite numbers.')
        self.simulation_start_position = {axis: float(value) for axis, value in simulation_start_position.items()}
        for axis in ('AP', 'ML', 'DV', 'PISTON'):
            check_target(self.travel_limits, axis, self.simulation_start_position[axis])
        self.allow_drill = bool(allow_drill) or (self.simulate and not self.persist_simulation)
        self.allow_dv = bool(allow_dv) or (self.simulate and not self.persist_simulation)
        self.allow_piston = bool(allow_piston) or (self.simulate and not self.persist_simulation)
        base = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'StereoDrivePythonAPI'
        self.path = Path(state_path) if state_path else base / ('simulation.json' if simulate else 'live.json')
        self.scales = {a: float(counts_per_mm[a]) for a in ('AP','ML','DV')} if isinstance(counts_per_mm, dict) else dict.fromkeys(('AP','ML','DV'), float(counts_per_mm))
        self.scales['PISTON'] = float(piston_counts_per_nl)
        self._lock = threading.Lock()
        self._stopping = threading.Event()
        self._session = None

    def _log(self, event, **values):
        record = dict(utc=datetime.now(timezone.utc).isoformat(), event=event, **values)
        if self.simulate and not self.persist_simulation:
            self.simulation_log.append(record)
            if len(self.simulation_log) > 10000:
                del self.simulation_log[:len(self.simulation_log) - 10000]
            return
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.path.with_suffix('.jsonl').open('a', encoding='utf-8') as f:
            f.write(json.dumps(record, allow_nan=False) + '\n')

    def _simulation_raw(self, backlash_state):
        # Start from realistic synthetic counts. Absolute calibration, if supplied,
        # is honored; otherwise Session establishes an in-memory relative origin.
        baseline = dict(AP=105280, ML=75864, DV=41767, PISTON=-8572)
        reference = self.calibration.reference(self.scales) if self.calibration else baseline
        return {
            axis: round(reference[axis] + SIGNS[axis] * self.scales[axis] * self.simulation_start_position[axis]
                        + backlash_state[axis])
            for axis in AXES
        }

    def connect(self, *, verified_backlash=None, new_reference=False):
        """Restore valid matching state, or explicitly establish a verified reference.

        New references require all four backlash states in motor counts.
        Positive raw direction: maximum backlash; negative raw direction: zero.
        Never use new_reference to recover unknown or interrupted movement.
        """
        with self._lock:
            if self._session is not None: raise RuntimeError('Already connected')
            store = StateStore(self.path) if (not self.simulate or self.persist_simulation) else MemoryStateStore()
            if self.simulate and not self.persist_simulation and verified_backlash is None:
                verified_backlash = dict.fromkeys(AXES, 0)
            saved = store.read()
            if verified_backlash is not None:
                if not isinstance(verified_backlash, dict) or set(verified_backlash) != set(AXES):
                    raise ValueError('Supply verified AP/ML/DV/PISTON backlash states')
                if any(type(verified_backlash[a]) is not int or verified_backlash[a] not in (0, BACKLASH[a]) for a in AXES):
                    raise ValueError('Invalid verified backlash states')
                if saved and not new_reference and verified_backlash != saved.get('backlash_state'):
                    raise ValueError('Current verified direction history differs from saved state; investigate before reconnecting')
            if (new_reference or saved is None) and not self.simulate:
                if verified_backlash is None or set(verified_backlash) != set(AXES):
                    raise ValueError('New reference requires verified AP/ML/DV/PISTON backlash states')
            if not (self.simulate and not self.persist_simulation):
                self.path.parent.mkdir(parents=True, exist_ok=True)
            transport = (
                Simulator(self._log, raw=self._simulation_raw(verified_backlash or dict.fromkeys(AXES, 0)),
                          device_file=self.path.with_suffix('.sim-device.json') if self.persist_simulation else None)
                if self.simulate else WindowsSerial(self._log)
            )
            try:
                self._session = Session(transport, store, self.scales,
                                        verified_backlash or {}, self._log, new_reference,
                                        self.calibration.reference(self.scales) if self.calibration else None,
                                        self.speed_mm_s, True, self.travel_limits,
                                        self.simulation_start_position if self.simulate else None)
            except Exception:
                transport.close()
                raise
        return self

    def _require(self):
        if self._session is None: raise RuntimeError('Not connected')
        return self._session

    def configure_motion(self, *, speed_mm_s, travel_limits):
        """Change future command profiles/ranges only while verified idle."""
        if isinstance(speed_mm_s, bool) or speed_mm_s not in (1, 2):
            raise ValueError('Axis speed must be 1 or 2 mm/s')
        limits = validate_limits(travel_limits)
        if not self._lock.acquire(blocking=False): raise RuntimeError('Controller busy')
        try:
            s = self._require()
            if s.fault: raise RuntimeError('Session fault: ' + s.fault)
            try: s.refresh()
            except Exception as exc:
                try: s.invalidate(exc)
                finally: s.emergency_stop()
                raise
            self.speed_mm_s = s.speed_mm_s = int(speed_mm_s)
            self.travel_limits = s.travel_limits = limits
        finally: self._lock.release()

    def _require_zero_calibration(self, session):
        if self.require_calibration and not self.simulate and not session.absolute_calibration:
            raise ValueError('Movement disabled: load and verify measured AP/ML/DV zero and piston anchors before connecting for motion.')

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
        if isinstance(delta,bool):raise ValueError('Movement must be a number, not a boolean')
        delta = float(delta)
        # Validate before accessing hardware or clearing cancellation.
        import math
        if not math.isfinite(delta) or delta == 0: raise ValueError('Finite nonzero movement required')
        if not self._lock.acquire(blocking=False): raise RuntimeError('Controller busy; move rejected')
        try:
            s = self._require()
            self._require_zero_calibration(s)
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
        """One axis, constrained by configured mechanical Axis travel ranges.

        Uses captured 2 mm/s profile by default (optional 1 mm/s). No automatic speed zones or collision path.
        """
        axis = str(axis).upper()
        if axis not in ('AP','ML','DV'): raise ValueError('Axis must be AP, ML or DV')
        if axis == 'DV' and not self.allow_dv: raise ValueError('DV disabled; opt in after checking clearance')
        return self._move(axis, delta_mm)

    def piston_step(self, direction, volume_nl):
        """Native up/down labels; captured steps within configured piston range.

        Supported volumes: 10,20,50,100. Reversal may aspirate fluid/air.
        This uses the captured free-piston profile, not controlled injection rate.
        """
        if not self.allow_piston: raise ValueError('Piston disabled; opt in after verifying syringe setup')
        if direction not in ('up','down') or volume_nl not in (10,20,50,100):
            raise ValueError('Use up/down and 10,20,50,100 nL')
        return self._move('PISTON', volume_nl if direction == 'up' else -volume_nl)

    def _axis_plan(self, session, targets, start=None):
        import math
        if not isinstance(targets, dict) or not targets:
            raise ValueError('Supply an AP/ML/DV target mapping')
        self._require_zero_calibration(session)
        if session.fault: raise RuntimeError('Session fault: ' + session.fault)
        start = dict(start or {a:(session.command_normal[a]-session.reference[a])/(SIGNS[a]*session.scales[a]) for a in ('AP','ML','DV')})
        planned = dict(start)
        for axis, value in targets.items():
            if axis not in planned or isinstance(value, bool) or not isinstance(value, (int,float)) or not math.isfinite(value):
                raise ValueError('Absolute targets require finite AP/ML/DV numbers')
            # Held axes are often supplied from rounded motor telemetry. Keep
            # their fractional commanded target within half a motor count;
            # otherwise a held edge-of-envelope axis can spuriously fail.
            planned[axis] = start[axis] if abs(value-start[axis])<=.5/session.scales[axis] else float(value)
        for axis in planned:
            delta = planned[axis] - start[axis]
            if axis == 'DV' and abs(delta) > .5/session.scales[axis] and not self.allow_dv:
                raise ValueError('DV disabled')
            if abs(delta) > .5/session.scales[axis]:
                check_target(session.travel_limits, axis, planned[axis])
            normal=session.reference[axis]+SIGNS[axis]*session.scales[axis]*planned[axis]
            backlash=BACKLASH[axis] if SIGNS[axis]*delta>0 else 0
            if abs(delta)>.5/session.scales[axis] and not -(2**31)<=round(normal+backlash)<2**31:
                raise ValueError('Absolute target exceeds signed 32-bit motor count range')
        return planned

    def validate_axis_path(self, waypoints):
        """Preflight an entire absolute path without sending target packets."""
        if not self._lock.acquire(blocking=False): raise RuntimeError('Controller busy')
        try:
            s=self._require(); s.refresh()
            current=None
            for targets in waypoints: current=self._axis_plan(s,targets,current)
        except ValueError:raise
        except Exception as exc:
            if self._session is not None:
                try:self._session.invalidate(exc)
                finally:self._session.emergency_stop()
            raise
        finally: self._lock.release()

    def validate_piston_steps(self, steps):
        """Preflight every signed free-piston step, without sending targets."""
        if not self._lock.acquire(blocking=False):raise RuntimeError('Controller busy')
        try:
            s=self._require();self._require_zero_calibration(s)
            if not self.allow_piston:raise ValueError('Piston disabled')
            if s.fault:raise RuntimeError('Session fault: '+s.fault)
            s.refresh();self._piston_plan(s, steps)
        except ValueError:raise
        except Exception as exc:
            if self._session is not None:
                try:self._session.invalidate(exc)
                finally:self._session.emergency_stop()
            raise
        finally:self._lock.release()

    def _piston_plan(self, session, steps):
        normal=session.command_normal['PISTON']
        for step in steps:
            if isinstance(step,bool) or step not in (-100,-50,-20,-10,10,20,50,100):
                raise ValueError('Use signed 10/20/50/100 nL free steps')
            normal+=step*session.scales['PISTON']
            estimate=(normal-session.reference['PISTON'])/session.scales['PISTON']
            if not -1e-7<=estimate<=5000+1e-7:raise ValueError('Piston capacity exceeded')
            check_target(session.travel_limits, 'PISTON', estimate)
            raw=round(normal+(BACKLASH['PISTON'] if step>0 else 0))
            if not -(2**31)<=raw<2**31:raise ValueError('Piston raw target overflow')

    def plan_piston_to(self, position_nl):
        """Preflight captured steps toward a calibrated target, without motion.

        Stop short by <10 nL if the target is not reachable with captured units.
        Uses fractional commanded counts, not rounded display telemetry.
        """
        import math
        if isinstance(position_nl,bool) or not isinstance(position_nl,(int,float)) or not math.isfinite(position_nl):
            raise ValueError('Finite piston target required')
        if not self._lock.acquire(blocking=False):raise RuntimeError('Controller busy')
        try:
            s=self._require();self._require_zero_calibration(s)
            if not self.allow_piston:raise ValueError('Piston disabled')
            if s.fault:raise RuntimeError('Session fault: '+s.fault)
            check_target(s.travel_limits,'PISTON',position_nl)
            s.refresh()
            current=(s.command_normal['PISTON']-s.reference['PISTON'])/s.scales['PISTON']
            if not -1e-7<=current<=5000+1e-7:
                raise ValueError('Current piston estimate is outside syringe capacity; verify calibration')
            remaining=10*math.floor(abs(position_nl-current)/10+1e-9)
            sign=1 if position_nl>=current else -1
            steps=[]
            for size in (100,50,20,10):
                count,remaining=divmod(remaining,size)
                steps.extend([sign*size]*count)
            self._piston_plan(s,steps)
            return steps
        except ValueError:raise
        except Exception as exc:
            if self._session is not None:
                try:self._session.invalidate(exc)
                finally:self._session.emergency_stop()
            raise
        finally:self._lock.release()

    def move_axis_to(self, axis, position_mm):
        """Move to a calibrated absolute Axis position; no GUI text boxes."""
        return self.move_axes_to({str(axis).upper():position_mm})

    def move_axes_to(self, targets):
        """Preflight all axes, then sequential DV-retract/AP/ML/DV moves.

        This is not collision planning or a simultaneous path. Each API call
        owns all its legs; Stop cancels remaining legs without a recovery move.
        """
        if not self._lock.acquire(blocking=False): raise RuntimeError('Controller busy')
        try:
            s=self._require(); s.refresh()
            planned=self._axis_plan(s,targets)
            if self._stopping.is_set(): raise InterruptedError('Stop in progress')
            s.cancel.clear()
            current={a:(s.command_normal[a]-s.reference[a])/(SIGNS[a]*s.scales[a]) for a in planned}
            order=['AP','ML','DV']
            if planned['DV'] < current['DV']: order=['DV','AP','ML']
            for axis in order:
                delta=planned[axis]-(s.command_normal[axis]-s.reference[axis])/(SIGNS[axis]*s.scales[axis])
                if abs(delta) <= .5/s.scales[axis]: continue
                if s.cancel.is_set() or self._stopping.is_set(): raise InterruptedError('Stop requested')
                s.move(axis, 1 if delta>0 else -1,abs(delta))
            return dict(s.positions())
        except ValueError: raise
        except Exception as exc:
            if self._session is not None:
                try:self._session.invalidate(exc)
                finally:self._session.emergency_stop()
            raise
        finally:self._lock.release()

    def live_position(self):
        """Motor-derived telemetry during motion; not a verified idle reference."""
        s=self._require()
        if s.fault:raise RuntimeError('Session fault: '+s.fault)
        try:readings=s.snapshot()
        except Exception:
            self.stop()
            raise
        states=dict(s.states)
        active=getattr(s,'active_axis',None)
        if active: states[active]=s.active_backlash
        p={a:(r.raw-states[a]-s.reference[a])/(SIGNS[a]*s.scales[a]) for a,r in readings.items()}
        return dict(axes_mm={a:p[a] for a in ('AP','ML','DV')},piston_nl_estimate=p['PISTON'],
                    raw_counts={a:r.raw for a,r in readings.items()},backlash_counts=states,
                    simulated=self.simulate,calibrated=s.absolute_calibration,
                    verified_idle=not s.busy and not any(r.moving for r in readings.values()))

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
                self._require_zero_calibration(s)
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
