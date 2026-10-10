"""Single-owner motor logic, persistent verified state and bounded moves."""
import json
import math
from pathlib import Path
import threading
import time
from datetime import datetime, timezone
from .protocol import AXES, BACKLASH, SIGNS, STEPS, poll, stop, target, decode_status, validate_ack
from .limits import validate_limits, check_target

class StateStore:
    def __init__(self, path):
        self.path = Path(path)
    def read(self):
        if not self.path.exists(): return None
        data = json.loads(self.path.read_text(encoding='utf-8'))
        if data.get('version') != 1: raise RuntimeError('Unsupported saved state version.')
        return data
    def write(self, data):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix('.tmp')
        # Flush state to disk before a target command is allowed.
        import os
        with tmp.open('w', encoding='utf-8') as handle:
            handle.write(json.dumps(data, indent=2, allow_nan=False))
            handle.flush(); os.fsync(handle.fileno())
        tmp.replace(self.path)

class MemoryStateStore:
    """Session state store for throwaway simulation sessions."""
    def read(self):
        return None
    def write(self, _data):
        return None

class Session:
    def __init__(self, transport, store, scales, initial_states, log, new_reference=False, calibration_reference=None, speed_mm_s=2, drill_enabled=False, travel_limits=None, simulated_initial_position=None):
        self.transport, self.store, self.log = transport, store, log
        self.simulated = transport.identity.startswith('SIMULATOR:')
        self.io_lock = threading.RLock()
        self.speed_mm_s = speed_mm_s
        self.travel_limits = validate_limits(travel_limits)
        self.drill_enabled = drill_enabled
        self.cancel = threading.Event(); self.fault = None; self.busy = False
        self.scales = {a: float(scales[a]) for a in AXES}
        if any(not math.isfinite(v) or not 100 <= v <= 100000 for v in self.scales.values()):
            raise ValueError('Invalid counts/mm calibration.')
        snapshot = self.snapshot()
        if any(s.moving for s in snapshot.values()): raise RuntimeError('Controller is moving; connection cannot initialize.')
        for _ in range(2):
            time.sleep(.1)
            repeat = self.snapshot()
            if repeat != snapshot and any(repeat[a].raw != snapshot[a].raw or repeat[a].moving for a in AXES):
                raise RuntimeError('Motor position changed during initialization.')
        self.raw = {a: snapshot[a].raw for a in AXES}
        saved = self.store.read()
        calibrated = dict(calibration_reference) if calibration_reference is not None else None
        if saved and not new_reference:
            if not saved.get('valid') or saved['identity'] != transport.identity or saved['raw'] != self.raw or saved['scales'] != self.scales:
                raise RuntimeError('Saved state does not match this idle controller. Verify direction history and choose New verified reference.')
            if calibrated is not None and saved['reference_normal'] != calibrated:
                raise RuntimeError('Calibration anchors differ from saved reference')
            if bool(saved.get('absolute_calibration', False)) != (calibrated is not None):
                raise RuntimeError('Supply the same calibration mode used by saved state')
            self.states = saved['backlash_state']; self.reference = saved['reference_normal']
            if any(type(self.states[a]) is not int or self.states[a] not in (0, BACKLASH[a]) for a in AXES): raise RuntimeError('Invalid saved backlash state.')
            if any(not math.isfinite(self.reference[a]) for a in AXES): raise RuntimeError('Invalid saved reference.')
            self.log('RESTORED', raw=self.raw, backlash=self.states)
        else:
            if any(type(initial_states[a]) is not int or initial_states[a] not in (0, BACKLASH[a]) for a in AXES): raise ValueError('Invalid initial backlash state.')
            self.states = dict(initial_states)
            self.reference = calibrated if calibrated is not None else {
                a: self.raw[a] - self.states[a] - SIGNS[a] * self.scales[a] *
                (simulated_initial_position[a] if simulated_initial_position is not None else 0.0)
                for a in AXES
            }
            self.log('NEW_REFERENCE', raw=self.raw, backlash=self.states)
        actual_normal = {a: self.raw[a] - self.states[a] for a in AXES}
        self.command_normal = dict(saved.get('command_normal', actual_normal)) if saved and not new_reference else actual_normal
        if any(not math.isfinite(self.command_normal[a]) or abs(self.command_normal[a]-actual_normal[a]) > .500001 for a in AXES):
            raise RuntimeError('Saved fractional count target is inconsistent with motor position.')
        self.absolute_calibration = calibrated is not None
        self.connection_normal = dict(self.command_normal)
        self.save(True)

    def snapshot(self):
        with self.io_lock:
            return {a: decode_status(a, self.transport.exchange(poll(a), 21 if a == 'AP' else 26)) for a in AXES}

    def positions(self):
        return {a: (self.raw[a] - self.states[a] - self.reference[a]) / (SIGNS[a]*self.scales[a]) for a in AXES}

    def save(self, valid, reason=None):
        self.store.write(dict(version=1, identity=self.transport.identity, valid=valid,
            utc=datetime.now(timezone.utc).isoformat(), raw=self.raw,
            relative_units=self.positions(), backlash_state=self.states,
            absolute_calibration=self.absolute_calibration, reference_normal=self.reference, command_normal=self.command_normal, scales=self.scales, reason=reason))

    def refresh(self):
        readings = self.snapshot()
        if any(s.moving or s.raw != self.raw[a] for a, s in readings.items()):
            raise RuntimeError('Unexpected motor movement/position change. Saved state needs verification.')
        return self.positions()

    def emergency_stop(self):
        errors = []
        for a in AXES:
            try:
                with self.io_lock: self.transport.write(stop(a))
            except Exception as exc: errors.append(f'{a}: {exc}')
        if self.drill_enabled:
            try:
                with self.io_lock: self.transport.write(bytes((0xaf, 0x11, 0)))
            except Exception as exc: errors.append(f'DRILL: {exc}')
        if errors: raise RuntimeError('Stop could not be confirmed: ' + '; '.join(errors))

    def invalidate(self, reason):
        self.fault = str(reason)
        self.save(False, self.fault)

    def move(self, axis, direction, step, publish=lambda readings: None):
        if self.fault: raise RuntimeError('Faulted session; reconnect with verified state.')
        if axis not in AXES or direction not in (-1, 1) or not math.isfinite(step) or step <= 0 or (axis == 'PISTON' and step not in (10,20,50,100)): raise ValueError('Invalid movement.')
        self.refresh()
        old = dict(self.raw); normal = self.command_normal[axis]
        next_normal = normal + SIGNS[axis] * direction * step * self.scales[axis]
        if axis=='PISTON':
            estimate=(next_normal-self.reference[axis])/self.scales[axis]
            if not -1e-7 <= estimate <= 5000+1e-7:
                raise ValueError('Nano 5 µL piston target exceeds 0–5000 nL travel range.')
        estimate = (next_normal-self.reference[axis])/(SIGNS[axis]*self.scales[axis])
        # Relative-only simulation is a protocol test mode, not calibrated travel.
        if axis != 'PISTON' or self.absolute_calibration or self.simulated:
            check_target(self.travel_limits, axis, estimate)
        new_state = BACKLASH[axis] if SIGNS[axis]*direction > 0 else 0
        requested = round(next_normal + new_state)
        if not -(2**31) <= requested < 2**31:
            raise ValueError('Target exceeds signed 32-bit motor count range.')
        # A crash/disconnect anywhere after this write must not restore stale state.
        self.save(False, 'Movement in progress; completion not yet verified.')
        self.busy = True
        self.active_axis, self.active_backlash = axis,new_state
        try:
            if self.cancel.is_set(): raise InterruptedError('Stopped before movement.')
            with self.io_lock:
                reply = self.transport.exchange(target(axis, requested, direction, self.speed_mm_s), 9)
            validate_ack(axis, reply, old[axis], requested)
            deadline = time.monotonic() + (10 if axis == 'PISTON' else max(10, step/self.speed_mm_s*2+5)); settled = None
            while time.monotonic() < deadline:
                if self.cancel.is_set(): raise InterruptedError('Stop requested.')
                readings = self.snapshot()
                if any(readings[a].moving or readings[a].raw != old[a] for a in AXES if a != axis):
                    raise RuntimeError('An uncommanded axis changed.')
                actual = readings[axis]
                if not min(old[axis], requested)-2 <= actual.raw <= max(old[axis], requested)+2:
                    raise RuntimeError('Motor left the commanded raw range.')
                publish(readings)
                if actual.raw == requested and not actual.moving:
                    if settled is None: settled = time.monotonic()
                    elif time.monotonic() - settled >= .2:
                        self.raw[axis] = requested; self.states[axis] = new_state
                        self.command_normal[axis] = next_normal
                        self.save(True); self.log('ARRIVED', axis=axis, raw=requested, relative_units=self.positions())
                        return
                else: settled = None
                time.sleep(.05)
            raise RuntimeError('Movement timed out; no automatic reverse or retry.')
        except InterruptedError as exc:
            if not self.simulated:
                self.fault = str(exc)
                try: self.invalidate(exc)
                except Exception: pass
                try: self.emergency_stop()
                except Exception as stop_exc:
                    try: self.log('STOP_ERROR', error=str(stop_exc))
                    except Exception: pass
                raise
            # A simulator can report an exact stop state, unlike live hardware.
            # Stop all pending axes, then reconcile only the requested axis if
            # its deterministic simulated target had already completed.
            try:
                self.emergency_stop()
                readings = self.snapshot()
                if any(reading.moving for reading in readings.values()):
                    raise RuntimeError('Simulator still reports movement after Stop.')
                reached_target = readings[axis].raw == requested
                if readings[axis].raw not in (old[axis], requested):
                    raise RuntimeError('Simulator stopped at an unknown position.')
                if any(readings[name].raw != old[name] for name in AXES if name != axis):
                    raise RuntimeError('An uncommanded simulated axis changed.')
                if reached_target:
                    self.raw[axis] = requested
                    self.states[axis] = new_state
                    self.command_normal[axis] = next_normal
                self.save(True)
                self.log('MOVE_CANCELLED', axis=axis, reached_target=reached_target,
                         relative_units=self.positions())
            except Exception as stop_exc:
                self.fault = str(stop_exc)
                try: self.invalidate(stop_exc)
                except Exception: pass
                try: self.emergency_stop()
                except Exception: pass
                raise
            raise
        except Exception as exc:
            self.fault = str(exc)
            try: self.invalidate(exc)
            except Exception: pass  # A disk error must never prevent Stop.
            try: self.emergency_stop()
            except Exception as stop_exc:
                try: self.log('STOP_ERROR', error=str(stop_exc))
                except Exception: pass
            raise
        finally:
            self.busy = False
            self.active_axis = None

    def stopped(self):
        """Idle Stop verifies that no motion occurred; interrupted motion stays invalid."""
        self.emergency_stop()
        if not self.fault:
            self.refresh(); self.save(True)

    def close(self):
        try:
            if not self.fault: self.refresh(); self.save(True)
        finally: self.transport.close()
