"""Single-owner motor logic, persistent verified state and bounded moves."""
import json
import math
from pathlib import Path
import threading
import time
from datetime import datetime, timezone
from .protocol import (AXES, BACKLASH, SIGNS, STEPS, poll, stop, target, decode_status,
                       validate_ack, drill_query, decode_drill, drill_power)
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
    def __init__(self, transport, store, scales, initial_states, log, new_reference=False, calibration_reference=None, speed_mm_s=2, drill_enabled=False, travel_limits=None, simulated_initial_position=None, drill_state_timeout_seconds=3.0, drill_state_retries=4):
        self.transport, self.store, self.log = transport, store, log
        self.simulated = transport.identity.startswith('SIMULATOR:')
        self.io_lock = threading.RLock()
        self.speed_mm_s = speed_mm_s
        self.travel_limits = validate_limits(travel_limits)
        self.drill_enabled = drill_enabled
        self.drill_state_timeout_seconds = float(drill_state_timeout_seconds)
        self.drill_state_retries = int(drill_state_retries)
        self.cancel = threading.Event(); self.fault = None; self.busy = False
        self.invalidation_scope = None
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
            recovery_failure = None
            if not saved.get('valid') and self._is_drill_only_invalidation(saved):
                recovery_failure = self._try_restore_drill_only_state(saved, calibrated)
                if recovery_failure is None:
                    saved = self.store.read()
            if not saved.get('valid') or saved['identity'] != transport.identity or saved['raw'] != self.raw or saved['scales'] != self.scales:
                detail = f' Drill-only recovery refused: {recovery_failure}.' if recovery_failure else ''
                raise RuntimeError('Saved state does not match this idle controller. Verify direction history and choose New verified reference.' + detail)
            if calibrated is not None and saved['reference_normal'] != calibrated:
                raise RuntimeError('Calibration anchors differ from saved reference')
            if bool(saved.get('absolute_calibration', False)) != (calibrated is not None):
                raise RuntimeError('Supply the same calibration mode used by saved state')
            self.states = saved['backlash_state']; self.reference = saved['reference_normal']
            if any(type(self.states[a]) is not int or not 0 <= self.states[a] <= BACKLASH[a] for a in AXES): raise RuntimeError('Invalid saved backlash state.')
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

    @staticmethod
    def _is_drill_only_invalidation(saved):
        reason = str(saved.get('reason') or '')
        # The reason prefix supports state files written before invalidation scopes
        # were added. Every recovery still requires fresh matching hardware checks.
        return (saved.get('invalidated_by') == 'drill_power'
                or reason.startswith('Drill power state did not match'))

    def _try_restore_drill_only_state(self, saved, calibrated):
        """Revalidate saved position state only after fresh, exact idle checks."""
        def refuse(reason):
            self.log('DRILL_ONLY_STATE_RESTORE_REFUSED', reason=reason)
            return reason

        self.log('DRILL_ONLY_STATE_RESTORE_ATTEMPTED', identity=self.transport.identity)
        if saved.get('identity') != self.transport.identity:
            return refuse('controller identity changed')
        if saved.get('scales') != self.scales:
            return refuse('motion scales changed')
        if not isinstance(saved.get('raw'), dict) or saved['raw'] != self.raw:
            return refuse('current motor counts differ from saved counts')
        states = saved.get('backlash_state')
        if (not isinstance(states, dict) or set(states) != set(AXES)
                or any(type(states[a]) is not int or not 0 <= states[a] <= BACKLASH[a] for a in AXES)):
            return refuse('saved direction history is invalid')
        reference = saved.get('reference_normal')
        if not isinstance(reference, dict) or set(reference) != set(AXES):
            return refuse('saved calibration reference is incomplete')
        if any(not isinstance(reference[a], (int, float)) or not math.isfinite(reference[a]) for a in AXES):
            return refuse('saved calibration reference is not finite')
        if bool(saved.get('absolute_calibration', False)) != (calibrated is not None):
            return refuse('calibration mode differs from saved state')
        if calibrated is not None and reference != calibrated:
            return refuse('measured calibration anchors differ from saved reference')
        command_normal = saved.get('command_normal')
        if not isinstance(command_normal, dict) or set(command_normal) != set(AXES):
            return refuse('saved commanded position is incomplete')
        actual_normal = {a: self.raw[a] - states[a] for a in AXES}
        if any(not isinstance(command_normal[a], (int, float)) or not math.isfinite(command_normal[a])
               or abs(command_normal[a] - actual_normal[a]) > .500001 for a in AXES):
            return refuse('saved commanded position does not match current counts')
        try:
            drill_on = decode_drill(self.transport.exchange(drill_query(), 14))
            if drill_on:
                confirmed_off = False
                for attempt in range(self.drill_state_retries + 1):
                    self.log('DRILL_ONLY_RECOVERY_OFF_ATTEMPT', attempt=attempt + 1,
                             max_attempts=self.drill_state_retries + 1)
                    with self.io_lock:
                        self.transport.write(drill_power(False))
                    deadline = time.monotonic() + self.drill_state_timeout_seconds
                    while time.monotonic() < deadline:
                        if not decode_drill(self.transport.exchange(drill_query(), 14)):
                            confirmed_off = True
                            self.log('DRILL_ONLY_RECOVERY_OFF_CONFIRMED', attempt=attempt + 1)
                            break
                        time.sleep(.1)
                    if confirmed_off:
                        break
                if not confirmed_off:
                    return refuse(f'could not confirm drill OFF after {self.drill_state_retries + 1} attempts')
        except Exception as exc:
            return refuse(f'could not verify or turn drill OFF: {exc}')

        restored = dict(saved, valid=True, reason=None)
        restored.pop('invalidated_by', None)
        self.store.write(restored)
        self.log('DRILL_ONLY_STATE_RESTORED', raw=self.raw,
                 note='Fresh idle counts exactly matched the saved position and drill was reported OFF.')
        return None

    def snapshot(self):
        with self.io_lock:
            return {a: decode_status(a, self.transport.exchange(poll(a), 21 if a == 'AP' else 26)) for a in AXES}

    def positions(self):
        return {a: (self.raw[a] - self.states[a] - self.reference[a]) / (SIGNS[a]*self.scales[a]) for a in AXES}

    def save(self, valid, reason=None, invalidated_by=None):
        state = dict(version=1, identity=self.transport.identity, valid=valid,
            utc=datetime.now(timezone.utc).isoformat(), raw=self.raw,
            relative_units=self.positions(), backlash_state=self.states,
            absolute_calibration=self.absolute_calibration, reference_normal=self.reference,
            command_normal=self.command_normal, scales=self.scales, reason=reason)
        if invalidated_by is not None:
            state['invalidated_by'] = invalidated_by
        self.store.write(state)

    def refresh(self):
        readings = self.snapshot()
        if any(s.moving or s.raw != self.raw[a] for a, s in readings.items()):
            raise RuntimeError('Unexpected motor movement/position change. Saved state needs verification.')
        return self.positions()

    def emergency_stop(self):
        errors = []
        def record(event, **details):
            try: self.log(event, **details)
            except Exception: pass

        record('STOP_REQUESTED', axes=list(AXES), drill_enabled=self.drill_enabled)
        for a in AXES:
            try:
                with self.io_lock: self.transport.write(stop(a))
                record('STOP_COMMAND_WRITE_COMPLETED', axis=a)
            except Exception as exc:
                errors.append(f'{a}: {exc}')
                record('STOP_COMMAND_WRITE_FAILED', axis=a, error=str(exc))
        if self.drill_enabled:
            try:
                with self.io_lock: self.transport.write(bytes((0xaf, 0x11, 0)))
                record('DRILL_OFF_COMMAND_WRITE_COMPLETED', source='emergency_stop')
            except Exception as exc:
                errors.append(f'DRILL: {exc}')
                record('DRILL_OFF_COMMAND_WRITE_FAILED', source='emergency_stop', error=str(exc))
        if errors:
            record('STOP_COMMANDS_FAILED', errors=errors)
            raise RuntimeError('Stop could not be confirmed: ' + '; '.join(errors))
        record('STOP_COMMAND_WRITES_COMPLETED', axes=list(AXES), drill_enabled=self.drill_enabled,
               note='Command writes completed; stationary hardware state is not confirmed by this event.')

    def invalidate(self, reason, *, scope=None):
        self.fault = str(reason)
        self.invalidation_scope = scope
        self.save(False, self.fault, invalidated_by=scope)

    def move(self, axis, direction, step, publish=lambda readings: None):
        if self.fault: raise RuntimeError('Faulted session; reconnect with verified state.')
        if axis not in AXES or direction not in (-1, 1) or not math.isfinite(step) or step <= 0 or (axis == 'PISTON' and step not in (10,20,50,100)): raise ValueError('Invalid movement.')
        self.refresh()
        old = dict(self.raw); normal = self.command_normal[axis]
        next_normal = normal + SIGNS[axis] * direction * step * self.scales[axis]
        if axis=='PISTON':
            low, high = self.travel_limits['PISTON']
            current = (normal-self.reference[axis])/self.scales[axis]
            if not low-1e-7 <= current <= high+1e-7:
                raise ValueError(f'Current piston position is outside configured {low:g}–{high:g} nL travel range.')
            estimate=(next_normal-self.reference[axis])/self.scales[axis]
            if not low-1e-7 <= estimate <= high+1e-7:
                raise ValueError(f'Piston target exceeds configured {low:g}–{high:g} nL travel range.')
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
        self.log(
            'MOVE_REQUESTED',
            axis=axis,
            direction=direction,
            step=step,
            speed_mm_s=self.speed_mm_s,
            start_position=self.positions(),
            target_position={**self.positions(), axis: estimate},
            start_raw=old[axis],
            target_raw=requested,
        )
        self.busy = True
        self.active_axis, self.active_backlash = axis,new_state
        move_started = time.monotonic()
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
            try:
                self.emergency_stop()
                readings = self._read_stable_stopped_position(old, requested, axis)
                actual = readings[axis].raw
                old_state = self.states[axis]
                raw_delta = actual - old[axis]
                raw_direction = 1 if requested > old[axis] else -1
                if raw_delta * raw_direction < 0:
                    raise RuntimeError('Stopped count moved opposite to the commanded direction.')
                state_delta = new_state - old_state
                if state_delta * raw_direction < 0:
                    raise RuntimeError('Backlash take-up direction is inconsistent with the stopped count.')
                takeup = abs(state_delta)
                if abs(raw_delta) <= takeup:
                    actual_state = old_state + raw_delta
                    actual_normal = normal
                else:
                    actual_state = new_state
                    actual_normal = normal + raw_delta - state_delta
                if not 0 <= actual_state <= BACKLASH[axis]:
                    raise RuntimeError('Stopped backlash offset lies outside its calibrated range.')
                expected_travel = next_normal - normal
                actual_travel = actual_normal - normal
                if actual_travel * expected_travel < 0 or abs(actual_travel) > abs(expected_travel) + 1.0:
                    raise RuntimeError('Stopped count cannot be reconciled with the commanded travel.')
                self.raw.update({name: readings[name].raw for name in AXES})
                self.states[axis] = actual_state
                self.command_normal[axis] = actual_normal
                self.save(True)
                self.log('MOVE_CANCELLED_RECOVERED', axis=axis, start_raw=old[axis],
                         target_raw=requested, stopped_raw=actual, backlash_state=actual_state,
                         position=self.positions()[axis], relative_units=self.positions())
            except Exception as stop_exc:
                self.fault = str(stop_exc)
                try: self.log('MOVE_CANCEL_RECONCILIATION_FAILED', axis=axis, error=str(stop_exc))
                except Exception: pass
                try: self.invalidate(stop_exc)
                except Exception: pass
                try: self.emergency_stop()
                except Exception: pass
                raise
            raise
        except Exception as exc:
            self.fault = str(exc)
            try:
                self.log('MOVE_FAILED', axis=axis, direction=direction, step=step,
                         start_raw=old[axis], target_raw=requested, error_type=type(exc).__name__,
                         error=str(exc), elapsed_seconds=time.monotonic() - move_started)
            except Exception: pass
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

    def _read_stable_stopped_position(self, old, requested, active_axis,
                                      *, timeout_seconds=3.0, stable_samples=3):
        """Require stable idle telemetry and counts bounded by the in-flight target."""
        deadline = time.monotonic() + timeout_seconds
        previous = None
        stable = 0
        last = None
        while time.monotonic() < deadline:
            readings = self.snapshot()
            if any(reading.moving for reading in readings.values()):
                stable = 0
                previous = None
            else:
                current = {axis: readings[axis].raw for axis in AXES}
                if any(current[axis] != old[axis] for axis in AXES if axis != active_axis):
                    raise RuntimeError('An uncommanded axis changed during stopped movement.')
                count = current[active_axis]
                if not min(old[active_axis], requested) - 2 <= count <= max(old[active_axis], requested) + 2:
                    raise RuntimeError('Stopped motor count is outside the commanded raw range.')
                stable = stable + 1 if current == previous else 1
                previous = current
                last = readings
                if stable >= stable_samples:
                    return last
            time.sleep(.05)
        raise RuntimeError('Could not verify stable idle counts after Stop.')

    def stopped(self):
        """Idle Stop verifies that no motion occurred; interrupted motion stays invalid."""
        self.emergency_stop()
        if not self.fault:
            try:
                position = self.refresh()
                self.save(True)
                self.log('STOP_STATE_VERIFIED', relative_units=position)
            except Exception as exc:
                try: self.log('STOP_STATE_VERIFICATION_FAILED', error_type=type(exc).__name__, error=str(exc))
                except Exception: pass
                raise

    def close(self):
        try:
            if not self.fault: self.refresh(); self.save(True)
        finally: self.transport.close()
