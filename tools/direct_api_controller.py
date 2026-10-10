"""Planner adapter: every physical command uses stereodrive_api, never native UI."""
import json
import math
from pathlib import Path
import sys
import threading
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from stereodrive_api import StereoDrive, Calibration


class StereoDriveError(RuntimeError):
    pass


class StereoDriveController:
    direct_api = True
    supports_drilling_protocol = False
    supports_injection_protocol = False
    pulsed_protocol = True

    def __init__(self, *, live=False):
        self.live = live
        self.drive = None
        self.origin = None
        self.steps = dict.fromkeys(("AP", "ML", "DV"), .01)
        self.cancelled = threading.Event()
        self.lock = threading.Lock()
        self.busy = False
        self.worker = None
        self.error = None
        self.directions = {}
        self.last_rejection = None

    def connect(self, calibration, states, state_path, *, new_reference=False,
                allow_dv=False, allow_piston=False, allow_drill=False, speed=1,allow_pulsed=False,travel_limits=None):
        if type(allow_pulsed) is not bool:raise StereoDriveError('Pulsed workflow opt-in must be an explicit boolean')
        if self.drive is not None:
            raise StereoDriveError("Disconnect first; calibration cannot change while connected.")
        if not isinstance(calibration, Calibration):
            raise StereoDriveError("Load measured zero calibration before connecting.")
        drive = StereoDrive(simulate=not self.live, calibration=calibration,
            state_path=state_path, allow_dv=allow_dv, allow_piston=allow_piston,
            allow_drill=allow_drill, speed_mm_s=speed, travel_limits=travel_limits)
        try:
            drive.connect(verified_backlash=states, new_reference=new_reference)
            position = drive.position()["axes_mm"]
            self.origin = tuple(position[a] for a in ("AP","ML","DV"))
        except Exception:
            drive.close()
            raise
        self.drive = drive
        self.supports_drilling_protocol=self.supports_injection_protocol=bool(allow_pulsed)
        self.error = None
        self.last_rejection = None
        self.cancelled.clear()

    def _require(self):
        if self.drive is None:
            raise StereoDriveError("Direct USB disconnected. Use USB setup: verify zero calibration and direction history first.")
        if self.error:
            raise StereoDriveError("Direct USB fault: " + self.error + ". Verify physical position/history before reconnecting.")
        return self.drive

    def prepare_motion(self):
        self._require()
        if self.has_active_motion(): raise StereoDriveError("Movement already in progress.")
        self.cancelled.clear()

    def _check_motion_cancelled(self, stop_requested=None):
        if self.cancelled.is_set() or (stop_requested and stop_requested()):
            raise StereoDriveError("Movement cancelled.")

    def _run(self, action, *, asynchronous=False, directions=None, stop_requested=None):
        drive = self._require()
        with self.lock:
            if self.busy: raise StereoDriveError("Controller busy; request rejected.")
            self._check_motion_cancelled(stop_requested)
            self.busy = True
            self.directions = directions or {}
        def execute():
            finished = threading.Event()
            def watch_cancel():
                while not finished.wait(.03):
                    if self.cancelled.is_set() or (stop_requested and stop_requested()):
                        try: drive.stop()
                        except Exception as exc: self.error = str(exc)
                        return
            watcher = threading.Thread(target=watch_cancel, daemon=True)
            watcher.start()
            try:
                self._check_motion_cancelled(stop_requested)
                return action(drive)
            except ValueError as exc:
                if not asynchronous: raise StereoDriveError(str(exc)) from exc
                # Validation failure is visible but is not a lost physical reference.
                self.last_rejection = str(exc)
            except Exception as exc:
                cancelled=self.cancelled.is_set() or (stop_requested and stop_requested())
                if cancelled:
                    try:drive.position() # An idle cancellation between packets need not lose calibration.
                    except Exception:self.error=str(exc)
                else:self.error = str(exc)
                if not asynchronous: raise StereoDriveError(str(exc)) from exc
            finally:
                finished.set()
                watcher.join(timeout=3)
                with self.lock:
                    self.busy = False
                    self.directions = {}
        if asynchronous:
            self.worker = threading.Thread(target=execute, daemon=True)
            self.worker.start()
        else:
            return execute()

    def get_current_axis_position(self):
        if self.last_rejection:
            message=self.last_rejection; self.last_rejection=None
            raise StereoDriveError(message)
        drive=self._require()
        # Idle reads and preflight share the adapter gate with command startup.
        # A display read must not cause a false "Controller busy" protocol failure.
        with self.lock:
            p=drive.live_position() if self.busy else drive.position()
        return tuple(p["axes_mm"][a] for a in ("AP","ML","DV"))

    def get_current_axis(self, axis):
        return self.get_current_axis_position()[("AP","ML","DV").index(axis.upper())]

    def get_motion_direction(self, axis):
        return self.directions.get(axis.upper()) if self.busy else None

    def has_active_motion(self):
        return self.busy

    def set_nudge_step(self, axis, step_mm):
        if axis not in self.steps or not math.isfinite(step_mm) or not 0 < step_mm:
            raise StereoDriveError("Direct API steps must be finite and positive.")
        self.steps[axis]=step_mm

    def configure_motion(self, speed, limits):
        with self.lock:
            if self.busy: raise StereoDriveError('Controller busy; settings unchanged')
            self._require().configure_motion(speed_mm_s=speed, travel_limits=limits)

    def nudge_axis(self, axis, positive, **kwargs):
        axis=axis.upper()
        delta=self.steps[axis]*(1 if positive else -1)
        # Preflight before the asynchronous worker so rejection is immediate.
        drive=self._require()
        p=dict(zip(('AP','ML','DV'),self.get_current_axis_position()))
        p[axis]+=delta
        self.validate_axis_path([tuple(p[a] for a in ('AP','ML','DV'))])
        return self._run(lambda d:d.move_mm(axis,delta), asynchronous=True, directions={axis:positive})

    def validate_axis_path(self, positions):
        with self.lock:
            if self.busy:raise StereoDriveError('Controller busy; path preflight rejected')
            return self._require().validate_axis_path([dict(zip(("AP","ML","DV"),p)) for p in positions])

    def validate_piston_steps(self, steps):
        with self.lock:
            if self.busy:raise StereoDriveError('Controller busy; dose preflight rejected')
            return self._require().validate_piston_steps(steps)

    def goto_axis_position(self, ap, ml, dv, delay_seconds=0, stop_requested=None):
        self.validate_axis_path([(ap,ml,dv)])
        current=self.get_current_axis_position()
        directions={a:t>c for a,t,c in zip(("AP","ML","DV"),(ap,ml,dv),current) if abs(t-c)>.0001}
        return self._run(lambda d:d.move_axes_to(dict(AP=ap,ML=ml,DV=dv)), directions=directions,
                         stop_requested=stop_requested)

    def wait_for_axis_position(self, ap, ml, dv, tolerance_mm=.003, timeout_seconds=15,
                               stop_requested=None, position_callback=None, **kwargs):
        self.wait_until_stopped(timeout_seconds)
        self._check_motion_cancelled(stop_requested)
        p=self.get_current_axis_position()
        if position_callback: position_callback(p)
        if any(abs(a-b)>tolerance_mm for a,b in zip(p,(ap,ml,dv))):
            raise StereoDriveError("Direct API did not verify the requested calibrated target.")

    def move_axis_to_target(self, axis, target, stop_requested=None, step_mm=.005,dwell_seconds=0,**kwargs):
        p=list(self.get_current_axis_position())
        p[("AP","ML","DV").index(axis.upper())]=target
        return self.move_to_position_nudged(*p,step_mm=step_mm,dwell_seconds=dwell_seconds,stop_requested=stop_requested)

    def move_to_position_nudged(self, ap, ml, dv, stop_requested=None, step_mm=.005,dwell_seconds=0,**kwargs):
        from pulsed_protocol import line,finite
        dwell_seconds=finite(dwell_seconds,'pulse period')
        start=self.get_current_axis_position()
        points=line(start,(ap,ml,dv),step_mm)
        self.validate_axis_path(points)
        def move(d):
            previous=start
            for target in points:
                self._check_motion_cancelled(stop_requested)
                started=time.monotonic()
                result=d.move_axes_to(dict(zip(('AP','ML','DV'),target)))
                period=dwell_seconds*math.dist(previous,target)/step_mm
                while time.monotonic()-started<period:
                    self._check_motion_cancelled(stop_requested)
                    time.sleep(max(0,min(.03,period-(time.monotonic()-started))))
                previous=target
            return result
        directions={a:t>c for a,t,c in zip(('AP','ML','DV'),(ap,ml,dv),start) if abs(t-c)>.0001}
        return self._run(move,directions=directions,stop_requested=stop_requested)

    def syringe_step(self, volume_label, up=True, stop_requested=None, asynchronous=False, on_completed=None):
        try: volume=float(volume_label.lower().replace("nl", "").strip())
        except ValueError as exc: raise StereoDriveError("Invalid syringe volume") from exc
        steps=self.syringe_volume_steps(volume,up)
        self.validate_piston_steps(steps)
        return self._run(lambda d:self._execute_piston_steps(d,steps,stop_requested,on_completed),
                         stop_requested=stop_requested,asynchronous=asynchronous)

    @staticmethod
    def syringe_volume_steps(volume,up):
        if isinstance(volume,bool) or not isinstance(volume,(int,float)) or not math.isfinite(volume) or not 10<=volume<=5000 or volume%10:
            raise StereoDriveError('Syringe volume must be a multiple of 10 nL within 10–5000 nL')
        remaining=int(volume);steps=[]
        for size in (100,50,20,10):
            count,remaining=divmod(remaining,size)
            steps.extend([(1 if up else -1)*size]*count)
        return steps

    def _execute_piston_steps(self,drive,steps,stop_requested,on_completed):
        drive.validate_piston_steps(steps) # Recheck at exclusive command ownership.
        result={'PISTON':drive.position()['piston_nl_estimate']}
        if not steps and on_completed:on_completed(result['PISTON'])
        for step in steps:
            self._check_motion_cancelled(stop_requested)
            result=drive.piston_step('up' if step>0 else 'down',abs(step))
            if on_completed:on_completed(result['PISTON'])
        return result

    def read_injectomate_calibrate_scale_nl(self, **kwargs):
        with self.lock:
            if self.busy:raise StereoDriveError('Wait for verified idle piston position')
            return self._require().position()["piston_nl_estimate"]

    def reported_drill_state(self):
        with self.lock:
            if self.busy:raise StereoDriveError('Controller busy; drill query rejected')
            return self._require().drill_state()

    def turn_drill_off(self):
        return self._run(lambda d:d.drill_off())

    def stop(self):
        self.cancelled.set()
        if self.drive: self.drive.stop()

    def stop_injectomate_motion(self, **kwargs):
        self.stop()

    def wait_until_stopped(self, timeout_seconds=15):
        deadline=time.monotonic()+timeout_seconds
        while self.busy and time.monotonic()<deadline: time.sleep(.03)
        if self.busy: raise StereoDriveError("Stop has not completed; use physical Stop.")
        if self.drive and self.drive.is_moving(): raise StereoDriveError("Controller still reports movement.")

    def activate_drill_toggle(self):
        def toggle(d):
            if d.drill_state(): d.drill_off()
            else: d.drill_on()
        return self._run(toggle,asynchronous=True)

    def close(self):
        self.cancelled.set()
        self.supports_drilling_protocol=self.supports_injection_protocol=False
        if self.drive:
            try:self.drive.close()
            finally:self.drive=None
        if self.worker:self.worker.join(timeout=3)

    def syringe_to_limit(self, *, fill=False, stop_requested=None, asynchronous=False, on_completed=None):
        with self.lock:
            if self.busy:raise StereoDriveError('Controller busy')
            drive=self._require()
            endpoint=drive.travel_limits['PISTON'][1 if fill else 0]
            steps=drive.plan_piston_to(endpoint)
        return self._run(lambda d:self._execute_piston_steps(d,steps,stop_requested,on_completed),
                         stop_requested=stop_requested,asynchronous=asynchronous)

    def empty_syringe(self, **kwargs):
        return self.syringe_to_limit(fill=False,**kwargs)

    def fill_syringe(self, **kwargs):
        return self.syringe_to_limit(fill=True,**kwargs)

    def benchmark_axis_moves(self, axes=None, distances_mm=None, repeats=3, tolerance=.003,
                             stop_requested=None, progress_callback=None):
        axes=['AP','ML'] if axes is None else axes
        distances_mm=[.01,.02,.05] if distances_mm is None else distances_mm
        if (not axes or len(set(axes))!=len(axes) or any(a not in ('AP','ML','DV') for a in axes)
                or type(repeats) is not int or not 1<=repeats<=10 or not distances_mm or len(distances_mm)>20
                or any(isinstance(d,bool) or not isinstance(d,(int,float)) or not math.isfinite(d) or not 0<d<=1 for d in distances_mm)
                or isinstance(tolerance,bool) or not math.isfinite(tolerance) or tolerance<=0):
            raise StereoDriveError('Invalid benchmark: select axes, 1–10 repeats and distances >0, <=1 mm')
        start=self.get_current_axis_position();legs=[]
        for axis in axes:
            index=('AP','ML','DV').index(axis)
            for distance in distances_mm:
                target=list(start);target[index]+=distance
                for repeat in range(1,repeats+1):
                    legs.extend([(axis,distance,repeat,'+',tuple(target)),(axis,distance,repeat,'-',start)])
        self.validate_axis_path([leg[-1] for leg in legs])
        def run(drive):
            if drive.drill_state():raise ValueError('Turn the drill OFF before benchmarking')
            drive.validate_axis_path([dict(zip(('AP','ML','DV'),leg[-1])) for leg in legs]);rows=[]
            for i,(axis,distance,repeat,direction,target) in enumerate(legs):
                self._check_motion_cancelled(stop_requested)
                index=('AP','ML','DV').index(axis)
                before=drive.position()['axes_mm'][axis];started=time.monotonic()
                result=drive.move_axes_to(dict(zip(('AP','ML','DV'),target)))
                elapsed=time.monotonic()-started;end=result[axis];achieved=abs(end-before)
                if abs(end-target[index])>tolerance:raise RuntimeError('Benchmark endpoint verification failed')
                rows.append(dict(axis=axis,distance_mm=distance,direction=direction,repeat=repeat,
                    start=before,target=target[index],end=end,achieved_mm=achieved,elapsed_s=elapsed,
                    mm_per_s=achieved/elapsed if elapsed else None,error_mm=end-target[index]))
                if progress_callback:progress_callback(i+1,len(legs),axis,target[index])
            return rows
        return self._run(run,stop_requested=stop_requested)

    def scan_stereodrive_controls(self):
        return json.dumps(dict(backend="direct stereodrive_api",native_gui_control_ids="not used"),indent=2)

    def get_mmc_depth_gauge_handle(self): return None
    def get_mmc_depth_gauge_rect(self): return None
    def get_injectomate_calibrate_snapshot(self): return {}
    def confirm_below_skull_warning(self, **kwargs): return False
    def confirm_no_actual_movement_dialog(self, **kwargs): return False

    def wait_for_named_motion(self, *args, **kwargs):
        raise StereoDriveError("Native Home/Work is not available; use calibrated saved Axis targets.")
