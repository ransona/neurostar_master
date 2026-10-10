"""Deterministic bench pulse plans; no device access or GUI at import.

Rates are requested average ceilings, not instantaneous motor speed/flow.
All events are serial, verified, cancellable, and never replayed to catch up.
"""
from dataclasses import dataclass, replace
import math
import time


class PulseCancelled(RuntimeError):pass
class PulsePaused(RuntimeError):pass


@dataclass(frozen=True)
class Pulse:
    due_s: float
    phase: str
    target: tuple | None = None
    volume_nl: int = 0
    point_index: int | None = None
    depth_mm: float | None = None
    surface_dv: float | None = None


def finite(value, name, minimum=0):
    if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value) or value<minimum:
        raise ValueError(f"Invalid {name}")
    return float(value)


def line(start, end, step=.005):
    step=finite(step,'pulse step',1/5225)
    start=tuple(finite(v,'Axis coordinate',-math.inf) for v in start)
    end=tuple(finite(v,'Axis coordinate',-math.inf) for v in end)
    if len(start)!=3 or len(end)!=3:raise ValueError('Three Axis coordinates required')
    count=max(1,math.ceil(math.dist(start,end)/min(step,1)))
    if count>20000:raise ValueError('Pulse path too large; review travel limits')
    return [tuple(a+(b-a)*i/count for a,b in zip(start,end)) for i in range(1,count+1)]


def clearance_path(start,end,clearance_dv):
    safe=min(start[2],end[2],clearance_dv)
    targets=((start[0],start[1],safe),(end[0],end[1],safe),end)
    points=[]
    for target in targets:
        points.extend(line(start,target,1));start=target
    return points


def insertion_volume(settings):
    seconds=(settings.injection_depth_mm+2*settings.overshoot_mm)/(settings.insert_retract_speed_um_s/1000)
    # Round upwards to the captured minimum 10 nL free-step unit; report it.
    return 10*math.ceil(max(0,seconds*settings.insertion_rate_nl_min/60)/10-1e-10)


def injection_plan(start,site,settings,clearance_dv,clearance_mm=.5):
    for name in ('injection_depth_mm','overshoot_mm','post_inject_pause_s'):
        finite(getattr(settings,name),name)
    for name in ('insert_retract_speed_um_s','insertion_rate_nl_min','main_rate_nl_min'):
        finite(getattr(settings,name),name,1e-9)
    volume=finite(settings.main_volume_nl,'main volume',10)
    if volume%10:raise ValueError('Main volume must be a multiple of 10 nL')
    clearance_mm=finite(clearance_mm,'clearance',1/5225)
    surface=(site.ap,site.ml,site.dv)
    above=(site.ap,site.ml,site.dv-clearance_mm)
    events=[Pulse(0,'approach',target=p) for p in clearance_path(start,above,clearance_dv)]
    events += [Pulse(0,'surface',target=p) for p in line(above,surface,1)]
    speed=settings.insert_retract_speed_um_s/1000
    overshoot=(site.ap,site.ml,site.dv+settings.injection_depth_mm+settings.overshoot_mm)
    depth=(site.ap,site.ml,site.dv+settings.injection_depth_mm)
    movement=[]; elapsed=0; current=surface
    for phase,target in (('insert',overshoot),('overshoot_retract',depth)):
        for p in line(current,target):
            elapsed+=math.dist(current,p)/speed
            movement.append(Pulse(elapsed,phase,target=p));current=p
    insert_volume=insertion_volume(settings)
    doses=[Pulse(i*10/settings.insertion_rate_nl_min*60,'insertion_dose',volume_nl=10)
           for i in range(1,insert_volume//10+1)]
    phase_end=max(elapsed,doses[-1].due_s if doses else 0)
    events+=sorted(movement+doses,key=lambda e:e.due_s) # Stable: axis before dose on ties.
    for i in range(1,int(volume)//10+1):
        events.append(Pulse(phase_end+i*10/settings.main_rate_nl_min*60,'main_dose',volume_nl=10))
    elapsed=events[-1].due_s+settings.post_inject_pause_s
    if settings.post_inject_pause_s>0:events.append(Pulse(elapsed,'hold'))
    current=depth
    for p in line(current,surface):
        elapsed+=math.dist(current,p)/speed
        events.append(Pulse(elapsed,'surface_retract',target=p));current=p
    events += [Pulse(elapsed,'return_above',target=p) for p in line(surface,above,1)]
    return events


def drilling_plan(start,surfaces,current_depths,target_depths,frozen,clearance_dv,clearance_mm,
                  center_above,round_seconds,depth_rate):
    finite(clearance_mm,'clearance',1/5225);finite(round_seconds,'round time',1e-9);finite(depth_rate,'depth rate',1e-9)
    if len(surfaces)<2 or any(len(v)!=len(surfaces) for v in (current_depths,target_depths,frozen)):
        raise ValueError('Drilling arrays must match the closed perimeter')
    if math.dist(surfaces[0],surfaces[-1])>1e-6:raise ValueError('Perimeter must be closed')
    count=len(surfaces)-1
    needed=[not frozen[i] and current_depths[i]+.0005<target_depths[i] for i in range(count)]
    if not any(needed):return []
    first=needed.index(True); order=[(first+i)%count for i in range(count+1)]
    events=[];elapsed=0;current=start;previous=None
    for index in order:
        ap,ml,surface=surfaces[index]
        if frozen[index]:previous=None;continue
        target=(ap,ml,surface+finite(target_depths[index],'depth'))
        if previous is None:
            at_depth=(ap,ml,surface+finite(current_depths[index],'current depth'))
            safe=min(clearance_dv,surface-clearance_mm)
            events += [Pulse(elapsed,'approach',target=p,surface_dv=surface) for p in clearance_path(current,at_depth,safe)]
            current=at_depth
            for p in line(current,target):
                elapsed+=math.dist(current,p)/depth_rate
                events.append(Pulse(elapsed,'enter',target=p,surface_dv=surface));current=p
        else:
            points=line(current,target,.01)
            for p in points:
                events.append(Pulse(elapsed,'cut',target=p,surface_dv=p[2]-target_depths[index]));current=p
        events.append(Pulse(elapsed,'point_complete',point_index=index,depth_mm=target_depths[index],surface_dv=surface))
        previous=index
    cuts=sum(e.phase=='cut' for e in events)
    # Spread cutting pulses across requested round time; completion can make it longer.
    added=0
    for i,e in enumerate(events):
        if e.phase=='cut':added+=round_seconds/max(1,cuts)
        events[i]=replace(e,due_s=e.due_s+added)
    elapsed=events[-1].due_s if events else 0
    events += [Pulse(elapsed,'return_center',target=p) for p in clearance_path(current,center_above,clearance_dv)]
    return events


def execute(controller,events,*,stop_requested=lambda:False,pause_requested=lambda:False,
            retract_on_pause=False,on_event=lambda e:None,on_delivered=lambda value:None,
            clock=time.monotonic,sleep=time.sleep):
    base=clock();shift=0
    for event in events:
        while True:
            if stop_requested():raise PulseCancelled('Pulsed workflow cancelled')
            if pause_requested():
                if retract_on_pause:raise PulsePaused('Pause requested at verified idle')
                paused_at=clock()
                while pause_requested():
                    if stop_requested():raise PulseCancelled('Pulsed workflow cancelled while paused')
                    sleep(.03)
                shift+=clock()-paused_at
            remaining=base+event.due_s+shift-clock()
            if remaining<=0:break
            sleep(min(.03,remaining))
        # Lost wall time is discarded rather than recovered by overdue pulses.
        shift+=max(0,clock()-(base+event.due_s+shift))
        if event.target is not None:
            controller.goto_axis_position(*event.target,stop_requested=stop_requested)
        if event.volume_nl:
            result=controller.syringe_step(f'{event.volume_nl} nl',up=False,stop_requested=stop_requested)
            on_delivered(result['PISTON'])
        on_event(event)
        if stop_requested():raise PulseCancelled('Pulsed workflow cancelled')
