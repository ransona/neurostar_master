# Timed pulsed workflows: supervised hardware bench test

This implements the original app's step-and-wait strategy using the shared USB API.
It does **not** introduce guessed firmware commands, continuous velocity control
or a continuous-flow pump. No physical hardware testing occurred during implementation.
Initial use must be on a supervised empty bench, no specimen, tool safely clear,
safe fluid collection and physical Stop immediately accessible.

## Enable it

1. Pull `codex/direct-api-control`, install PySide6, and first run simulation:
   `py -3 tools\craniotomy_qt.py`. USB setup has simulation-only calibration/history.
   Use the Python/source launcher: existing `dist` executables have not been rebuilt
   for this direct-control branch.
2. Hardware launch: `py -3 tools\craniotomy_qt.py --live`. Close native StereoDrive,
   probe server, standalone API GUI and other controllers. The app stays disconnected.
3. Load measured Axis-zero/piston-anchor calibration, verify current direction
   history for all channels, and independently verify physical reference/clearance.
4. Select DV, piston and/or drill permission as appropriate. Also select **Enable
   PULSED drilling/injection workflows**, then confirm verified setup and Connect.
   The preference is saved in mode-specific `direct-control.json`; connection and
   safety confirmation are never restored automatically.
5. Re-establish GUI Bregma and optional Anchor for the actual attached tool. Captured
   surfaces must be revalidated after changing tools/reconnecting. Every programmed
   Bregma target is converted to Axis using the current internal reference.

Run the two test commands in the root README before hardware tests. Start with
independently observed 0.01 mm manual AP moves, then permitted ML/DV movements and
10 nL free-piston collection tests. Motor counts alone do not prove displacement.

## Injection behavior

- Start and Resume from selected use the same complete-plan preflight and explicit
  bench confirmation. Sites must have verified surfaces, not merely map/grid XY.
- Approach via configured clearance, reach the surface, insert in steps no larger
  than 0.005 mm, reach the full overshoot endpoint, then retract to target depth.
- During that phase, deliver insertion free-piston pulses at the requested nominal
  average rate; then deliver the **additional** main volume in 10 nL pulses.
- Insertion volume is rounded **UP** to a multiple of 10 nL. It can materially exceed
  a tiny requested insertion volume: inspect the displayed total before starting.
- Hold, retract to surface using microsteps, return above surface and optionally
  perform blockage testing with the existing beep/prompt/retest flow.
- All commands are serial and verified idle. Axis motion and piston are not physically
  simultaneous. Pausing holds at the last verified position and freezes future
  events; Resume does not replay overdue doses. Stop cancels remaining events.
- Counters reflect verified piston-count displacement, not measured fluid delivery.
  The main volume, rounded insertion dose and two test volumes/site are reserved in
  preflight. Further retests recheck the entire test dose before starting.

A small initial *software-settings example*, not medical guidance: one validated
site at GUI Bregma AP/ML/DV zero; main volume 10 nL; insertion depth and overshoot
zero; hold zero; blockage testing off; nominal main rate 600 nL/min. This is one
10 nL pulse followed by return, not a steady 600 nL/min stream. Check all clearance
travel against the actual setup before enabling movement. Subsequently test a small
nonzero insertion depth on an independently measured fixture, not tissue.

## Drilling behavior

- All seed surfaces must be captured; a closed trajectory and target depth are required.
  The entire perimeter, depths, return route and possible pause retractions are
  preflighted before any target is sent.
- The app asks whether the drill is ON, asks for bench confirmation, and verifies
  reported USB drill power. It never turns the spindle on automatically.
- Depth entry uses up to 0.005 mm steps, perimeter following up to 0.01 mm steps.
  AP/ML/DV move sequentially and settle between commands: the path is a staircase,
  not a continuous circular cut. Requested rate/round duration can be exceeded.
- Frozen sections are crossed at clearance, never at cutting depth. A point's
  recorded drilled depth updates only after its target has been verified.
- Pause waits for the current pulse to finish, retracts to preflighted clearance
  and requests drill OFF. Turn it ON explicitly and verify spindle setup before
  Continue. Stop/fault/close attempt all-channel Stop and drill OFF.
- Successful completion returns above the center. Drill ON is left under operator
  control so existing automatic next-round/countdown behavior can work; turn it
  OFF explicitly after the agreed test. Never leave a spinning tool unattended.

Start with a tiny perimeter/depth fitting the verified bench envelope. Test without
contact/load first; use external displacement measurement and Stop observation.
Reported ON/OFF is not proof of spindle rotation/stationarity.

## Timing, limits and failure behavior

Each pulse uses the existing captured 1 or 2 mm/s axis profile or fixed free-piston
profile. The requested slow rate is an **average timing ceiling**, not instantaneous
speed/flow. A command can occupy multiple 200 ms settle periods. Overruns push later
events out; no pulse enlargement, parallel movement, early reversal or catch-up
burst is used to meet an unrealistic duration.

Pulsed workflow clearance uses **Options → Validation / pulsed clearance (mm)**
(default 0.5). It applies above surfaces and GUI Bregma. Choose a verified positive
clearance for the entire tool path; do not lower it just to fit software bounds.
Changing it affects future workflows, not a running plan's frozen configuration.

Existing limits are unchanged: ±1 mm per Axis from connection, ≤1 mm combined
distance per individual command; piston ±100 nL from connection and estimated
0–5000 nL capacity. Preflight starts from actual verified counts, so previous manual
piston movement changes available budget. Larger default clinical plans/grids can
be refused. Do not reconnect/reset/edit state to enlarge an envelope or evade faults.

Unexpected position/history, serial faults, timeout, external movement or uncertain
acknowledgement stop the workflow. There is no automatic recovery movement. Establish
physical condition and reference/history independently before reconnecting.

Record requested versus externally observed displacement, collected volume, timing,
pause behavior and Stop performance. Only supervised hardware testing can establish
real mechanical/fluid behavior; simulation results are not surgical validation.
