# Direct-control implementation review

Authoritative for `codex/direct-api-control`. Planner and HTTP server use the public
API exclusively, via `tools/direct_api_controller.py`. The historical Win32
controller remains for old diagnostics/regression tests, not as a runtime fallback.

## Coordinates and zero calibration

Calibration JSON contains `axis_zero_counts` (AP/ML/DV measured at mechanical zero),
`piston_3000_count`, and `anchor_backlash` (all four states when anchors were measured).
Current `verified_backlash` is separate; select it from independently observed history:

| Channel | Increasing / decreasing numerical coordinate | Backlash counts |
|---|---|---|
| AP | increased / decreased | 0 / 522 |
| ML | increased / decreased | 261 / 0 |
| DV | increased / decreased | 0 / 52 |
| Piston | native up / down | 5825 / 0 |

Never infer direction history from raw counts or copy simulation anchors. Axis sign
is AP −1, ML +1, DV −1; scale is 5225 counts/mm. The tested Nano 5 µL piston scale is
approximately 161.36 counts/nL. Calibration changes interpretation, not physical
position. No homing, hardware zero command or automatic calibration movement occurs.

`GUI Bregma = Axis − bregma_axis`. At Anchor uses
`bregma_axis = current Axis − stored anchor offset`; planned sites remain unchanged
in Bregma space. No StereoDrive Bregma fields/buttons are used. Each USB connection
conservatively clears tool references, captured surfaces and quick targets for
physical reverification. Home/Work persist only with identical Axis calibration;
changed/unknown calibration clears them. Simulation/live settings are separate.

Direct setup is stored in a dedicated `direct-control.json` in each mode's config
folder, not general `settings.json` or the API motion journal. It atomically saves
calibration, speed profile, enable preferences and calibrated Home/Work metadata.
Previously shared Home/Work/fingerprint fields migrate automatically. Restoring
preferences never connects, verifies physical setup, chooses current direction
history or enables New Reference automatically. Invalid/wrong-mode files are
ignored for control and retained until the user explicitly saves new setup.

## Reviewed physical entry points

| Entry point | Direct behavior / safety guard |
|---|---|
| Keyboard and validation-modal movement | Bounded `move_mm`; opposite-direction request cancels without immediate reversal |
| GoTo and named locations | Calibrated absolute targets; GUI Bregma converted at execution |
| Home/Work/A/B/Bregma | Same absolute API/progress/cancellation path; no native menu click |
| Map move and validation approach | Entire retract/XY/approach path preflighted before the first leg |
| Selected-site/grid validation | Direct keyboard fine adjustment; verified Axis capture converted back to Bregma |
| Position labels/map crosses | Verified idle counts or moving motor telemetry; stale displays cleared on read failure |
| Manual piston/test volume | Larger volumes expand into captured free steps; whole-dose preflight and verified per-step counters |
| Drill toggle | Worker-owned API ON/OFF; ON requires calibration/opt-in; Stop always attempts OFF |
| Injection start, Resume, sequence and worker | Explicit pulsed opt-in; complete travel/dose preflight, serial 10 nL delivery, no catch-up |
| Automatic drilling entry and worker | Pulsed opt-in and drill confirmation/reported ON; microsteps, frozen gaps, verified pause/retraction |
| Empty/fill | Confirmed captured-step sequence toward configured piston limit; cancellable progress, actual counter and <10 nL remainder if necessary |
| Axis benchmark | Selectable supervised out-and-back path; whole-path preflight, drill OFF, DV opt-in, cancellable progress and CSV |
| Local HTTP server | Same calibrated API adapter; mandatory measured setup, free-piston only |
| Branch Update / Restart | Verified direct branch/root, disconnect first, recoverable Git stash, pinned fetched commit, restart gate and preserved live flag |
| Native control-ID/screenshot diagnostics | Direct diagnostics/no native handles; not a source of direct counters |

New public API calls: `move_axis_to(axis, value_mm)`,
`move_axes_to({"AP": ..., "ML": ..., "DV": ...})`, `validate_axis_path(waypoints)`
and `live_position()`. Omitted axes hold position. Absolute requests reject unknown
axes, booleans, NaN/infinity, disabled DV changes and exceeded envelopes. Full-path
preflight sends no targets. Moves are sequential (retract DV before AP/ML where
requested; otherwise AP/ML/DV), **not collision planning or simultaneous trajectories**.

## Edge cases fixed during review

- Mandatory zero calibration before movement and drill ON; read-only diagnostics
  and Stop remain possible. Relative-reference bypass is simulator-test-only.
- Fractional normalized targets retained across absolute requests; count rounding
  does not accumulate from repeatedly reading rounded coordinates.
- Whole paths preflighted before moving; a bad final approach cannot cause a partial
  earlier move. All legs owned by one API call; Stop prevents subsequent legs.
- Manual piston/drill commands moved to workers to keep GUI Stop responsive.
  Concurrent moves rejected, not queued or retried.
- Stop/fault/close attempt all four channel Stops and drill OFF. Logging failure
  cannot suppress Stop/OFF packets. Exact target and idle settle for 200 ms.
- State flushed invalid before sending targets, valid after completion. Restore
  checks device identity, calibration, scales, raw counts and direction history.
  External changes/faults refuse motion; no automatic recovery or reversal.
- Disconnected simulation startup; live requires flag plus verified setup. Separate
  settings/state, one planner per mode, exclusive live transport, StereoDrive refused.
- Options ranges default to 0–40 mm per mechanical Axis and 0–5000 nL piston;
  connection-relative windows are removed. Estimated 0–5000 nL capacity remains;
  empty/fill stop within the configured range, not unlimited or controlled-rate delivery.
- Resume injection uses whole-plan preflight/bench confirmation; faults/disconnects clear stale
  crosses/readings. Branch updater only uses the verified direct branch; local edits
  are stashed, another branch/drifting checkout is refused, and restart is required
  before USB reconnects. Restart explicitly releases the old single-instance lock.
- Options probe waits for verified idle before reversal, not transient target counts;
  progress dialogs poll moving telemetry so map crosses stay current.
- Current history disagreement at restoration, boolean history/speed values and
  signed 32-bit count overflow are rejected before target writes.

## Pulsed workflows and remaining validation

Captured profiles cover axis moves at 1/2 mm/s and fixed free-piston steps. They do
not establish tissue-safe slow insertion/retraction, controlled nL/min injection,
continuous drilling, Auto-Speed, native skull limits or homing. The user-approved
pulsed implementation serializes small axis moves and 10 nL free-piston doses
through existing captured profiles. It never invents speed/flow packets. No overlap,
automatic recovery or overdue-dose catch-up is permitted. Actual duration can exceed
the requested schedule. See [bench guide](PULSED_BENCH_TEST.md).

Insertion volume is additional to main volume and rounded UP to 10 nL units.
Preflight checks all sites, depths, overshoot, return paths and two test volumes/site
against the actual remaining piston window/capacity. Every repeated blockage test
is checked again before its first dose. Piston counters come from verified API
results, not subtraction of assumed delivery. Larger manual/blockage test volumes
expand into captured steps, and repeated tests are preflighted as a complete dose.
Empty/fill use commanded fractional counts to plan without rounding across a bound;
any unachievable <10 nL remainder is reported. Benchmarks reject a bad later axis
before any target and stop on Cancel/fault without automatic reversal. Ordinary
Options/confirmation modals do not enable physical movement shortcuts.
Pause freezes future events; drilling
waits for idle, retracts via a preflighted path and requests drill OFF. Continue
requires explicit drill ON again. Frozen gaps are crossed at clearance.

Simulation covers calibration guards, absolute/fractional targets, path rejection,
multi-leg cancellation, Stop/OFF, API adapter routes, actual Qt setup/progress/
keyboard/piston/reference/disconnect behavior, and HTTP behavior. Run both root README
test commands. Original capture evidence is not end-to-end validation of this branch.

Simulation also exercises pulse quantization, insertion-plus-main accounting,
overdue timing, pause/cancel, frozen gaps, whole-plan rejection, end-to-end Qt
injection/Bregma/counters and drilling pause/completion. All test connections use
simulators, not hardware. Run both suites before bench testing.

The original pulsed-integration check passed 136 application tests and 22 API/GUI
tests. Subsequent regressions cover travel/speed Options, larger free-step doses,
empty/fill limits and counters, benchmark preflight/cancellation/CSV, and modal
navigation safety. Run the full suites for the current check. Added
regressions cover fractional/rounded held coordinates at envelope edges and an idle
display read racing command startup. The adapter serializes idle reads/preflight
with command startup; moving telemetry shares only the transport's I/O lock.

No actual hardware motion, drill activation or fluid delivery was performed during
integration. Reported counts can agree while physical steps are missed. Externally
validate both directions, reversals, tool changes, delivery and physical/software
Stop. Software bounds and passing simulations do not establish surgical safety.
