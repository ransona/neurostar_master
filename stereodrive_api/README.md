# StereoDrive Python USB API

Importable, standard-library-only Python 3.10+ package for the tested Windows
StereoDrive / Nano 5 µL bench configuration. It talks directly to the identified
USB virtual serial controller; it does not need the movement-probe server.
Importing the package does not connect or move anything. Simulation is the default.
On `codex/direct-api-control`, the planner and movement server use this API too.
See the [branch instructions](../README.md) and [integration review](../docs/DIRECT_API_CONTROL.md).
Movement and drill ON require measured absolute zero calibration by default.
Read-only diagnostics and Stop remain available without it; bypass is forbidden live.

## Safety and what position means

**This controller reports motor counts, not independently measured tool position.**
A motor may miss physical steps while still reporting its requested count. Arrival,
persistent state and backlash compensation cannot detect that failure. Never use
reported arrival alone as proof of physical clearance, delivered volume or tissue
position. This experimental API is for supervised bench validation, not unattended
surgery.

The axis profile is fixed to the captured **2 mm/s** request (field 1306). Pass `speed_mm_s=1`
for the captured 1 mm/s profile (field 653); higher profiles are rejected. This is a conservative starting setting, not a guarantee
against missed steps. Validate under your actual load with an external displacement
measurement, in both directions and after reversals. Start at 0.01 mm. If counts
and physical motion disagree, use physical Stop, end the session, inspect the
mechanics and re-establish a measured reference; do not compensate by blindly
adding counts. Power loss, stalls, manual repositioning and movement by other
software also invalidate physical position assumptions even when raw counts match.

Keep the bench clear, drill off and physical Stop accessible. Verify clearance
throughout the entire allowed envelope. DV and piston movement require explicit
constructor opt-in. No collision planning, homing, skull zones,
controlled injection-rate programming, calibration or automatic reversal is provided.
Software Stop is best effort and requires a functioning connection and Python process.

For piston tests verify syringe type/calibration, plunger travel and safe fluid
collection with the pipette out of tissue. Up/down are native labels, not promises
of aspiration or dispensing direction. Reversing may aspirate air/contaminants and
does not undo delivered fluid. Piston calibration below applies only to the tested
Nano 5 µL configuration and is approximate; validate volume independently.

## Import and try simulation

Run from the repository root, or put that root on your program's Python import path:

```python
from stereodrive_api import StereoDrive, Calibration

# These initial states are for the simulator only; never assume them for hardware.
initial = {"AP": 0, "ML": 261, "DV": 0, "PISTON": 0}
# Synthetic simulator counts only, never a live calibration:
calibration = Calibration({"AP": 105280, "ML": 75864, "DV": 41767}, -8572, initial)
drive = StereoDrive(simulate=True, calibration=calibration,
                    allow_dv=True, allow_piston=True)
drive.connect(verified_backlash=initial, new_reference=True)
with drive:
    print(drive.position())
    drive.move_mm("AP", +0.01)  # blocks until independently reported idle at raw target
    drive.move_mm("AP", -0.01)  # explicit, after successful completion only
    drive.piston_step("up", 10)
    drive.piston_step("down", 10)
```

To test package logic without hardware:

```powershell
py -3 -B -m unittest discover -s stereodrive_api/tests -v
```

## Connect to hardware

1. Validate the actual bench setup and physical reference. Stop and close
   StereoDrive, movement-probe, the standalone USB GUI and other automation.
   The API rejects connection while StereoDrive.exe runs and opens the port
   exclusively. Do not run two controllers.
2. Record the verified last physical/numerical movement direction for each axis
   and piston before closing the native app. Raw counts do not reveal backlash
   state. If history is unknown, establish it through an independently observed,
   safe native procedure first; do not guess or toggle state to fit desired readings.
3. Load independently measured `Calibration` anchors (see below), then construct
   `StereoDrive(simulate=False, calibration=calibration, allow_dv=False, allow_piston=False)`.
   DV/piston can be explicitly enabled only after checking their clearance/setup.
4. On the first connection pass `verified_backlash` with all four entries from
   the table below. Coordinates are interpreted from the measured anchors;
   connection does not reset the device, Bregma or StereoDrive coordinates.
5. On subsequent connections call `connect()` with no reference reset. Restoration
   requires valid saved state, identical device identity/calibration and exactly
   matching idle raw counts. Verify the physical reference too. A mismatch refuses
   connection. Never use `new_reference=True` merely to bypass that refusal.

| Channel | Verified most recent direction | Backlash state to pass |
|---|---|---:|
| AP | mechanical AP increased / decreased | 0 / 522 |
| ML | mechanical ML increased / decreased | 261 / 0 |
| DV | mechanical DV increased / decreased | 0 / 52 |
| PISTON | native up / down | 5825 / 0 |

The Windows registry resolves **VID 0483, PID 5743, serial 206334AC5031** to its
current COM port. No arbitrary COM fallback, driver installation, baud guessing
or DTR/RTS changes. A different controller requires separate identification and
validation before changing the transport identity.

## Public methods and limits

| Method | Behavior |
|---|---|
| `connect(verified_backlash=None, new_reference=False)` | Restore matching saved state, or establish verified calibrated reference; supplied current history must match saved state |
| `position()` | Fresh idle verification; returns `axes_mm`, `piston_nl_estimate`, `raw_counts`, `backlash_counts`, `simulated` |
| `move_mm("AP"/"ML"/"DV", signed_delta)` | Blocking single-axis relative move, max 1 mm, axes default captured 2 mm/s profile |
| `move_axis_to(axis, position_mm)` | Blocking calibrated absolute Axis target |
| `move_axes_to(targets)` | Preflight AP/ML/DV mapping, then sequential moves; omitted axes hold; combined distance <=1 mm |
| `validate_axis_path(waypoints)` | Validate complete absolute path/envelope before sending any target |
| `validate_piston_steps(signed_steps)` | Preflight signed 10/20/50/100 nL doses, connection/capacity/raw limits; no target writes |
| `live_position()` | Motor-derived moving telemetry with `verified_idle`, not encoder feedback |
| `piston_step("up"/"down", volume_nl)` | Blocking free-piston step; supported 10,20,50,100 nL; captured fixed profile |
| `stop()` | Request cancellation from another thread, then send Stop to all four channels |
| `close()` / context-manager exit | Stop, preserve valid idle state or retain invalid state, release port |

All movements must remain within **±1 mm per axis and ±100 nL estimated piston
position from the connection position**, including cumulative moves. Calibrated piston
targets must also stay within the tested Nano 5 µL estimated 0–5000 nL capacity. The envelope
is a software limit, not collision protection. A reconnect establishes a new
connection envelope; do not reconnect to expand travel without reviewing clearance.
Absolute multi-axis moves are sequential: requested DV retraction first, then AP/ML;
otherwise AP/ML/DV. Stop cancels remaining legs. There is no collision planning or
combined out-and-back command. A new move is rejected
while another call owns the controller. Movement methods are blocking: use a worker
thread in GUIs and call `stop()` from another thread. Stop waits for the current
bounded serial exchange to finish; it is not instantaneous. Poll/reply timeouts
are about 0.6 seconds per exchange (a four-channel snapshot can take longer); each move has a 10-second arrival timeout. Do not terminate
Python instead of stopping the hardware: a controller may continue its last target.

With no calibration, positions are relative to the verified saved reference.
With a Calibration, positions are on its absolute scales. They are not automatically
native Axis or Bregma values; that depends on the measured anchors supplied. Piston nL is a motor-derived estimate, not measured fluid volume.
Axes use the observed approximate 5225 counts/mm; piston uses 161.36 counts/nL.
Constructor calibration values can be overridden after independent validation;
calibration changes make existing saved state incompatible. Backlash values are
bench-specific observations (AP 522, ML 261, DV 52, piston 5825 counts), not universal
specifications. Fractional targets are retained to reduce rounding accumulation.

The packet profile stays constant across movement sizes. Effective speed changes
with acceleration and braking; the profile's other fields have not been fully
identified. Auto-Speed and native safety zones are **not implemented**. Piston uses
field 1820 and its captured free-step flags; do not equate this with axis mm/s or
a specified injection flow rate.

## Persistence, faults and logs

By default files are stored in `%LOCALAPPDATA%\StereoDrivePythonAPI`:
`live.json` or `simulation.json`, with a corresponding UTC `.jsonl` traffic/event log.
Simulation also stores its mock motor counts separately. A custom `state_path` is
supported. Use a separate path for simulation and hardware and only one process per
state path. Never edit state manually to recover motion or delete it to evade a fault.

State is atomically saved and flushed **invalid before sending a target**, then
saved valid only after matching raw target and idle readings settle for 200 ms.
Saved state includes reference, fractional targets, raw counts and backlash state.
An interruption, timeout, unexpected reply, external movement or write failure
attempts Stop and invalidates state. No automatic target retry or recovery reversal.
Resolve the cause and independently verify physical position and direction history
before starting a new reference. `connect()` does not clear a physical fault.

Testing validates packet construction, limits, persistence and cancellation in a
simulator. It cannot establish mechanical accuracy. This package has not yet been
bench-tested directly as a complete API; its protocol is based on captured native
movements and the earlier standalone controller implementation.

## Absolute calibration anchors

```python
from stereodrive_api import StereoDrive, Calibration

# Example numbers only: replace with measured anchors for YOUR apparatus.
calibration = Calibration(
    axis_zero_counts={"AP": 100000, "ML": 70000, "DV": 40000},
    piston_3000_count=-6000,
    anchor_backlash={"AP": 0, "ML": 0, "DV": 0, "PISTON": 0},
)
drive = StereoDrive(simulate=True, calibration=calibration, speed_mm_s=2)
# First connection still requires verified CURRENT backlash history.
# Anchor backlash is history AT THE MEASURED ANCHOR, which may differ.
drive.connect(verified_backlash={"AP": 0, "ML": 261, "DV": 0, "PISTON": 0})
with drive:
    print(drive.position())
    print(drive.is_moving())
```

Axis anchors are the raw count observed at 0 mm on each desired scale. The piston
anchor is the raw count at 3000 nL on its scale. `anchor_backlash` defaults to zero
for all channels: this means anchors were measured in the zero-backlash direction,
or already normalized by subtracting backlash. If they are uncorrected raw counts,
pass their actual verified anchor states using the direction table above.

The calculation is `position = (raw - current_backlash - reference) / (sign * scale)`.
Axis reference is `zero_count - anchor_backlash`; piston reference is
`count_at_3000 - anchor_backlash - 3000 * 161.36`. Axis signs are AP -1, ML +1, DV -1;
piston native up is positive. Calibration changes coordinate interpretation only;
it sends no zeroing or movement command. Reuse the same calibration when restoring
state. Changed anchors are rejected unless an independently verified new reference
is explicitly established. Travel envelopes stay relative to connection position.

## Moving and drill state

`moving_status()` returns `{ "any_moving": bool, "channels": {"AP": bool,
"ML": bool, "DV": bool, "PISTON": bool} }`. `is_moving()` returns any-channel motion;
`is_moving("DV")` queries one channel. These fresh controller queries can run from
another thread during a blocking movement; they share a serial lock with movement
polling. They do not prove physical motion or absence of missed steps.

`drill_state()` queries reported power on/off. To enable `drill_on()`, explicitly
construct with `allow_drill=True` after checking the drill setup. `drill_off()` is
available without enabling ON. Both verify the reported power transition for up
to three seconds; state can lag the write by approximately half a second.
Stop/close and motion faults always attempt drill OFF, even without ON permission;
Stop/close verify the OFF state. Software OFF does not prove the spindle has stopped
rotating. Do not touch the drill until physically stationary. This interface does
not set drill RPM. Drill commands are rejected while a movement owns the controller;
use `stop()` first. ON is never automatically retried after uncertainty.

The drill bytes come from one operator-controlled ON/OFF capture. The new direct
USB drill methods are tested in simulation only, and require supervised hardware
validation before practical use. No drill ON command was sent during development.

## Investigation evidence

See [technical report](TECHNICAL_REPORT.md), [capture archive](data/README.md),
and the figures in `figures/`. The archive includes a metadata JSON and readable
metadata document for every scoped capture, including failed and passive trials.

## GUI using this same API

Launch `pyw -3 -m stereodrive_api.gui` from the repository root, or use the existing
StereoDrive USB Controller desktop shortcut. The separate GUI folder now only
launches this shared implementation. It starts disconnected in Simulation.
It requires loaded measured zero calibration before connection/movement. It provides
AP/ML/DV arrows, mm step sizes, default 2 mm/s (optional 1), bounded
10/20/50/100 nL piston up/down steps, drill ON/OFF and reported moving/power status.
DV, injector and drill ON must be explicitly enabled after setup verification.

Load an absolute `Calibration` JSON with keys `axis_zero_counts`,
`piston_3000_count`, and `anchor_backlash`; example values above must be replaced
with actual measured anchors. GUI calibration and speed persist per mode. GUI
state/logs are in `%LOCALAPPDATA%\StereoDriveUSBController` using new `api-state-*.json`
files; old standalone three-axis state is not silently reused.
