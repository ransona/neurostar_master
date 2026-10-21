# Agent guide: system design

This file is the repository entry point for coding agents. It explains the
current direct-USB architecture and the constraints that shape changes. Read it
before editing hardware-control code. For operator procedures and detailed
calibration behavior, follow the linked documents rather than inferring behavior
from UI labels.

## System at a glance

The current planner is a Qt application. It controls the device through a
Python API and a single-owner Windows USB-serial transport; it does not automate
the native StereoDrive GUI. The broad application layer manages plans and UI,
the adapter translates UI operations into API calls, the API validates and
coordinates operations, and the session owns serial I/O and persistent motion
state.

```text
Qt planner / calibration UI (tools/craniotomy_qt.py)
        │
        ├── direct planner adapter (tools/direct_api_controller.py)
        │       │
        │       └── public device API (stereodrive_api/api.py)
        │               │
        │               ├── Session and atomic motion journal
        │               │   (stereodrive_api/controller.py)
        │               ├── captured protocol definitions
        │               │   (stereodrive_api/protocol.py)
        │               └── one exclusive transport
        │                   (stereodrive_api/transport.py)
        └── planner/project/configuration state

Separate tools:
  tools/movement_probe/  local experimental USB probe server and its UI
  stereodrive_api/gui.py standalone API demonstration GUI
```

`tools/stereodrive_controller.py` is historical native-GUI automation retained
for diagnostics and regression coverage. It is not a fallback for the current
planner. The movement-probe server is a separate experimental tool, not part of
the planner's normal control path.

## Ownership and safety invariants

- **One serial owner.** A connected `StereoDrive` session owns the transport and
  serial I/O lock. Read through that session while connected; do not open a
  second COM port or call the legacy GUI automation alongside it. Close native
  StereoDrive and other device controllers before direct USB use.
- **No side effects at import.** Importing API/protocol modules must not open a
  port or move hardware. `StereoDrive.connect()` is explicit; commands require
  a connected session.
- **Simulation is not hardware evidence.** Keep simulation and live profiles
  and journals separate. Never copy synthetic positions, calibration, or
  direction history into a live profile.
- **Calibration is an interpretation of counts, not homing.** The saved zero
  anchors map raw counts to user-entered mechanical positions. Calibration
  does not move, home, or zero the physical mechanism. Count-derived position
  is not independent tool-tip feedback; piston nL is an estimate, not a
  measurement of delivered fluid.
- **Direction history is state.** Backlash state cannot be inferred safely
  from raw position alone. Maintain it through each move, persist it, and only
  initialize or replace it from independently verified direction history.
  Do not silently recover an invalid journal by creating a new reference.
- **Validate before motion.** Check input, configured travel limits, signed
  count bounds, and complete planned paths before issuing targets. Treat the
  calibrated scales, signs, backlash values, protocol captures, and default
  limits as evidence-backed constants; do not adjust them to make a test pass.
- **Persist uncertainty before sending a target.** `Session.move()` saves an
  invalid/in-progress journal before transmission. It marks state valid again
  only after stable completion or a verified stopped-position reconciliation.
  On an unknown result or transport/protocol fault, preserve the fault and
  require physical verification; do not retry a target or automatically
  reverse it.
- **Stop is a controller operation.** `StereoDrive.stop()` can be called from a
  separate thread while a blocking operation owns the session. The session
  handles stop commands and count reconciliation. Do not bypass this flow with
  a concurrent transport write. Physical Stop remains the independent fallback.
- **Do not send guessed packets.** Protocol builders and decoders live in
  `stereodrive_api/protocol.py`; transport owns packet exchange. Unsupported
  behavior must stay unavailable until captured and verified.
- **Hardware actions need explicit user authorization.** Routine code edits and
  simulation tests are fine. Never start physical movement, injection, drill
  power, calibration jogs, or probe commands as an incidental test. Follow the
  repository operating guide and ask for authorization if the user has not
  already authorized the specific physical action.

## Position, motion, and backlash model

The API uses channels `AP`, `ML`, `DV`, and `PISTON`. Axis coordinates are
mechanical millimeters; piston position is estimated nanoliters. The captured
axis scale is 5225 counts/mm with different signs by axis; the tested Nano 5 µL
piston scale is approximately 161.36 counts/nL. The API derives position from
raw counts, calibrated reference, sign/scale, and backlash offset. These are
count-based coordinates, not encoder-confirmed physical tool position.

A requested move changes a normalized count target. On reversal, the target
accounts for the channel's backlash state. Backlash differs by axis. During a
partial Stop, `Session.move()` reads stable idle counts, determines how much of
the backlash take-up and real travel occurred, and saves that reconciled state.
If the counts cannot be reconciled within the commanded path, the journal is
invalidated and further movement must be blocked pending verification.

Axis motion profile is configured separately from step size: supported API
profiles are 1 or 2 mm/s. GUI arrow steps are distances; repeated presses do
not alter the configured speed. The GUI serializes queued motion. Consecutive
pending steps in the same direction on one axis combine into a longer move;
steps separated by another axis remain in order. An opposite-direction press
removes stale pending moves on that axis and stops an active move on that axis,
then waits for Stop reconciliation and idle verification before continuing in
the new direction. Escape cancels the whole pending queue and stops the active
move. Preserve this distinction when changing queue behavior.

The manual piston-step queue is separate: pulses execute serially, and Stop or
Escape cancels pending pulses and reconciles the active step. Automated injection
and drilling are opt-in timed-pulse workflows, not continuous flow/speed
controls. Drill state is queried and confirmed through the API; it does not
prove spindle rotation or RPM.

## Persistence and startup

The GUI stores user preferences (including keyboard mappings) in
`%APPDATA%\\Neurostar_Master\\settings.json`; repo updates must not overwrite
that file. Direct-control configuration is separated by live/simulation mode
under `%APPDATA%\\Neurostar_Master\\DirectUSB\\...` and includes calibration,
speed/limits, and saved location metadata. The API has a separate motion-state
journal (`api-state.json`) and event log. Keep project/planner data, app
preferences, calibration configuration, and the API motion journal distinct.

Startup may attempt a live restore only when the saved calibration and journal
match the connected controller state. If verification fails, the planner uses
simulation and warns the user; it must not guess a live reference. Simulation
positions are volatile unless a specific persistent simulation mode is
requested. A GUI Bregma origin is an application coordinate transform layered
on mechanical Axis position; it is separate from USB calibration and native
device Bregma controls.

## Where to work

| Concern | Primary files |
| --- | --- |
| Main planner UI, coordinates, keyboard queues, persistence glue | `tools/craniotomy_qt.py` |
| Planner-facing adapter and worker/busy state | `tools/direct_api_controller.py` |
| Public API, connect/restore, limits, move orchestration, drill | `stereodrive_api/api.py` |
| Session, backlash transitions, interruption reconciliation, journal | `stereodrive_api/controller.py` |
| Packet definitions, status/ack decoders, captured constants | `stereodrive_api/protocol.py` |
| USB identity and exclusive serial transport; simulator | `stereodrive_api/transport.py` |
| Calibration anchors and validation | `stereodrive_api/calibration.py` |
| Application-facing operation and safety guide | `README.md` |
| API operation and design review | `stereodrive_api/README.md`, `docs/DIRECT_API_CONTROL.md` |
| Protocol evidence and limitations | `stereodrive_api/TECHNICAL_REPORT.md` |
| Separate experimental probe server | `tools/movement_probe/README.md`, `tools/movement_probe/AGENTS.md` |

## Development and verification

Prefer simulator-backed tests. Tests must not require hardware unless clearly
isolated and explicitly intended for a supervised bench. A focused test set is
usually faster and more diagnostic; run the relevant package suites after the
focused checks when a shared API or protocol change warrants it.

```powershell
py -3 -B -m unittest discover -s tests -v
py -3 -B -m unittest discover -s stereodrive_api/tests -v
py -3 -m py_compile tools/craniotomy_qt.py tools/direct_api_controller.py stereodrive_api/*.py
git diff --check
```

For a UI behavior change, add/adjust Qt integration tests in
`tests/test_direct_planner_qt.py`. For API/state/protocol behavior, use tests in
`stereodrive_api/tests/` and `tests/test_direct_api.py`. Test partial Stop and
state recovery paths when changing movement or persistence logic. A passing
simulator test verifies software logic only; it does not establish physical
displacement, injected volume, drill behavior, or safety.

## Read next

- [Operator setup and feature guide](README.md)
- [Direct API operation and integration contract](docs/DIRECT_API_CONTROL.md)
- [API package guide](stereodrive_api/README.md)
- [Protocol findings and test evidence](stereodrive_api/TECHNICAL_REPORT.md)
