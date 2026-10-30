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
then waits for Stop reconciliation and idle verification. The first
opposite-direction press is cancellation-only; a
second press is required to start motion in the new direction. Escape cancels
the whole pending queue and stops the active move. Preserve this distinction
when changing queue behavior.

The manual piston-step queue is separate: pulses execute serially, and Stop or
Escape cancels pending pulses and reconciles the active step. Automated injection
and drilling are opt-in timed-pulse workflows, not continuous flow/speed
controls. Drill state is queried and confirmed through the API; it does not
prove spindle rotation or RPM.

## Design logic behind the main workflows

### Calibration separates physical measurement from motion state

Calibration answers “what physical scale reading corresponds to these counts?”
It does not answer “which direction did the mechanism last move?” Those are
stored separately: `Calibration` holds the measured zero/anchor counts, while
the API journal holds current raw counts, normalized command counts, and
backlash state. Keeping these separate prevents a new scale reading from
silently resetting direction history or recovering a fault.

The calibration screen connects to the identified hardware (or an explicit
simulation) and reads the controller's own current position. If direction
history is not already established, it asks before making small positive moves
on AP/ML/DV and the piston. That known move gives the controller a defensible
starting direction state. Jogging then remains available so the operator can
align each axis with a readable physical scale mark. A physical-reading input
unlocks only after that axis has accumulated at least 0.200 mm of travel (50 nL
for the piston); this avoids saving a zero from a display value without a
meaningful physical comparison. Save is enabled only after all four physical
readings have been entered. The resulting zero anchors are derived from the
live raw counts, the supplied physical readings, and tracked backlash state.

When the device lacks a measured absolute calibration, the calibration jog
helper provides a bounded relative session for positioning. Its bounds and
journal are deliberately separate from the normal calibrated motion journal.
It must not be treated as an unrestricted movement mode or as a source of
absolute calibration by itself. When already connected, reuse that session and
its single port; never open another serial owner just to read counters.

### A move is a persisted state transition

`Session.move()` computes the requested normalized target before sending it.
It persists an invalid/in-progress state first because a process or cable can
fail after the motor receives a command but before software sees completion.
After a stable arrival it commits raw counts, backlash state, and normalized
position together. If Stop interrupts motion, it stops all motor channels,
samples until counts are stably idle, and reconciles partial backlash take-up
against actual travel. Only a count path consistent with the command can be
saved as valid. A timeout, uncommanded axis change, implausible count, protocol
error, or failed Stop leaves the session faulted; creating a new reference is
not a recovery shortcut.

The planner adapter has its own `busy` flag because UI workers may be running
while API calls are asynchronous. Display reads during motion use live telemetry
and do not claim verified idle. Before a new operation, the adapter waits until
both the worker and hardware motion flags are idle. Stop is routed to the
existing worker/session so only one thread owns serial I/O.

### Keyboard queue order is intentional

Keyboard moves are serialized because the API accepts one move at a time. The
queue stores axis, direction, label, and distance. Consecutive pending steps on
the same axis and in the same direction are summed into one longer command;
steps with an intervening axis move stay separate to preserve input order. A
direction reversal removes stale pending steps for that axis, stops its active
move if needed, waits for stopped-count reconciliation and idle verification,
then waits for the operator to press the reverse direction again. Other axes'
pending work stays in order. Escape has broader meaning: cancel the entire
pending queue and stop the active move.
Do not merge across a reversal, across another axis command, or across an
already-started motor command.

### Coordinate frames stay layered

The API and safety limits use mechanical Axis coordinates. GUI Bregma is a
planner transform: `Bregma = Axis − bregma_axis`. Named locations and site plans
are stored in their documented frame, then converted to Axis targets before
preflight and motion. Setting or recalibrating GUI Bregma must refresh both
planner views so maps and overlays use the same frame. A GUI coordinate change
must never rewrite API zero anchors or native device coordinates.

### Timed pulses are not continuous rate control

The injector and drill workflows sequence captured, bounded serial movements
with dwell/settling between them. Their configured rates describe a schedule;
they do not command firmware-controlled continuous speed or fluid flow. Whole
paths and doses should be preflighted before starting, each pulse should be
verified from counters, and cancellation should stop remaining work without an
automatic return/retry. Keep that distinction explicit in APIs, labels, and
operator docs.

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

## Planner workflow handover and non-obvious behaviour

This section records the application-level behaviour that is easy to miss when
following a single callback. The implementation is concentrated in
`tools/craniotomy_qt.py`; the adapter and API remain the only route to physical
commands. Treat the function names below as navigation hints, not as stable
public interfaces.

### Startup, state ownership and coordinate frames

- The planner has three related but distinct coordinate/state layers:
  mechanical Axis positions in the controller/API, the GUI's Bregma origin,
  and project geometry such as craniotomy seeds and injection sites. A Bregma
  change updates the planner transform and displayed map, not the API's USB
  calibration or raw controller counts.
- The anchor is a Bregma-relative location. Setting Relative Coordinates moves
  the Bregma origin to the operator-entered coordinates of the current physical
  position; it must preserve the anchor's Bregma-relative offset. The anchor is
  not the Bregma origin and must not be discarded when that origin changes.
- Startup offers to clear the saved Bregma location. Project/session restore,
  GUI settings, direct-control settings, and API movement journal are different
  stores. Do not conflate clearing one with clearing the others.
- Saved Go To locations are Axis positions persisted in direct-control setup.
  Home/Work and quick locations are subject to calibration compatibility and
  connection-time reverification. Plans and injection sites are Bregma-relative
  project data. Converting a plan to Axis targets should happen at the operation
  boundary, so the worker receives a fixed, immutable plan while it runs.
- The GUI displays moving telemetry while commands are active. During a move it
  is an estimate from controller counts, not independent physical tool-tip
  feedback. On read failure, stale position/map indications should be cleared,
  not left looking current.

### Dialogs, keyboard routing and workers

- Most long operations run on a Python worker thread. Qt widgets must only be
  changed on the GUI thread, using the existing signals, timers or queued
  callbacks. The adapter `busy` flag and the API session busy/moving state are
  separate. A worker being absent does not by itself prove that the controller
  is idle.
- The main `eventFilter` in `tools/craniotomy_qt.py` handles configured
  movement/syringe shortcuts. Ordinary options and confirmation dialogs
  suppress movement shortcuts. Some purpose-built workflows intentionally
  accept them: notably injection-site validation and setting craniotomy seed
  surfaces. A modal workflow that needs shortcuts must explicitly keep the
  event filter active and avoid the generic `_require_idle` guard blocking the
  intended movement.
- The seed-surface modal has a special drill-power shortcut. It toggles power
  ON or OFF while the modal is open. That command is synchronous by design:
  closing or advancing the dialog must not cancel an asynchronous drill worker
  before it has issued the toggle. Ignore key auto-repeat, and keep the dialog's
  displayed shortcut tied to the configured binding.
- Escape is context-sensitive. It cancels/stops ordinary queued movement, but
  in a craniotomy or single-hole run it is an emergency stop path, not a
  resumable pause. A deliberate Pause/Continue preserves progress; Stop/cancel
  has stronger semantics. Do not map Escape to merely closing a modal during
  active motion.
- Starting an operation disables and visually greys out unrelated controls.
  Stop/Pause and relevant progress remain available. Restore controls in every
  completion, error, cancellation and thread-start failure path.
- The two application logs are diagnostic records, not proof of physical
  action. GUI event logging records user actions, status/dialog text, errors and
  workflow milestones. API JSONL records commands, replies and state changes.
  A request entry is not a confirmation; look for a verified status/result.

### Craniotomy and single-hole drilling

- Craniotomy mode is either continuous perimeter traversal or spaced burr
  holes. In spaced mode, the generated perimeter points are seed/surface
  references; actual hole positions are generated using the configured spacing
  and the inferred/interpolated surface profile. Freeze state belongs to
  individual generated holes. Frozen holes are skipped while the remaining
  holes can be deepened.
- Surface-setting is a guided sequence over seeds. AP and ML are fixed by the
  selected seed; the operator adjusts DV to the surface, then chooses Set and
  Move to Next, Move to Next without setting, or Cancel. Returning to a prior
  seed is supported. Once surfaces are complete, the workflow returns above
  the craniotomy centre, prompts whether to start drilling, and offers to
  establish the desired drill power before continuing. The drill shortcut is
  intentionally available in this workflow.
- The surface prompt is not just a data-entry dialog. It is a controlled motion
  workflow, so retain its position updates, keyboard routing, cancel handling
  and final move to centre when editing its UI.
- A change to drilling geometry or relevant setup invalidates stored hole
  depths. The code resets software progress and requires explicit acknowledgement
  before restarting because physical partial holes cannot be undone. Editing a
  target depth is different: increasing the target can continue from the
  recorded depth. If the requested target exceeds Max Depth, the operator is
  asked whether to raise Max Depth as well.
- The Start/Pause control has several states: Start before the first round,
  Pause while active, Continue after pausing or between rounds, and it should
  remain Continue after a target is reached if a deeper target may be set. A
  completed round can prompt to increase depth or stop. Accepting a higher
  target should launch the next round, retaining per-hole depth. Auto-start
  between rounds uses a countdown and configurable delay.
- In spaced mode, rate means the maximum DV advance rate and there is no
  requested time-per-circuit pacing. Each pulse is scheduled against elapsed
  time so delayed worker activity catches up to the depth due at the current
  rate. Pausing retracts to clearance, but the stored hole depth is retained;
  Continue rapidly returns to that depth and resumes advancing. This distinction
  prevents re-drilling a partially completed hole from its surface.
- Continuous-mode round timing and spaced-mode rate are intentionally different
  models. Do not share a “round duration” wait loop with spaced mode. Keep the
  status/map current as the worker moves between points, report elapsed time,
  completed holes/round fraction and target depth, and omit the removed
  unreliable remaining-time estimate.
- At target completion, the drill is commanded OFF and verified. Failure to
  confirm OFF is an actionable warning. When extending the target, the workflow
  can turn it back on only through the normal explicit confirmation/verification
  path. `start_drilling_round`, `_run_drilling_round`,
  `_run_pulsed_drilling`, `on_drill_round_finished`, and
  `_prompt_increase_drilling_depth_or_stop` form the central path.
- The single-hole tab drills relative to the current AP/ML/DV position. Its
  target depth defaults to 0.2 mm; its rate is separately editable and saved,
  initially inherited from the craniotomy rate. `delayAtDepth` is configurable
  and defaults to 3 seconds. Pause retracts to the configured clearance above
  the start point, while retaining reached depth; Continue returns to that depth
  before advancing. Cancel stops the run. Escape emergency-stops and retracts
  to the clearance point. Target extension should continue from already-drilled
  depth, not repeat the initial descent.
- Depth colour scaling must cover the larger of skull thickness and target
  depth. Geometry/progress signatures, restoration and invalidation are handled
  around `_borehole_pattern`, `_ensure_borehole_state`,
  `_invalidate_drilling_progress`, and `_restore_borehole_progress`.

### Injection workflow

- Injection sites are stored in Bregma coordinates. The sequence converts all
  selected sites to Axis coordinates once at start and preflights the full
  approach, insertion, dose, retraction, return and optional test-volume path
  before issuing the first target. Keep the complete-plan preflight; do not
  validate each leg only as it is reached.
- “Start From Selected” begins at the selected site, including earlier
  unvalidated sites only when the relevant validation flow has been invoked.
  If Go finds unvalidated sites, the operator can enter site validation; after
  validation completes, the original request should proceed into injection,
  not return to idle and require a second Go. Preserve the start index and
  original intent through that modal transition.
- Site validation is deliberately interactive: keyboard nudges refine AP/ML
  and DV, then Validate and Next records the site and advances. Next Without
  Validating advances without marking valid; Delete Point removes the site;
  Cancel exits. A single-site validation may be opened before starting the
  sequence. Avoid changing the selected-site indexing while the modal is open.
- During the sequence, manual control, hand injection and recipe/settings
  controls are disabled and greyed out. The operator must retain access to
  sequence progress, current movement progress, Pause/Resume and Stop. Status
  should state site N of total and the current operation. Insertion progress
  includes depth in micrometres and percent; injection progress includes
  movement volume and percent. Do not call it a dose.
- The pulsed injector uses serial free-piston movements, usually 10 nL units,
  and small axis movements. It is neither continuous flow nor a guaranteed
  delivery rate. Piston position is estimated from counts. Manual syringe
  moves are queued and bounded to 500..4500 nL; Empty moves to the lower limit,
  Fill to the upper limit, and Go To sets a requested piston value. The blockage
  test volume defaults to 20 nL and is an option. Every operation must preflight
  the available piston range.
- Optional insertion confirmation asks whether the pipette is visibly inserted
  at the planned depth. Yes continues. The operator may request another 50 um
  overshoot, repeated as needed, or cancel, which pauses/stops the sequence.
  Any additional overshoot is reversed before the planned injection/retraction
  path continues. Keep test-volume and injection accounting consistent when
  repeating a site.
- If a blockage test reports blocked, the operator can rerun it. If a later
  blockage test reports clear after one or more tests, the sequence asks whether
  to return to and repeat the last site. Yes must actually loop that same site
  before advancing; this was a prior regression. Re-preflight remaining syringe
  capacity before each repeat.
- `pause_resume_injection` changes the button to Resume while paused. Stop calls
  the controller Stop path and cancels pending syringe pulses. Do not use a
  modal warning for routine Stop. Progress should be monotonic across sites;
  completing a site must not report 100 percent unless it is the final site.
  Relevant navigation points are `_offer_validation_before_injection`,
  `_run_injection_site_validation`, `_start_injection_sequence`,
  `_run_pulsed_injections`, `_run_block_test`, and `stop_injection`.

### Drill power, recovery and faults

- “Drill ON/OFF” is a reported power-state request. ON is polled until confirmed
  or a configurable per-attempt timeout expires; retry count is configurable in
  Options (current defaults: 3 seconds per attempt and 4 retries after the first
  attempt). OFF is not assumed to
  have succeeded just because a command was written. The reply decoder's
  reported state is authoritative for confirmation, though it still cannot
  prove spindle rotation.
- Drill-only failure can sometimes be recovered without discarding movement
  calibration. `Session` narrowly recognises a drill-only invalidation, verifies
  controller identity and unchanged motor counts, then re-establishes the
  session only if drill OFF is confirmed. It must not recover if any AP/ML/DV or
  piston count, calibration or identity differs. The GUI Recover button warns
  first and runs this guarded procedure. Recovery is not a generic “clear fault”.
- Axis/motion faults remain latched. Do not automatically retry an axis target,
  reverse motion, create a new reference, or treat cached requested state as
  confirmed physical state. A failed Stop/OFF write is separately logged and
  physical Stop may be required.
- Craniotomy surface setting asks whether the drill should be on at workflow
  start. The answer is applied and state-confirmed before motion. “No” means
  ensure it is off, including turning it off if currently on. During the modal,
  the configured drill shortcut toggles it; this is an explicit exception to
  normal modal keyboard suppression.

### Configuration, logs and practical development

- General settings, project state, direct-control config, craniotomy/injection
  named config files, and API state are separate. New persistent options need
  both a default/widget and load/save logic. Preserve backward compatibility:
  older JSON files will lack newer keys, so loading needs defaults and type/range
  checks. Do not save transient progress as if it were durable physical truth.
- The Windows app data root is `%APPDATA%\Neurostar_Master`. General options
  are in `settings.json`; direct USB files are under `DirectUSB\live` or
  `DirectUSB\simulation`, including `direct-control.json`, `api-state.json`
  and API JSONL events. The application event log is `app-events.jsonl`.
  Simulation and live paths must remain separate. On macOS/Linux the helper
  resolves an OS-appropriate user config root; inspect `_user_app_data_root`
  rather than hard-coding Windows paths into code or docs.
- Event logs use UTC timestamps. GUI log events include message text and
  responses, action milestones and exception details; API logs include packet
  exchange and command/result events. Logs help reconstruct chronology but
  cannot prove a motor or drill physically moved. Do not log secrets or add
  unbounded high-frequency telemetry that makes the log unusable.
- For edits, first locate the owning state transition, then its worker and GUI
  completion callback. Search current tests before assuming a UI label captures
  semantics. Add Qt regression coverage in `tests/test_direct_planner_qt.py`,
  API/session coverage under `stereodrive_api/tests/` or `tests/test_direct_api.py`,
  and protocol tests when packets or decoders change. Common focused commands:

  ```powershell
  py -3 -B -m unittest tests.test_direct_planner_qt -v
  py -3 -B -m unittest discover -s stereodrive_api/tests -v
  py -3 -B -m unittest discover -s tests -v
  py -3 -m py_compile tools/craniotomy_qt.py tools/direct_api_controller.py stereodrive_api/*.py
  git diff --check
  ```

  GUI tests can require PySide6 and platform display support. Use the existing
  headless Qt setup in the test module. Never use real hardware to validate an
  incidental code change.
