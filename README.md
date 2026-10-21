# Neurostar planner — direct USB branch

Branch: **`codex/direct-api-control`**. The Qt planner and local movement server
now use [`stereodrive_api`](stereodrive_api/README.md), not StereoDrive text boxes,
Bregma fields or button clicks. **StereoDrive must be closed.**

This is an experimental, supervised bench branch. Motor counts are not independent
tool-position feedback. No hardware motion was performed during implementation.
Automatic drilling and injection sequences are implemented using **timed pulses**
behind a USB-setup opt-in. These are verified serial microsteps/free-piston doses,
not continuous-speed cuts or firmware-controlled injection flow. See the
[pulsed bench-test guide](docs/PULSED_BENCH_TEST.md) before hardware testing.

## Install and launch

Use Windows and Python 3.10+. From the repository root:

```powershell
git fetch origin
git switch codex/direct-api-control
git pull --ff-only
py -3 -m pip install PySide6
py -3 tools\craniotomy_qt.py
```

On normal startup, the planner tries the saved live calibration and matching
API state first. It never moves or creates a new hardware reference during this
attempt. If USB is unavailable or the saved state does not match, it starts the
in-memory simulator and displays a warning. Use `--simulation` to skip USB:

```powershell
py -3 tools\craniotomy_qt.py --simulation
```

`--live` remains accepted for compatibility; automatic startup already tries
hardware. The calibration screen's **Save Calibration** button stores measured anchors in
the mode-specific `direct-control.json`, which is loaded at the next start.
`run_craniotomy_gui.bat` launches from its own folder and accepts either flag.
Update does not reset this branch to `main`; close the app and use `git pull --ff-only`.

## Before enabling hardware

1. Clear the bench: no specimen, tool/pipette retracted, drill off, physical Stop
   accessible. Close StereoDrive, all other controllers and the probe server.
2. Open **USB setup / zero calibration**. The single calibration window closes
   any simulator connection and connects directly to physical USB, showing the
   connection state, current position/counts and last-motion direction for each
   axis. The window asks before making small positive-direction initialization
   moves (+0.100 mm on AP/ML/DV and +50 nL on the plunger), which establish the
   initial movement-direction history. After they finish, use the jog buttons or
   normal keyboard shortcuts; later moves and reversals update direction automatically.
   Pressing **D** toggles the drill; this shortcut can be changed in Options → Keyboard Controls.
   Repeated manual syringe pulses are queued in order. Stop/Escape cancels the remaining
   pulses and stops the pulse in progress, then refreshes the verified plunger position.
   Repeated movement-key presses run as a queue; consecutive pending steps in the
   same direction on an axis are combined into one longer move. Pressing the
   opposite direction replaces pending moves on that axis while preserving other axes' queue. Escape
   cancels all queued steps and stops the active move. The controller verifies the stopped counts, saves any
   partial movement/backlash offset, and stays connected when that state is stable.
   Use the physical Stop if software Stop or count verification fails.
   A physical scale field unlocks after that axis has moved at least 0.200 mm
   (50 nL for the plunger). Enter all four physical readings,
   then press **Save Calibration**. This stores the live calibration profile in
   `DirectUSB\live\direct-control.json`, then reconnects the planner to hardware.
   Each completed move stores the new position and backlash direction in
   `DirectUSB\live\api-state.json`; an interrupted move is reconciled from stable
   stopped counts when possible, including partial backlash take-up.
   The calibration screen never substitutes simulator data for hardware.
3. Confirm physical reference and clearance. Planner default is captured 1 mm/s;
   2 mm/s is optional. Start with 0.01 mm AP, with physical observation.
4. Connect / restore. Investigate mismatches; do not reset references to evade
   faults. Zero calibration changes interpretation, not device position or homing.

GUI Bregma is a separate tool reference: Set Bregma, then Set Anchor if needed.
After tool exchange touch the anchor and confirm At Anchor. Requested Bregma
positions are converted to calibrated Axis targets; no native Bregma action occurs.

## Features and limits

Keyboard movements/shortcuts, absolute GoTo, named coordinates, Home/Work/A/B,
map moves, Bregma/Anchor, site validation, overlay maps, pan/zoom, grid/site editing,
import/export and project autosave remain available through the direct adapter.
Manual piston steps are free-piston requests, **not controlled-rate injections**.
Reported piston volume is an estimate for the tested Nano 5 µL syringe, not delivery.

**Options → Direct control — speed and travel limits** provides independent minimum
and maximum ranges for AP, ML and DV in mechanical Axis mm (defaults **0–40 mm**)
and piston in nL (default **500–4500 nL**). Piston travel is hard-limited to this
safe operating band; the syringe's nominal capacity is 0–5000 nL. These are Axis coordinates even in Bregma
mode. Click **Apply and save speed / limits** while idle; settings persist in the
mode-specific `direct-control.json` and also apply to an already connected controller.
Axis speed selects the captured **1 or 2 mm/s** profile, separate from keyboard step
size and pulsed average rates. The previous ±1 mm/±100 nL connection envelopes and
1 mm combined-distance cap are removed. Complete approach paths and workflow doses
are preflighted. Limits are not collision protection; verify the full path independently.
Piston free steps remain 10/20/50/100 nL; configured limits cannot exceed 500–4500 nL
for the Nano 5 µL syringe. Empty/Fill move only to the configured lower/upper limit.
Signed motor-count overflow remains rejected.
Craniotomy **Drilling pattern and timing → Mode** defaults to **Spaced burr holes**;
select **Continuous path** for the existing perimeter-tracing behavior. Burr holes are placed
uniformly along the closed perimeter at no more than the selected center-to-center
spacing. Set every craniotomy seed surface first; the surface at each burr hole is
inferred from the interpolated seed-surface profile, so individual hole surfaces do
not need separate capture. Each advances by **Depth increment / round** until **Max Depth**, retracting
to clearance before moving to the next hole. Hole progress is retained only while
the sampled surface plan and spacing still match. Spaced burr-hole drilling requires
the verified pulsed controller workflow. **Time per circuit** paces visits around the
perimeter; serial movement and settling overhead can make actual time longer.
In this mode, **Freeze Holes** and **Unfreeze Holes** let you draw over individual
burr-hole markers on the map; frozen holes are gray and are skipped on deeper rounds.
Increase **Max Depth** to continue deepening the remaining unfrozen holes; the next
depth increment becomes available when the previous maximum had already been reached.
The same draw controls freeze/unfreeze perimeter sections in Continuous path mode.
Pulsed drilling includes frozen sections, pause/retract/Continue and round progression.
Pulsed injection includes Start/Resume, insertion and main doses, overshoot,
post-injection hold, Pause/Resume, surface return and blockage tests.
The craniotomy tab lists seed points; select one from the list
or click it on the map, then choose **Set Surface**. The planner moves to the
planned point above its surface and opens a shortcut-enabled dialog. Lower the
tool to touch the skull and choose **At Surface** to save its current GUI Bregma
position. Captured point surfaces are autosaved and become the surface targets
used by drilling. **Drill On/Off** toggles the reported drill state directly through
the API without a confirmation popup; the drilling sequence still separately asks
you to confirm the drill is on and verifies reported power. Reported power is not
proof of spindle rotation or rest.
Manual/test volumes up to 2000 nL are restored: larger requests expand into
captured 10/20/50/100 nL steps, with whole-dose preflight and per-step verified
counters. Keyboard movement step choices again include 2 and 5 mm, within travel limits.
**Empty Syringe / Fill Syringe** move toward the configured lower/upper piston
limit with a confirmation and cancellable progress dialog. Remove the pipette from
the specimen and verify safe collection/aspiration first. If the endpoint is not
reachable in 10 nL units, they stop short by less than 10 nL and show the real
verified piston estimate; they never pretend the syringe is exactly empty/full.
**Go To** in Manual Control moves to a chosen 500–4500 nL position in 10 nL
increments, using the calibrated piston API and cancellable progress when in direct mode.
The configurable blockage-test volume defaults to **20 nL**.
**Options → Benchmark Axis Moves** selects axes, distances (<=1 mm) and repeats;
AP/ML are selected by default, DV is off. The entire out-and-back path is preflighted,
drill power must be OFF, and Cancel/Esc stops without automatic return. Copyable CSV
reports count-derived displacement and elapsed command time, not encoder feedback.
Continuous firmware speed/flow, automatic hardware homing and native Auto-Speed/
safety zones are not established by the captured protocol; no guessed packets are sent.
Ordinary Options/confirmation dialogs suppress motor shortcuts; the dedicated
site-validation modal still accepts the requested movement and step-size shortcuts.
**Update** now works on this branch: confirm while idle, USB disconnects, the direct
branch is fetched, local tracked/non-ignored untracked edits are saved in a Git stash,
and the checkout is set to the verified remote commit. It refuses another branch
or a checkout changed during fetch/backup, or changes Git cannot stash. Do not run
other Git operations during updating. Ignored files and external settings are preserved.
The progress dialog remains responsive; Cancel waits for the current Git command.
It offers **Restart now** and retains live/simulation mode without auto-connecting.
If you defer restarting (or cancel/fail an update), reconnecting is blocked until
restart to prevent mixing old in-memory code with changed source. Restore a backup
on a closed app with `git stash list` and `git stash apply <saved-stash-hash>`.
Begin with small supervised bench plans. Larger site grids are allowed only if
their entire path and dose fit the configured travel ranges and available volume.
The dedicated network probe retains its independent experiment radius/command limits.

Connecting conservatively clears tool references, captured surfaces and quick
targets for reverification. Home/Work persist with identical Axis calibration,
but are cleared when it changes or is unknown; verify clearance before using them.
Keyboard shortcuts and application preferences are stored under
`%APPDATA%\Neurostar_Master\settings.json`. Live and simulation control
profiles/state are separated under `%APPDATA%\Neurostar_Master\DirectUSB`,
outside the repository and preserved across repository updates. On first launch,
existing settings and control files are copied from the former
`Documents\Neurostar_Master\Configs` location only when the AppData copy is
absent; the old files are left in place.
Each mode has a dedicated `direct-control.json` for measured calibration, captured
speed profile, travel ranges, Axis-zero fingerprint and Home/Work. Supported
capabilities are enabled by default and have no enable checkboxes.
These are separate from general `settings.json` and the live API motion journal
`api-state.json`. Setup preferences save immediately; existing Home/Work metadata
migrates automatically. Simulation startup auto-connects only the volatile
in-memory controller; its current AP/ML/DV and piston positions are never saved.
Live startup attempts a restore only when saved calibration and matching API
state are available; otherwise it falls back to simulation with a warning.

Stop/Esc/close cancel movement, stop all four channels and attempt drill OFF even
without ON permission. Software Stop depends on the connection/process; use physical
Stop for uncertainty. Faults invalidate state; no automatic retry/recovery reversal.

## Documentation and tests

- [Direct operation, calibration contract and full integration review](docs/DIRECT_API_CONTROL.md)
- [Same-computer HTTP server and agent instructions](tools/movement_probe/README.md)
- [Shared API and standalone GUI](stereodrive_api/README.md): `py -3 -m stereodrive_api.gui`
- [Protocol investigation/evidence](stereodrive_api/TECHNICAL_REPORT.md)
- [Archived native-GUI instructions](LEGACY_GUI_README.md), for the old branch only

```powershell
py -3 -B -m unittest discover -s tests -v
py -3 -B -m unittest discover -s stereodrive_api/tests -v
```

PySide6 enables actual Qt tests; Tkinter enables the API GUI tests (standard Windows
Python installer). Passing simulation cannot establish surgical safety. Validate
physical displacement, delivered volume, reversal behavior and Stop independently
on a supervised bench before practical hardware use.
