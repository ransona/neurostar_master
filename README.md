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

This opens simulation, disconnected. In **USB setup / zero calibration**, the
simulation fixture allows testing without hardware; it is not live calibration.

For hardware mode (also initially disconnected):

```powershell
py -3 tools\craniotomy_qt.py --live
```

`run_craniotomy_gui.bat` launches from its own folder and accepts `--live`.
Update does not reset this branch to `main`; close the app and use `git pull --ff-only`.

## Before enabling hardware

1. Clear the bench: no specimen, tool/pipette retracted, drill off, physical Stop
   accessible. Close StereoDrive, all other controllers and the probe server.
2. Load measured calibration JSON: AP/ML/DV raw counts at mechanical Axis zero,
   piston raw count at 3000 nL, and backlash states at those anchors. Separately
   select verified current direction/backlash history of all four channels.
   Never copy simulation counts or guess history from a raw count.
3. Confirm physical reference and clearance. Enable DV, piston or drill ON only
   after independently checking their setup. Planner default is captured 1 mm/s;
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

Axis limits: 1 mm per command, ±1 mm per axis from connection, and 1 mm combined
distance per absolute multi-axis request. Complete approach paths are preflighted.
Piston limits: 10/20/50/100 nL, ±100 nL from connection, estimated 0–5000 nL capacity.
These limits do not prove clearance. Reconnecting to expand travel is not a workaround.
Pulsed drilling includes frozen sections, pause/retract/Continue and round progression.
Pulsed injection includes Start/Resume, insertion and main doses, overshoot,
post-injection hold, Pause/Resume, surface return and blockage tests.
Continuous firmware rate control, empty/fill and native benchmarks remain unsupported.
Use small bench plans fitting existing envelopes: larger default/site-grid
protocols may be rejected in full before movement. No limits were widened.

Connecting conservatively clears tool references, captured surfaces and quick
targets for reverification. Home/Work persist with identical Axis calibration,
but are cleared when it changes or is unknown; verify clearance before using them.
Live and simulation settings/state are separated
under `Documents\Neurostar_Master\Configs\DirectUSB`, outside the repository.
Each mode has a dedicated `direct-control.json` for measured calibration, captured
speed profile, DV/piston/drill preferences, Axis-zero fingerprint and Home/Work.
The `allow_pulsed` preference is saved there too; it is off by default.
These are separate from general `settings.json` and the API motion journal
`api-state.json`. Setup preferences save immediately; existing Home/Work metadata
migrates automatically. Startup stays disconnected, with safety confirmation and
current direction history requiring fresh verification each connection.

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
