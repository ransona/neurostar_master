# Neurostar StereoDrive Craniotomy Planner

This repository contains a Windows Qt application for driving **StereoDrive** and the **Injectomate** during craniotomy drilling and injection workflows.

The main GUI is:

- [tools/craniotomy_qt.py](C:/code/repos/neurostar_master/tools/craniotomy_qt.py)

The main Win32 automation/controller layer is:

- [tools/stereodrive_controller.py](C:/code/repos/neurostar_master/tools/stereodrive_controller.py)

PowerShell diagnostics and direct control tests live in:

- [tools/control_stereodrive.ps1](C:/code/repos/neurostar_master/tools/control_stereodrive.ps1)

## What The Software Does

The application sits on top of the native StereoDrive application and automates:

- AP / ML / DV movement
- Bregma reset / synchronization setup
- Seed generation for circular craniotomies
- Repeated drilling rounds around a perimeter
- Freeze / unfreeze marking of perimeter points
- Injection-site storage
- Automated injection sequences
- Injectomate syringe stepping and plunger-position tracking
- Injectomate calibrate-popup reading for real syringe position
- Benchmarking of axis movement speed

It does **not** replace StereoDrive. StereoDrive must already be running.

## Requirements

- Windows
- Python 3
- `PySide6`
- StereoDrive running and detectable by its main window class/title
- Hardware connected and functioning in StereoDrive

## How To Run

From the repo root:

```powershell
python .\tools\craniotomy_qt.py
```

If StereoDrive is not running, the app will fail to attach.

## High-Level UI Layout

The window title is **Craniotomy Planner**.

Main areas:

- Top header for coordinate actions and quick stored locations
- **Setup** panel for craniotomy planning
- Two projection views for trajectory / seed visualization
- **Manual Control** panel for syringe manual actions and syringe-position display
- **Injection** panel for automated injection protocol settings and sequence progress
- **Injection Sites** panel for storing and resuming site lists

## Startup Behavior

On startup the app:

- Connects to StereoDrive
- Starts a live position refresh timer
- Starts a background watcher that auto-confirms some StereoDrive dialogs:
  - below-skull warning: clicks **Yes**
  - "no actual movement to execute": clicks **OK**
- Schedules an early syringe-position update from the Injectomate calibrate popup

## Coordinate Controls

Top-row actions include:

- `Set Bregma`
- `Bregma`
- `Home`
- `Work`
- `Go to`
- `Stop`
- quick stored locations `A`, `B`, `C` and corresponding `Set A`, `Set B`, `Set C`
- `Benchmark`

### Set Bregma

`Set Bregma` is intended to zero the current position to Bregma through the StereoDrive synchronize/reference workflow.

Internally it uses the direct StereoDrive command that opens the **Synchronize Drill and Syringe...** panel instead of navigating the Tools menu manually.

### Bregma / Home / Work / Go to

- `Bregma` moves to the Bregma location
- `Home` moves to StereoDrive home
- `Work` moves to StereoDrive work
- `Go to` opens a coordinate dialog populated with current AP / ML / DV and allows entry of a new destination

### Quick Stored Locations

You can store three extra locations:

- `Set A` / `A`
- `Set B` / `B`
- `Set C` / `C`

These are intended for frequently revisited points.

## Setup Panel

The **Setup** panel contains the craniotomy planning controls.

Important fields and actions include:

- craniotomy diameter
- seed count
- skull thickness
- max depth
- depth per round
- round time
- drill rate
- auto-start next round
- `Generate Seeds`
- `Next seed`
- `Set Surface`
- `Stop Motion`
- `Clear Surface Measurements`
- `Start Drilling`
- freeze/unfreeze controls

### Generate Seeds

This creates a circular perimeter of drilling targets.

Behavior:

- the perimeter is drawn immediately
- the craniotomy circle is frozen visually once seeds are generated
- generating seeds resets current drilling progress state
- the app can ask whether to move to the first seed

### Set Surface

At each seed, `Set Surface` stores the surface DV for that point.

Special shortcut:

- `Ctrl + click Set Surface` sets all surface values to `0`

This is mainly for fast debugging.

### Seed Navigation

`Next seed` moves through the seed list.

## Craniotomy Drilling Workflow

The drilling workflow is round-based.

For a round:

1. A current target depth is chosen.
2. The app traverses the circular perimeter.
3. At each point it moves to surface + current round depth.
4. Frozen points are skipped.
5. At the end of the round it returns above the center.

After a round:

- the app beeps
- a countdown popup appears
- it can auto-start the next round after the countdown
- the next target depth can be edited in that popup
- if auto-start is enabled, it proceeds automatically
- otherwise it pauses

### Start / Pause / Continue Drilling

`Start Drilling` becomes `Pause` while a round is running.

If paused:

- the drill is raised to `2 mm` above surface
- the button changes to `Continue`
- continuing returns to the first not-yet-drilled perimeter location for the current round depth

### Current Target Depth

There is a separate current target depth control under the drilling progress area.

- `Change` opens a dialog to set it directly
- this can be set deeper than skull thickness if needed

## Craniotomy Path Planning And Timing

Perimeter tracing uses a continuous AP/ML nudge path rather than simple point-to-point stepping.

Current behavior:

- AP/ML path stepping uses a DDA-style controller in `stereodrive_controller.py`
- the GUI estimates a `path_step_mm` for the round based on expected travel and round-time budget
- larger required travel or shorter round time results in larger step size and looser tolerance
- it prioritizes approximate path timing while still following the circle

Important limitation:

- `Round Time` is read when a round starts
- changing the round time during an already-running round does **not** currently re-plan that round
- the new value applies to the next round

## Manual Control Panel

The **Manual Control** panel contains syringe-related controls and readouts.

Main controls:

- `Step Syringe Up (F3)`
- `Step Syringe Down (F4)`
- red `Stop`
- `Empty Syringe`
- `Update Syringe Position`
- `Test for Blockage`

### Syringe Position

The vertical syringe gauge on the right tracks syringe position in `nl`.

Behavior:

- the full displayed range is `0` to `5000 nl`
- internal syringe state is tracked during manual stepping and injection programs
- on startup and before injection sequences, the real position is refreshed using the Injectomate calibrate popup

### Update Syringe Position

`Update Syringe Position` reads the Injectomate calibrate popup and updates the GUI/internal syringe position.

This uses the precise value exposed in the calibrate popup rather than trying to infer the on-panel custom gauge.

### Empty Syringe

This sends the Injectomate plunger to `0`.

## Injection Panel

The **Injection** panel defines a full single-site injection protocol.

Current settings:

- `Main injection volume (nl)`
- `Injection rate (nl/min)`
- `Insertion injection rate (nl/min)`
- `Injection depth (mm)`
- `Insert/retract speed (um/sec)`
- `Overshoot (mm)`
- `Post inject pause (s)`
- `Test volume (nl)`

Below that:

- program-sequence list
- overall sequence progress
- current injection / movement progress
- `Go`
- `Pause`
- `Stop`

### Injection Sequence Definition

For each site, the current protocol is:

1. Move to `1 mm` above the stored surface
2. Move normally to the stored surface
3. Insert from surface to `depth + overshoot` at `Insert/retract speed`
4. Inject during insertion at `Insertion injection rate`
5. Retract overshoot back to target depth at the same insert/retract speed
6. Continue main injection at target depth if needed
7. Post-injection pause at target if configured
8. Retract to the stored surface at insert/retract speed
9. Move normally to `1 mm` above the stored surface
10. Optionally run a blockage test

### Important Injection Details

- insertion and retraction are handled by controlled DV motion
- the pipette advances downward while injection is occurring
- overshoot defaults to `0.05 mm`
- the post-injection pause is shown in the feedback/status area
- the active sequence step is bolded in the program-sequence list during execution
- the active injection site is bolded in the injection-site list during execution

## Injection Sites Panel

This panel manages stored injection sites.

Controls:

- `Add Injection Site`
- `Remove Selected Site`
- `Clear Sites`
- `Resume From Selected`
- `Check blockage after each site`

### Injection Site Storage

Each stored site contains:

- AP
- ML
- surface DV

If no sites are stored, the current location is treated as the active site and its current DV is treated as the surface.

## Starting An Injection Sequence

Press `Go` in the Injection panel.

Before starting, the app:

- reads current GUI injection settings
- syncs real syringe position from the calibrate popup
- estimates required syringe volume including optional blockage tests
- refuses to start if the requested sequence would exceed syringe limits

## Stopping And Resuming Injection

### Stop

Pressing `Stop`:

- requests sequence stop
- attempts to stop active Injectomate motion
- refreshes real syringe position from the calibrate-popup read method

### Resume From Selected

After a stopped sequence:

1. Select a site in the `Injection Sites` list
2. Click `Resume From Selected`

The app will:

- restart the protocol from the selected site onward
- use the **current GUI settings**
- re-check syringe position before restarting

This is not a low-level resume inside a partially executed site. It resumes at the selected stored site boundary.

## Blockage Test Behavior

If `Check blockage after each site` is enabled:

- after returning to `1 mm` above the stored surface, the app performs a test injection
- once the test injection completes, it asks:
  - `Is the test injection confirmed not blocked?`

If you choose:

- `Not blocked`: continue to the next site
- `Blocked`: the app asks whether to do another test injection
- `Another test`: it repeats the blockage test at the same site
- `No`: it stops the sequence there

This repeats until:

- you confirm clear, or
- you decline further tests, or
- you stop the sequence manually

## Syringe Limits And Safety Checks

The GUI tracks syringe position in `nl` and enforces bounds.

Current allowed range:

- minimum: `0 nl`
- maximum: `5000 nl`

If a requested syringe movement or full sequence would exceed those bounds:

- the app shows a warning
- the action is rejected
- for sequence planning, the required total volume is checked in advance

## Automatic StereoDrive Dialog Handling

The app runs a background watcher that auto-confirms common dialogs:

- below-skull warning: clicks `Yes`
- no-actual-movement dialog: clicks `OK`

This helps keep automated drilling and movement from hanging on routine dialogs.

## Benchmarking

The `Benchmark` button runs a movement benchmark and shows the results in a popup text box suitable for copy/paste.

This is intended to measure actual movement speed for:

- AP
- ML
- DV

across several step sizes.

The benchmark is useful for tuning round-time expectations and continuous path behavior.

## Notes On Injectomate Control

The code can:

- show the Injectomate panel directly
- read the calibrate-popup plunger value
- trigger syringe steps up/down
- stop active Injectomate motion

The most reliable precise syringe-position read path is currently the **Injectomate calibrate popup**.

## Troubleshooting

### The app says StereoDrive was not found

- Start StereoDrive first
- Make sure the StereoDrive main window is actually present

### The app opens but movement does not happen

- Confirm StereoDrive is connected to the hardware
- Confirm the selected nudge sizes are valid in StereoDrive
- Watch for any external modal dialogs not covered by the auto-confirm watcher

### Set Bregma does not work

- The app expects to be able to open the synchronize/reference panel through StereoDrive
- If StereoDrive is in an unusual UI state, reopen the main StereoDrive window and try again

### Injectomate position looks wrong

- Press `Update Syringe Position`
- This re-reads the calibrate popup and overwrites the internal estimate with the real plunger value

### The sequence stops after a blockage test

- This is expected if you answered that the pipette is blocked and declined another test injection
- Select the next desired site and use `Resume From Selected` if needed

## Main Files

- GUI: [tools/craniotomy_qt.py](C:/code/repos/neurostar_master/tools/craniotomy_qt.py)
- Controller: [tools/stereodrive_controller.py](C:/code/repos/neurostar_master/tools/stereodrive_controller.py)
- PowerShell diagnostics: [tools/control_stereodrive.ps1](C:/code/repos/neurostar_master/tools/control_stereodrive.ps1)

## Typical Usage

A common session looks like this:

1. Start StereoDrive
2. Run `python .\tools\craniotomy_qt.py`
3. Press `Set Bregma` if needed
4. Set craniotomy diameter and seed count
5. Press `Generate Seeds`
6. Move around the perimeter and capture surface values with `Set Surface`
7. Start drilling rounds with `Start Drilling`
8. Store injection sites with `Add Injection Site`
9. Configure the injection protocol
10. Press `Go`
11. Confirm blockage-test outcomes between sites if enabled

## Current Scope

This README describes the current behavior implemented in the codebase as of the present repository state. It is an operator guide, not a formal validation document.
